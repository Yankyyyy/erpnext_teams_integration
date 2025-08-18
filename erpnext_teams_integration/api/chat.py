import frappe, requests
from .helpers import get_access_token, get_azure_user_id_by_email, get_login_url
from frappe.utils import now_datetime, get_datetime
from datetime import datetime

GRAPH_API = 'https://graph.microsoft.com/v1.0'
my_azure = frappe.db.get_single_value('Teams Settings', 'owner_azure_object_id')

# Add all supported doctypes here, along with the participant child table + email field
SUPPORTED_DOCTYPES = {
    "Event": {"participants_field": "event_participants", "email_field": "email"},
    "Project": {"participants_field": "users", "email_field": "email"},
    # Add more doctypes as needed
}

@frappe.whitelist()
def create_group_chat_for_doc(docname, doctype):
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"{doctype} is not supported for Teams chat creation.")

    doc = frappe.get_doc(doctype, docname)
    token = get_access_token()
    if not token:
        return {'error': 'auth_required', 'login_url': get_login_url(docname)}

    participants_field = SUPPORTED_DOCTYPES[doctype]["participants_field"]
    email_field = SUPPORTED_DOCTYPES[doctype]["email_field"]

    # Get current Azure Object IDs from doc participants
    target_azure_ids = set()
    if getattr(doc, participants_field, None):
        for p in getattr(doc, participants_field):
            email_val = getattr(p, email_field, None)
            azure = None
            if email_val:
                azure = frappe.db.get_value('User', email_val, 'azure_object_id')
                if not azure:
                    azure = get_azure_user_id_by_email(email_val)
            if azure:
                target_azure_ids.add(azure)

    if not target_azure_ids:
        frappe.throw('No valid Microsoft users found for chat.')

    # If chat already exists → only add missing members
    if getattr(doc, "custom_teams_chat_id", None):
        chat_id = doc.custom_teams_chat_id

        # Get existing members from Teams
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        res = requests.get(f"{GRAPH_API}/chats/{chat_id}/members", headers=headers)
        if res.status_code != 200:
            frappe.throw(f"Failed to fetch chat members: {res.text}")

        existing_ids = {m["userId"] for m in res.json().get("value", []) if "userId" in m}
        new_members = target_azure_ids - existing_ids

        for azure_id in new_members:
            member_payload = {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{azure_id}')"
            }
            add_res = requests.post(
                f"{GRAPH_API}/chats/{chat_id}/members",
                headers=headers,
                json=member_payload
            )
            if add_res.status_code not in (200, 201):
                frappe.log_error(add_res.text, f"Failed to add member {azure_id} to {chat_id}")

        return {"chat_id": chat_id, "message": "Chat updated with new members (if any)."}

    # Otherwise create a new chat
    members = [{
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'roles': ['owner'],
        'user@odata.bind': f"https://graph.microsoft.com/v1.0/users('{azure_id}')"
    } for azure_id in target_azure_ids]
    
    if my_azure and my_azure not in target_azure_ids:
        members.append({
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            'roles': ['owner'],
            'user@odata.bind': f"https://graph.microsoft.com/v1.0/users('{my_azure}')"
        })
        target_azure_ids.add(my_azure)

    payload = {
        'chatType': 'group',  # 'group' or 'oneOnOne'
        # 'topic': f"{doctype} {docname}",
        'members': members
    }

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    res = requests.post(f"{GRAPH_API}/chats", headers=headers, json=payload)
    if res.status_code in (200, 201):
        chat_id = res.json().get('id')
        if frappe.db.has_column(doctype, 'custom_teams_chat_id'):
            frappe.db.set_value(doctype, docname, 'custom_teams_chat_id', chat_id)
        doc_data = {
            "doctype": "Teams Conversation",
            "chat_id": chat_id,
            "document_type": doctype,
            "document_name": docname,
            "last_synced": now_datetime()
        }

        if payload.get("topic"):  # only add if it's set and not empty
            doc_data["topic"] = payload["topic"]

        frappe.get_doc(doc_data).insert(ignore_permissions=True)
        frappe.db.commit()
        return {"chat_id": chat_id, "message": "New chat created."}

    frappe.log_error(res.text, 'Teams create chat failed')
    frappe.throw(f"Failed to create chat: {res.status_code}: {res.text}")

        
@frappe.whitelist()
def send_message_to_chat(chat_id, message, docname=None, doctype=None):
    token = get_access_token()
    if not token:
        return {'error':'auth_required'}
    headers = {'Authorization':f'Bearer {token}','Content-Type':'application/json'}
    payload = {'body':{'contentType':'html','content': message}}
    res = requests.post(f"{GRAPH_API}/chats/{chat_id}/messages", headers=headers, json=payload)
    if res.status_code in (200,201):
        msg = res.json()
        _save_message_local(msg, chat_id, docname, doctype, 'Outbound')
        return msg
    elif res.status_code == 401:
        from .helpers import refresh_access_token
        refresh_access_token()
        token = get_access_token()
        headers['Authorization'] = f'Bearer {token}'
        res = requests.post(f"{GRAPH_API}/chats/{chat_id}/messages", headers=headers, json=payload)
        if res.status_code in (200,201):
            msg = res.json(); _save_message_local(msg, chat_id, docname, doctype, 'Outbound'); return msg
    frappe.log_error(res.text, 'Teams send message failed'); frappe.throw(f"Teams API Error: {res.status_code}: {res.text}")

@frappe.whitelist()
def get_local_chat_messages(chat_id, limit=200):
    msgs = frappe.get_all('Teams Chat Message', filters={'chat_id': chat_id}, fields=['message_id','sender_display','body','created_at','direction'], order_by='created_at asc', limit_page_length=limit)
    for m in msgs:
        if m.get('created_at'):
            m['created_at'] = str(m['created_at'])
    return msgs

def fetch_and_store_chat_messages(chat_id, event_docname=None, top=50):
    token = get_access_token()
    if not token:
        return None
    headers = {'Authorization':f'Bearer {token}'}
    res = requests.get(f"{GRAPH_API}/chats/{chat_id}/messages?$top={top}", headers=headers)
    if res.status_code == 200:
        items = res.json().get('value', [])
        for item in items:
            _save_message_local(item, chat_id, event_docname, 'Inbound')
        return items
    elif res.status_code == 401:
        from .helpers import refresh_access_token
        refresh_access_token(); token = get_access_token(); headers['Authorization'] = f'Bearer {token}'
        res = requests.get(f"{GRAPH_API}/chats/{chat_id}/messages?$top={top}", headers=headers)
        if res.status_code == 200:
            items = res.json().get('value', [])
            for item in items:
                _save_message_local(item, chat_id, event_docname, 'Inbound')
            return items
    frappe.log_error(res.text, 'Teams fetch messages failed'); return None

def _save_message_local(msg_json, chat_id, docname=None, doctype=None, direction='Inbound'):
    try:
        mid = msg_json.get('id')
        if frappe.db.exists('Teams Chat Message', {'message_id': mid}):
            return  # Avoid duplicates

        body = msg_json.get('body', {}).get('content')

        # Ensure created_at is in Frappe's datetime format
        created_str = msg_json.get('createdDateTime') or now_datetime()
        if created_str:
            try:
                # Replace 'Z' with '+00:00' for proper ISO parsing
                created_dt = datetime.fromisoformat(created_str.replace('Z', '+00:00'))
                # Format without timezone for MySQL DATETIME
                created_at = created_dt.strftime('%Y-%m-%d %H:%M:%S')
                # If DATETIME(3) in DB, keep milliseconds:
                # created_at = created_dt.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
            except Exception:
                created_at = now_datetime().strftime('%Y-%m-%d %H:%M:%S')
        else:
            created_at = now_datetime().strftime('%Y-%m-%d %H:%M:%S')

        sender = msg_json.get('from', {}).get('user') or msg_json.get('from')
        sender_id = None
        sender_display = None

        if sender and isinstance(sender, dict):
            sender_id = sender.get('id') or sender.get('user', {}).get('id')
            sender_display = sender.get('displayName') or sender.get('user', {}).get('displayName')

        # Build doc payload
        payload = {
            'doctype': 'Teams Chat Message',
            'chat_id': chat_id,
            'message_id': mid,
            'sender_id': sender_id,
            'sender_display': sender_display,
            'body': body,
            'created_at': created_at,
            'direction': direction
        }

        # Link to correct doctype
        if doctype == 'Event':
            payload['event'] = docname
        elif doctype == 'Project':
            payload['project'] = docname

        doc = frappe.get_doc(payload)
        doc.insert(ignore_permissions=True)
        frappe.db.commit()

    except Exception as e:
        frappe.log_error(f"Error saving Teams message: {str(e)}", "Teams Chat Message Save Failed")

@frappe.whitelist()
def post_message_to_channel(team_id, channel_id, message, docname=None):
    token = get_access_token()
    if not token:
        return {'error':'auth_required'}
    headers = {'Authorization':f'Bearer {token}','Content-Type':'application/json'}
    payload = {'body':{'contentType':'html','content':message}}
    res = requests.post(f"{GRAPH_API}/teams/{team_id}/channels/{channel_id}/messages", headers=headers, json=payload)
    if res.status_code in (200,201):
        return res.json()
    elif res.status_code == 401:
        from .helpers import refresh_access_token
        refresh_access_token()
        token = get_access_token()
        headers['Authorization'] = f'Bearer {token}'
        res = requests.post(f"{GRAPH_API}/teams/{team_id}/channels/{channel_id}/messages", headers=headers, json=payload)
        if res.status_code in (200,201):
            return res.json()
    frappe.log_error(res.text, 'Teams channel post failed')
    frappe.throw(f"Teams API Error: {res.status_code}: {res.text}")

@frappe.whitelist()
def sync_all_conversations(chat_id=None):
    try:
        access_token = get_access_token()
        if not access_token:
            frappe.throw("Could not fetch Teams access token. Please log in again.")

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        base_url = f"https://graph.microsoft.com/v1.0/chats"
        chats = []

        # If chat_id provided → only sync that one, else sync all
        if chat_id:
            chats = [chat_id]
        else:
            resp = requests.get(base_url, headers=headers)
            resp.raise_for_status()
            data = resp.json()
            chats = [c.get("id") for c in data.get("value", []) if c.get("id")]

        for cid in chats:
            messages_url = f"{base_url}/{cid}/messages"
            resp = requests.get(messages_url, headers=headers)
            resp.raise_for_status()
            messages = resp.json().get("value", [])

            for msg in messages:
                _save_message_local(
                    msg_json=msg,
                    chat_id=cid,
                    docname=None,   # Could link to Event/Project if known
                    doctype=None,   # Change if you want direct linking
                    direction="Inbound"
                )

        frappe.msgprint(f"Synced {len(chats)} conversation(s) successfully.")

    except Exception as e:
        frappe.log_error(f"Error syncing conversations: {str(e)}", "Teams Sync Conversations Failed")
        frappe.throw("Failed to sync Teams conversations. Check error logs for details.")
