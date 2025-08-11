import frappe, requests
from .helpers import get_access_token, get_azure_user_id_by_email, get_login_url
from frappe.utils import now_datetime
from datetime import datetime

GRAPH_API = 'https://graph.microsoft.com/v1.0'

@frappe.whitelist()
def create_group_chat_for_doc(docname, doctype):
    doc = frappe.get_doc(doctype, docname)
    token = get_access_token()
    if not token:
        return {'error': 'auth_required', 'login_url': get_login_url(docname)}

    members = []
    added_users = set()  # track azure IDs to avoid duplicates

    my_azure = frappe.db.get_value('User', frappe.session.user, 'azure_object_id')
    if my_azure and my_azure not in added_users:
        members.append({
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            'roles': ['owner'],
            'user@odata.bind': f"https://graph.microsoft.com/v1.0/users('{my_azure}')"
        })
        added_users.add(my_azure)

    if doctype == 'Event' and getattr(doc, 'event_participants', None):
        for p in doc.event_participants:
            azure = None
            if p.email:
                azure = frappe.db.get_value('User', p.email, 'azure_object_id')
            if not azure and p.email:
                azure = get_azure_user_id_by_email(p.email)

            if azure and azure not in added_users:
                members.append({
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    'roles': ['owner'],
                    'user@odata.bind': f"https://graph.microsoft.com/v1.0/users('{azure}')"
                })
                added_users.add(azure)

    if not members:
        frappe.throw('No valid Microsoft users to create chat.')

    payload = {
        'chatType': 'group',
        'topic': f"{doctype} {docname}",
        'members': members
    }

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    res = requests.post(f"{GRAPH_API}/chats", headers=headers, json=payload)

    if res.status_code in (200, 201):
        chat = res.json()
        chat_id = chat.get('id')

        try:
            if frappe.db.has_column(doctype, 'custom_teams_chat_id'):
                frappe.db.set_value(doctype, docname, 'custom_teams_chat_id', chat_id)
        except Exception:
            pass

        conv = frappe.get_doc({
            'doctype': 'Teams Conversation',
            'chat_id': chat_id,
            'event': (docname if doctype == 'Event' else None),
            'topic': payload['topic'],
            'last_synced': now_datetime()
        })
        conv.insert(ignore_permissions=True)
        frappe.db.commit()

        return {'chat_id': chat_id}
    else:
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
    mid = msg_json.get('id')
    if frappe.db.exists('Teams Chat Message', {'message_id': mid}):
        return

    body = msg_json.get('body', {}).get('content')

    # Get and format created_at
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

    doc = frappe.get_doc({
        'doctype': 'Teams Chat Message',
        'chat_id': chat_id,
        'message_id': mid,
        'sender_id': sender_id,
        'sender_display': sender_display,
        'body': body,
        'created_at': created_at,
        'event': (docname if doctype == 'Event' else None),
        'direction': direction
    })
    doc.insert(ignore_permissions=True)
    frappe.db.commit()

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
def sync_all_conversations():
    convs = frappe.get_all('Teams Conversation', fields=['name','chat_id','event'])
    for c in convs:
        try:
            fetch_and_store_chat_messages(c.chat_id, c.event, top=50)
            return "Done"
        except Exception as e:
            frappe.log_error(str(e), f'Teams Sync {c.chat_id}')
