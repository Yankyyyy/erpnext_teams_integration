import frappe
import requests
from .helpers import get_access_token, get_azure_user_id_by_email, get_login_url
from frappe.utils import now_datetime, get_datetime, sanitize_html
from datetime import datetime
import json
import html

GRAPH_API = 'https://graph.microsoft.com/v1.0'

# Supported doctypes with their configuration
SUPPORTED_DOCTYPES = {
    "Event": {
        "participants_field": "event_participants", 
        "email_field": "email",
        "subject_field": "subject"
    },
    "Project": {
        "participants_field": "users", 
        "email_field": "email",
        "subject_field": "project_name"
    },
    # Add more doctypes as needed
}


def get_my_azure_id():
    """Get the current user's Azure ID safely"""
    try:
        settings = frappe.get_single('Teams Settings')
        if settings.owner_azure_object_id:
            return settings.owner_azure_object_id
        
        # Try to get from current user
        current_user = frappe.session.user
        if current_user != "Guest":
            azure_id = frappe.db.get_value('User', current_user, 'azure_object_id')
            if azure_id:
                return azure_id
        
        return None
    except Exception as e:
        frappe.log_error(f"Failed to get current user's Azure ID: {str(e)}", "Teams Azure ID Error")
        return None


@frappe.whitelist()
def create_group_chat_for_doc(docname, doctype):
    """Create or update Teams group chat for a document"""
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"{doctype} is not supported for Teams chat creation.")

    try:
        doc = frappe.get_doc(doctype, docname)
        token = get_access_token()
        
        if not token:
            return {'error': 'auth_required', 'login_url': get_login_url(docname)}

        # Get participant configuration
        config = SUPPORTED_DOCTYPES[doctype]
        participants_field = config["participants_field"]
        email_field = config["email_field"]

        # Collect Azure Object IDs from participants
        target_azure_ids = set()
        participants_data = getattr(doc, participants_field, None)
        
        if participants_data:
            for participant in participants_data:
                email_val = getattr(participant, email_field, None)
                if email_val:
                    azure_id = get_azure_user_id_by_email(email_val)
                    if azure_id:
                        target_azure_ids.add(azure_id)

        # Add current user to chat
        my_azure_id = get_my_azure_id()
        if my_azure_id:
            target_azure_ids.add(my_azure_id)

        if not target_azure_ids:
            frappe.throw('No valid Microsoft Teams users found for chat creation.')

        # Check if chat already exists
        existing_chat_id = getattr(doc, "custom_teams_chat_id", None)
        
        if existing_chat_id:
            return update_existing_chat(existing_chat_id, target_azure_ids, token)
        else:
            return create_new_chat(docname, doctype, target_azure_ids, token)

    except Exception as e:
        frappe.log_error(f"Error creating chat for {doctype} {docname}: {str(e)}", "Teams Chat Creation Error")
        frappe.throw(f"Failed to create Teams chat: {str(e)}")


def update_existing_chat(chat_id, target_azure_ids, token):
    """Update existing chat with new members"""
    try:
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        # Get existing members
        response = requests.get(f"{GRAPH_API}/chats/{chat_id}/members", headers=headers, timeout=30)
        
        if response.status_code != 200:
            frappe.log_error(f"Failed to fetch chat members: {response.text}", "Teams API Error")
            frappe.throw(f"Failed to fetch existing chat members: {response.status_code}")

        existing_members = response.json().get("value", [])
        existing_ids = {member.get("userId") for member in existing_members if member.get("userId")}
        
        # Find new members to add
        new_member_ids = target_azure_ids - existing_ids
        
        if not new_member_ids:
            return {"chat_id": chat_id, "message": "Chat is up to date with all participants."}

        # Add new members
        added_count = 0
        for azure_id in new_member_ids:
            member_payload = {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{azure_id}')"
            }
            
            add_response = requests.post(
                f"{GRAPH_API}/chats/{chat_id}/members",
                headers=headers,
                json=member_payload,
                timeout=30
            )
            
            if add_response.status_code in (200, 201):
                added_count += 1
            else:
                frappe.log_error(
                    f"Failed to add member {azure_id} to chat {chat_id}: {add_response.text}",
                    "Teams Add Member Error"
                )

        return {
            "chat_id": chat_id, 
            "message": f"Added {added_count} new member(s) to existing chat."
        }

    except Exception as e:
        frappe.log_error(f"Error updating existing chat {chat_id}: {str(e)}", "Teams Chat Update Error")
        frappe.throw(f"Failed to update existing chat: {str(e)}")


def create_new_chat(docname, doctype, target_azure_ids, token):
    """Create a new Teams group chat"""
    try:
        # Build members list
        members = []
        for azure_id in target_azure_ids:
            members.append({
                '@odata.type': '#microsoft.graph.aadUserConversationMember',
                'roles': ['owner'],
                'user@odata.bind': f"https://graph.microsoft.com/v1.0/users('{azure_id}')"
            })

        payload = {
            'chatType': 'group',
            'members': members
        }

        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        response = requests.post(f"{GRAPH_API}/chats", headers=headers, json=payload, timeout=30)
        
        if response.status_code not in (200, 201):
            frappe.log_error(f"Failed to create chat: {response.text}", "Teams Create Chat Error")
            frappe.throw(f"Failed to create Teams chat: {response.status_code}")

        chat_data = response.json()
        chat_id = chat_data.get('id')
        
        if not chat_id:
            frappe.throw("Failed to get chat ID from Teams response")

        # Update document with chat ID
        if frappe.db.has_column(doctype, 'custom_teams_chat_id'):
            frappe.db.set_value(doctype, docname, 'custom_teams_chat_id', chat_id)

        # Create Teams Conversation record
        conversation_doc = {
            "doctype": "Teams Conversation",
            "chat_id": chat_id,
            "document_type": doctype,
            "document_name": docname,
            "last_synced": now_datetime()
        }

        frappe.get_doc(conversation_doc).insert(ignore_permissions=True)
        frappe.db.commit()
        
        return {"chat_id": chat_id, "message": "New Teams chat created successfully."}

    except Exception as e:
        frappe.log_error(f"Error creating new chat for {doctype} {docname}: {str(e)}", "Teams New Chat Error")
        frappe.throw(f"Failed to create new Teams chat: {str(e)}")


@frappe.whitelist()
def send_message_to_chat(chat_id, message, docname=None, doctype=None):
    """Send message to Teams chat with proper error handling"""
    if not chat_id or not message:
        frappe.throw("Chat ID and message are required")
    
    try:
        token = get_access_token()
        if not token:
            return {'error': 'auth_required', 'message': 'Authentication required'}
        
        # Sanitize message content
        sanitized_message = html.escape(str(message))
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        payload = {
            'body': {
                'contentType': 'html',
                'content': sanitized_message
            }
        }
        
        response = requests.post(
            f"{GRAPH_API}/chats/{chat_id}/messages", 
            headers=headers, 
            json=payload,
            timeout=30
        )
        
        if response.status_code in (200, 201):
            message_data = response.json()
            _save_message_local(message_data, chat_id, docname, doctype, 'Outbound')
            return {
                "success": True,
                "message_id": message_data.get("id"),
                "message": "Message sent successfully"
            }
        
        elif response.status_code == 401:
            # Try token refresh
            try:
                from .helpers import refresh_access_token
                token = refresh_access_token()
                headers['Authorization'] = f'Bearer {token}'
                
                response = requests.post(
                    f"{GRAPH_API}/chats/{chat_id}/messages", 
                    headers=headers, 
                    json=payload,
                    timeout=30
                )
                
                if response.status_code in (200, 201):
                    message_data = response.json()
                    _save_message_local(message_data, chat_id, docname, doctype, 'Outbound')
                    return {
                        "success": True,
                        "message_id": message_data.get("id"),
                        "message": "Message sent successfully after token refresh"
                    }
            except Exception as refresh_error:
                frappe.log_error(f"Token refresh failed: {str(refresh_error)}", "Teams Token Refresh Error")
        
        # Log the error and return failure
        frappe.log_error(f"Failed to send message: {response.status_code} - {response.text}", "Teams Send Message Error")
        frappe.throw(f"Failed to send message to Teams: {response.status_code}")
        
    except requests.exceptions.Timeout:
        frappe.log_error("Send message request timed out", "Teams Message Timeout")
        frappe.throw("Message sending timed out. Please try again.")
    except Exception as e:
        frappe.log_error(f"Error sending message: {str(e)}", "Teams Send Message Error")
        frappe.throw(f"Failed to send message: {str(e)}")


@frappe.whitelist()
def get_local_chat_messages(chat_id, limit=200):
    """Get chat messages from local database with pagination"""
    if not chat_id:
        return []
    
    try:
        limit = min(int(limit), 500)  # Cap at 500 messages
        
        messages = frappe.get_all(
            'Teams Chat Message', 
            filters={'chat_id': chat_id}, 
            fields=[
                'message_id', 'sender_display', 'body', 'created_at', 
                'direction', 'sender_id'
            ], 
            order_by='created_at desc', 
            limit_page_length=limit
        )
        
        # Convert datetime objects to strings for JSON serialization
        for msg in messages:
            if msg.get('created_at'):
                msg['created_at'] = str(msg['created_at'])
            
            # Sanitize message body for display
            if msg.get('body'):
                msg['body'] = sanitize_html(msg['body'])
        
        # Reverse to show oldest first
        return list(reversed(messages))
        
    except Exception as e:
        frappe.log_error(f"Error fetching local messages for chat {chat_id}: {str(e)}", "Teams Local Messages Error")
        return []


@frappe.whitelist()
def fetch_and_store_chat_messages(chat_id, docname=None, doctype=None, top=50):
    """Fetch messages from Teams API and store locally"""
    if not chat_id:
        return None
    
    try:
        token = get_access_token()
        if not token:
            return None
        
        top = min(int(top), 100)  # Cap at 100 messages per request
        headers = {'Authorization': f'Bearer {token}'}
        
        response = requests.get(
            f"{GRAPH_API}/chats/{chat_id}/messages?$top={top}",
            headers=headers,
            timeout=30
        )
        
        if response.status_code == 200:
            messages = response.json().get('value', [])
            stored_count = 0
            
            for message in messages:
                if _save_message_local(message, chat_id, docname, doctype, 'Inbound'):
                    stored_count += 1
            
            return {
                "success": True,
                "fetched": len(messages),
                "stored": stored_count
            }
            
        elif response.status_code == 401:
            # Try token refresh
            try:
                from .helpers import refresh_access_token
                token = refresh_access_token()
                headers['Authorization'] = f'Bearer {token}'
                
                response = requests.get(
                    f"{GRAPH_API}/chats/{chat_id}/messages?$top={top}",
                    headers=headers,
                    timeout=30
                )
                
                if response.status_code == 200:
                    messages = response.json().get('value', [])
                    stored_count = 0
                    
                    for message in messages:
                        if _save_message_local(message, chat_id, docname, doctype, 'Inbound'):
                            stored_count += 1
                    
                    return {
                        "success": True,
                        "fetched": len(messages),
                        "stored": stored_count
                    }
            except Exception as refresh_error:
                frappe.log_error(f"Token refresh failed during message fetch: {str(refresh_error)}", "Teams Fetch Error")
        
        frappe.log_error(f"Failed to fetch messages: {response.status_code} - {response.text}", "Teams Fetch Messages Error")
        return None
        
    except Exception as e:
        frappe.log_error(f"Error fetching messages for chat {chat_id}: {str(e)}", "Teams Fetch Messages Error")
        return None


def _save_message_local(msg_json, chat_id, docname=None, doctype=None, direction='Inbound'):
    """Save Teams message to local database with better error handling"""
    try:
        if not msg_json or not isinstance(msg_json, dict):
            return False
        
        message_id = msg_json.get('id')
        if not message_id:
            return False
        
        # Check if message already exists
        if frappe.db.exists('Teams Chat Message', {'message_id': message_id}):
            return False  # Already stored
        
        # Extract message content
        body_data = msg_json.get('body', {})
        body_content = ""
        
        if isinstance(body_data, dict):
            body_content = body_data.get('content', '')
        elif isinstance(body_data, str):
            body_content = body_data
        
        # Parse created timestamp
        created_str = msg_json.get('createdDateTime')
        created_at = now_datetime()
        
        if created_str:
            try:
                # Handle ISO format with timezone
                if created_str.endswith('Z'):
                    created_str = created_str[:-1] + '+00:00'
                
                created_dt = datetime.fromisoformat(created_str)
                created_at = created_dt.strftime('%Y-%m-%d %H:%M:%S')
            except (ValueError, AttributeError):
                # Use current time if parsing fails
                created_at = now_datetime().strftime('%Y-%m-%d %H:%M:%S')
        
        # Extract sender information
        sender_info = msg_json.get('from', {})
        sender_id = None
        sender_display = "Unknown"
        
        if isinstance(sender_info, dict):
            user_info = sender_info.get('user', {})
            if user_info:
                sender_id = user_info.get('id')
                sender_display = user_info.get('displayName', 'Unknown')
            else:
                sender_id = sender_info.get('id')
                sender_display = sender_info.get('displayName', 'Unknown')
        
        # Build document payload
        doc_data = {
            'doctype': 'Teams Chat Message',
            'chat_id': chat_id,
            'message_id': message_id,
            'sender_id': sender_id,
            'sender_display': sender_display,
            'body': sanitize_html(body_content) if body_content else "",
            'created_at': created_at,
            'direction': direction
        }
        
        # Link to document if provided
        if doctype and docname:
            doc_data['document_type'] = doctype
            doc_data['document_name'] = docname
        
        # Create and save document
        message_doc = frappe.get_doc(doc_data)
        message_doc.insert(ignore_permissions=True)
        frappe.db.commit()
        
        return True
        
    except Exception as e:
        frappe.log_error(f"Error saving Teams message {message_id}: {str(e)}", "Teams Message Save Error")
        return False


@frappe.whitelist()
def post_message_to_channel(team_id, channel_id, message, docname=None):
    """Post message to Teams channel"""
    if not all([team_id, channel_id, message]):
        frappe.throw("Team ID, Channel ID, and message are required")
    
    try:
        token = get_access_token()
        if not token:
            return {'error': 'auth_required', 'message': 'Authentication required'}
        
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        payload = {
            'body': {
                'contentType': 'html',
                'content': html.escape(str(message))
            }
        }
        
        response = requests.post(
            f"{GRAPH_API}/teams/{team_id}/channels/{channel_id}/messages",
            headers=headers,
            json=payload,
            timeout=30
        )
        
        if response.status_code in (200, 201):
            return {
                "success": True,
                "message": "Posted to channel successfully"
            }
        elif response.status_code == 401:
            # Try token refresh
            try:
                from .helpers import refresh_access_token
                token = refresh_access_token()
                headers['Authorization'] = f'Bearer {token}'
                
                response = requests.post(
                    f"{GRAPH_API}/teams/{team_id}/channels/{channel_id}/messages",
                    headers=headers,
                    json=payload,
                    timeout=30
                )
                
                if response.status_code in (200, 201):
                    return {
                        "success": True,
                        "message": "Posted to channel successfully after token refresh"
                    }
            except Exception as refresh_error:
                frappe.log_error(f"Token refresh failed during channel post: {str(refresh_error)}", "Teams Channel Post Error")
        
        frappe.log_error(f"Failed to post to channel: {response.status_code} - {response.text}", "Teams Channel Post Error")
        frappe.throw(f"Failed to post to Teams channel: {response.status_code}")
        
    except Exception as e:
        frappe.log_error(f"Error posting to channel: {str(e)}", "Teams Channel Post Error")
        frappe.throw(f"Failed to post to channel: {str(e)}")


@frappe.whitelist()
def sync_all_conversations(chat_id=None):
    """Sync Teams conversations with better error handling and progress tracking"""
    try:
        access_token = get_access_token()
        if not access_token:
            frappe.throw("Could not fetch Teams access token. Please authenticate first.")

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        synced_count = 0
        error_count = 0

        if chat_id:
            # Sync specific chat
            result = _sync_single_chat(chat_id, headers)
            if result:
                synced_count = 1
            else:
                error_count = 1
        else:
            # Sync all chats
            try:
                # Get all chats from Teams
                chats_response = requests.get(f"{GRAPH_API}/chats", headers=headers, timeout=30)
                
                if chats_response.status_code == 200:
                    chats_data = chats_response.json()
                    chat_list = chats_data.get("value", [])
                    
                    for chat in chat_list:
                        chat_id = chat.get("id")
                        if chat_id:
                            if _sync_single_chat(chat_id, headers):
                                synced_count += 1
                            else:
                                error_count += 1
                else:
                    frappe.log_error(f"Failed to fetch chats list: {chats_response.status_code} - {chats_response.text}", "Teams Sync Error")
                    frappe.throw("Failed to fetch chats list from Teams")
                    
            except Exception as e:
                frappe.log_error(f"Error during bulk sync: {str(e)}", "Teams Bulk Sync Error")
                frappe.throw("Failed to sync conversations")

        # Update conversation records
        frappe.msgprint(f"Synced {synced_count} conversation(s) successfully. {error_count} errors occurred.")
        
        return {
            "success": True,
            "synced": synced_count,
            "errors": error_count
        }

    except Exception as e:
        frappe.log_error(f"Error syncing conversations: {str(e)}", "Teams Sync Conversations Error")
        frappe.throw("Failed to sync Teams conversations. Check error logs for details.")


def _sync_single_chat(chat_id, headers):
    """Sync a single chat's messages"""
    try:
        messages_url = f"{GRAPH_API}/chats/{chat_id}/messages?$top=50"
        messages_response = requests.get(messages_url, headers=headers, timeout=30)
        
        if messages_response.status_code == 200:
            messages = messages_response.json().get("value", [])
            stored_count = 0
            
            for msg in messages:
                if _save_message_local(msg, chat_id, None, None, "Inbound"):
                    stored_count += 1
            
            # Update conversation last synced time
            if frappe.db.exists("Teams Conversation", {"chat_id": chat_id}):
                frappe.db.set_value("Teams Conversation", {"chat_id": chat_id}, "last_synced", now_datetime())
            
            return True
        else:
            frappe.log_error(f"Failed to sync chat {chat_id}: {messages_response.status_code} - {messages_response.text}", "Teams Single Chat Sync Error")
            return False
            
    except Exception as e:
        frappe.log_error(f"Error syncing single chat {chat_id}: {str(e)}", "Teams Single Chat Sync Error")
        return False


@frappe.whitelist()
def get_chat_statistics(chat_id=None):
    """Get statistics about Teams chat messages"""
    try:
        filters = {}
        if chat_id:
            filters['chat_id'] = chat_id
        
        stats = {
            "total_messages": frappe.db.count("Teams Chat Message", filters),
            "inbound_messages": frappe.db.count("Teams Chat Message", {**filters, "direction": "Inbound"}),
            "outbound_messages": frappe.db.count("Teams Chat Message", {**filters, "direction": "Outbound"}),
            "unique_chats": frappe.db.sql("""
                SELECT COUNT(DISTINCT chat_id) 
                FROM `tabTeams Chat Message` 
                WHERE {}
            """.format("1=1" if not chat_id else "chat_id = %(chat_id)s"), {"chat_id": chat_id})[0][0]
        }
        
        return stats
        
    except Exception as e:
        frappe.log_error(f"Error getting chat statistics: {str(e)}", "Teams Stats Error")
        return {}