import frappe
import requests
from frappe import _
from .helpers import get_access_token, get_settings
import json
from frappe.utils import cstr


@frappe.whitelist()
def get_enabled_doctypes():
    """Get list of enabled doctypes for Teams integration"""
    try:
        settings = get_settings()
        enabled_doctypes = []
        
        if hasattr(settings, 'enabled_doctypes') and settings.enabled_doctypes:
            for row in settings.enabled_doctypes:
                if row.doctype_name:
                    enabled_doctypes.append(row.doctype_name)
        
        return enabled_doctypes
        
    except Exception as e:
        frappe.log_error(f"Error getting enabled doctypes: {str(e)}", "Teams Settings Error")
        return []


@frappe.whitelist()
def bulk_sync_azure_ids():
    """Bulk sync Azure Object IDs for all users"""
    try:
        settings = get_settings()
        token = get_access_token()
        
        if not token:
            frappe.throw('Please authenticate with Microsoft Teams first.')
        
        headers = {'Authorization': f'Bearer {token}'}
        
        # Fetch all users from Microsoft Graph API with pagination
        all_users = []
        url = 'https://graph.microsoft.com/v1.0/users'
        
        while url:
            response = requests.get(url, headers=headers, timeout=30)
            
            if response.status_code != 200:
                frappe.log_error(f"Failed to fetch users from Graph API: {response.text}", "Teams Bulk Sync Error")
                frappe.throw(f'Failed to fetch users from Microsoft Graph: {response.status_code}')
            
            data = response.json()
            all_users.extend(data.get('value', []))
            
            # Get next page URL if available
            url = data.get('@odata.nextLink')
        
        if not all_users:
            frappe.msgprint("No users found in Microsoft Graph API")
            return "No users found to sync"
        
        # Update Frappe users with Azure IDs
        updated_count = 0
        created_count = 0
        error_count = 0
        
        for graph_user in all_users:
            try:
                # Get email and Azure ID
                email = graph_user.get('mail') or graph_user.get('userPrincipalName')
                azure_id = graph_user.get('id')
                display_name = graph_user.get('displayName', '')
                
                if not email or not azure_id:
                    continue
                
                # Check if Frappe user exists
                existing_user = frappe.db.get_value('User', {'email': email}, ['name', 'azure_object_id'], as_dict=True)
                
                if existing_user:
                    # Update existing user
                    if existing_user.azure_object_id != azure_id:
                        frappe.db.set_value('User', existing_user.name, 'azure_object_id', azure_id)
                        updated_count += 1
                else:
                    # Optionally create new user (commented out for safety)
                    # You might want to enable this based on your requirements
                    """
                    try:
                        new_user = frappe.get_doc({
                            "doctype": "User",
                            "email": email,
                            "first_name": display_name.split(' ')[0] if display_name else email.split('@')[0],
                            "azure_object_id": azure_id,
                            "send_welcome_email": 0,
                            "enabled": 0  # Create as disabled by default
                        })
                        new_user.insert(ignore_permissions=True)
                        created_count += 1
                    except Exception as create_error:
                        frappe.log_error(f"Failed to create user {email}: {str(create_error)}", "Teams User Creation Error")
                        error_count += 1
                    """
                    pass
                    
            except Exception as user_error:
                frappe.log_error(f"Error processing user {email}: {str(user_error)}", "Teams User Sync Error")
                error_count += 1
                continue
        
        frappe.db.commit()
        
        # Update settings with owner info if not set
        if not settings.azure_owner_email_id and settings.access_token:
            try:
                me_response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers, timeout=30)
                if me_response.status_code == 200:
                    me_data = me_response.json()
                    owner_email = me_data.get('mail') or me_data.get('userPrincipalName')
                    owner_azure_id = me_data.get('id')
                    
                    if owner_email and owner_azure_id:
                        settings.azure_owner_email_id = owner_email
                        settings.owner_azure_object_id = owner_azure_id
                        settings.save(ignore_permissions=True)
            except Exception as owner_error:
                frappe.log_error(f"Failed to update owner info: {str(owner_error)}", "Teams Owner Update Error")
        
        result_message = f'Sync completed: {updated_count} users updated'
        # if created_count > 0:
        #     result_message += f', {created_count} users created'
        if error_count > 0:
            result_message += f', {error_count} errors occurred'
        
        frappe.msgprint(result_message)
        return result_message
        
    except Exception as e:
        frappe.log_error(f"Bulk Azure ID sync failed: {str(e)}", "Teams Bulk Sync Error")
        frappe.throw(f'Failed to sync Azure IDs: {str(e)}')


@frappe.whitelist()
def test_teams_connection():
    """Test connection to Microsoft Teams API"""
    try:
        token = get_access_token()
        if not token:
            return {
                "success": False,
                "message": "No access token available. Please authenticate first."
            }

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        # Test basic API access
        me_response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers, timeout=30)

        if me_response.status_code == 200:
            user_data = me_response.json()

            # Test chats access
            chats_response = requests.get("https://graph.microsoft.com/v1.0/chats?$top=1", headers=headers, timeout=30)
            chats_access = chats_response.status_code in (200, 204)

            # Test meetings access by creating a dummy meeting
            dummy_meeting = {
                "startDateTime": "2025-01-01T12:00:00Z",
                "endDateTime": "2025-01-01T12:30:00Z",
                "subject": "Permission Test Meeting"
            }

            meetings_access = False
            meetings_response = requests.post(
                "https://graph.microsoft.com/v1.0/me/onlineMeetings",
                headers=headers,
                json=dummy_meeting,
                timeout=30
            )

            if meetings_response.status_code == 201:
                meetings_access = True
                # cleanup to avoid leaving a meeting behind
                meeting_id = meetings_response.json().get("id")
                if meeting_id:
                    try:
                        requests.delete(
                            f"https://graph.microsoft.com/v1.0/me/onlineMeetings/{meeting_id}",
                            headers=headers,
                            timeout=30
                        )
                    except Exception:
                        pass  # safe ignore cleanup error

            return {
                "success": True,
                "message": "Connection successful",
                "user_info": {
                    "name": user_data.get("displayName"),
                    "email": user_data.get("mail") or user_data.get("userPrincipalName"),
                    "id": user_data.get("id")
                },
                "permissions": {
                    "chats": chats_access,
                    "meetings": meetings_access
                }
            }
        else:
            return {
                "success": False,
                "message": f"API connection failed: {me_response.status_code}"
            }

    except Exception as e:
        frappe.log_error(f"Teams connection test failed: {str(e)}", "Teams Connection Test Error")
        return {
            "success": False,
            "message": f"Connection test failed: {str(e)}"
        }


@frappe.whitelist()
def get_teams_statistics():
    """Get statistics about Teams integration usage"""
    try:
        stats = {
            "total_conversations": frappe.db.count("Teams Conversation"),
            "total_messages": frappe.db.count("Teams Chat Message"),
            "inbound_messages": frappe.db.count("Teams Chat Message", {"direction": "Inbound"}),
            "outbound_messages": frappe.db.count("Teams Chat Message", {"direction": "Outbound"}),
            "unique_chats": frappe.db.sql("SELECT COUNT(DISTINCT chat_id) FROM `tabTeams Chat Message`")[0][0],
            "users_with_azure_id": frappe.db.count("User", {"azure_object_id": ["!=", ""]})
        }
        
        # Get recent activity (last 7 days)
        recent_messages = frappe.db.sql("""
            SELECT DATE(created_at) as date, COUNT(*) as count 
            FROM `tabTeams Chat Message` 
            WHERE created_at >= DATE_SUB(NOW(), INTERVAL 7 DAY)
            GROUP BY DATE(created_at)
            ORDER BY date DESC
        """, as_dict=True)
        
        stats["recent_activity"] = recent_messages
        
        # Get top active chats
        top_chats = frappe.db.sql("""
            SELECT chat_id, COUNT(*) as message_count,
                   MAX(created_at) as last_activity
            FROM `tabTeams Chat Message`
            GROUP BY chat_id
            ORDER BY message_count DESC
            LIMIT 5
        """, as_dict=True)
        
        stats["top_chats"] = top_chats
        
        return stats
        
    except Exception as e:
        frappe.log_error(f"Error getting Teams statistics: {str(e)}", "Teams Statistics Error")
        return {}


@frappe.whitelist()
def cleanup_old_messages(days=30):
    """Clean up old Teams messages to save space"""
    try:
        if not isinstance(days, int) or days < 1:
            frappe.throw("Days must be a positive integer")
        
        # Delete messages older than specified days
        result = frappe.db.sql("""
            DELETE FROM `tabTeams Chat Message` 
            WHERE created_at < DATE_SUB(NOW(), INTERVAL %s DAY)
        """, (days,))
        
        deleted_count = result[0] if result else 0
        
        frappe.db.commit()
        
        message = f"Deleted {deleted_count} messages older than {days} days"
        frappe.msgprint(message)
        return message
        
    except Exception as e:
        frappe.log_error(f"Error cleaning up old messages: {str(e)}", "Teams Cleanup Error")
        frappe.throw(f"Failed to cleanup old messages: {str(e)}")


@frappe.whitelist()
def export_chat_history(chat_id=None, format="json"):
    """Export chat history for backup or analysis"""
    try:
        filters = {}
        if chat_id:
            filters["chat_id"] = chat_id
        
        messages = frappe.get_all(
            "Teams Chat Message",
            filters=filters,
            fields=["*"],
            order_by="created_at asc"
        )
        
        if format.lower() == "csv":
            import csv
            import io
            
            output = io.StringIO()
            writer = csv.DictWriter(output, fieldnames=messages[0].keys() if messages else [])
            writer.writeheader()
            writer.writerows(messages)
            
            return {
                "data": output.getvalue(),
                "filename": f"teams_chat_export_{chat_id or 'all'}_{frappe.utils.now()}.csv",
                "content_type": "text/csv"
            }
        else:
            return {
                "data": json.dumps(messages, indent=2, default=str),
                "filename": f"teams_chat_export_{chat_id or 'all'}_{frappe.utils.now()}.json",
                "content_type": "application/json"
            }
            
    except Exception as e:
        frappe.log_error(f"Error exporting chat history: {str(e)}", "Teams Export Error")
        frappe.throw(f"Failed to export chat history: {str(e)}")


@frappe.whitelist()
def validate_configuration():
    """Validate Teams integration configuration"""
    try:
        settings = get_settings()
        issues = []
        
        # Check required fields
        required_fields = {
            'client_id': 'Client ID',
            'client_secret': 'Client Secret',
            'tenant_id': 'Tenant ID',
            'redirect_uri': 'Redirect URI'
        }
        
        for field, label in required_fields.items():
            value = getattr(settings, field, None)
            if not value:
                issues.append(f"Missing {label}")
            elif len(cstr(value).strip()) < 5:
                issues.append(f"{label} appears too short")
        
        # Validate redirect URI
        if settings.redirect_uri:
            if not settings.redirect_uri.startswith(('http://', 'https://')):
                issues.append("Redirect URI must start with http:// or https://")
            
            # Check if redirect URI points to current site
            site_url = frappe.utils.get_url()
            if not settings.redirect_uri.startswith(site_url):
                issues.append("Redirect URI should point to current site")
        
        # Check authentication status
        token = get_access_token()
        if not token:
            issues.append("Not authenticated with Microsoft Teams")
        else:
            # Test token validity
            try:
                headers = {'Authorization': f'Bearer {token}'}
                response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers, timeout=10)
                if response.status_code != 200:
                    issues.append("Access token appears to be invalid")
            except:
                issues.append("Unable to validate access token")
        
        # Check enabled doctypes
        enabled_doctypes = get_enabled_doctypes()
        if not enabled_doctypes:
            issues.append("No doctypes enabled for Teams integration")
        
        return {
            "valid": len(issues) == 0,
            "issues": issues,
            "configuration_complete": len(issues) == 0
        }
        
    except Exception as e:
        frappe.log_error(f"Configuration validation error: {str(e)}", "Teams Config Validation Error")
        return {
            "valid": False,
            "issues": ["Configuration validation failed"],
            "configuration_complete": False
        }


@frappe.whitelist()
def reset_integration():
    """Reset Teams integration (clear all tokens and data)"""
    try:
        # Confirm this is intentional
        if not frappe.confirm("This will clear all Teams authentication data and conversation history. Are you sure?"):
            return
        
        settings = get_settings()
        
        # Clear authentication data
        settings.access_token = ""
        settings.refresh_token = ""
        settings.token_expiry = None
        settings.azure_owner_email_id = ""
        settings.owner_azure_object_id = ""
        settings.save(ignore_permissions=True)
        
        # Clear Azure Object IDs from users (optional)
        # frappe.db.sql("UPDATE `tabUser` SET azure_object_id = ''")
        
        # Clear conversation data (optional - commented out for safety)
        # frappe.db.sql("DELETE FROM `tabTeams Conversation`")
        # frappe.db.sql("DELETE FROM `tabTeams Chat Message`")
        
        frappe.db.commit()
        frappe.clear_cache()
        
        return "Teams integration has been reset. Please reconfigure and authenticate."
        
    except Exception as e:
        frappe.log_error(f"Error resetting integration: {str(e)}", "Teams Reset Error")
        frappe.throw(f"Failed to reset integration: {str(e)}")


@frappe.whitelist()
def get_oauth_scopes():
    """Get list of OAuth scopes required for the integration"""
    return [
        "User.Read",
        "User.ReadBasic.All", 
        "OnlineMeetings.ReadWrite",
        "offline_access",
        "Chat.ReadWrite",
        "Chat.Create",
        "Chat.ReadBasic",
        "ChannelMessage.Send",
        "Chat.ReadWrite.All"
    ]