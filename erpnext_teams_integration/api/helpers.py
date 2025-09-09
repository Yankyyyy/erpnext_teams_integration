import frappe
import requests
import urllib.parse
from datetime import timedelta
from frappe.utils import now_datetime, get_datetime

GRAPH_API = "https://graph.microsoft.com/v1.0"


def get_settings():
    """Get Teams Settings singleton with proper error handling"""
    try:
        settings = frappe.get_doc("Teams Settings")
        return settings
    except frappe.DoesNotExistError:
        frappe.throw("Teams Settings not found. Please configure Teams integration first.")
    except Exception as e:
        frappe.log_error(f"Failed to get Teams settings: {str(e)}", "Teams Settings Error")
        frappe.throw("Failed to load Teams settings")


@frappe.whitelist()
def get_access_token():
    """Get valid access token, refresh if needed"""
    try:
        settings = get_settings()
        
        # Check if we have a token
        if not settings.access_token:
            return None
        
        # Check if token is expired or will expire soon (5 minutes buffer)
        if settings.token_expiry:
            expiry_time = get_datetime(settings.token_expiry)
            current_time = now_datetime()
            time_until_expiry = expiry_time - current_time
            
            # If token expires in less than 5 minutes, refresh it
            if time_until_expiry.total_seconds() < 300:
                try:
                    return refresh_access_token()
                except Exception as e:
                    frappe.log_error(f"Token refresh failed: {str(e)}", "Teams Token Refresh Error")
                    return None
        
        return settings.access_token
        
    except Exception as e:
        frappe.log_error(f"Failed to get access token: {str(e)}", "Teams Token Error")
        return None


@frappe.whitelist()
def refresh_access_token():
    """Refresh access token using refresh token"""
    try:
        settings = get_settings()
        
        if not settings.refresh_token:
            frappe.throw("No refresh token available. Please re-authenticate.")
        
        # Prepare refresh request
        token_url = f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": settings.client_id,
            "client_secret": settings.client_secret,
            "grant_type": "refresh_token",
            "refresh_token": settings.refresh_token,
            "scope": "https://graph.microsoft.com/.default"
        }
        
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        
        response = requests.post(token_url, data=data, headers=headers, timeout=30)
        
        if response.status_code != 200:
            error_data = response.text
            try:
                error_json = response.json()
                error_data = error_json.get('error_description', error_data)
            except:
                pass
            
            frappe.log_error(f"Token refresh failed: {response.status_code} - {error_data}", "Teams Token Refresh Error")
            
            # If refresh token is invalid, clear all tokens
            if response.status_code == 400:
                settings.access_token = ""
                settings.refresh_token = ""
                settings.token_expiry = None
                settings.save(ignore_permissions=True)
                frappe.db.commit()
            
            frappe.throw("Failed to refresh access token. Please re-authenticate.")
        
        token_data = response.json()
        
        # Update tokens
        settings.access_token = token_data.get("access_token")
        
        # Update refresh token if provided (some OAuth flows provide new refresh token)
        if token_data.get("refresh_token"):
            settings.refresh_token = token_data.get("refresh_token")
        
        # Calculate new expiry (subtract 5 minutes for safety buffer)
        expires_in = token_data.get("expires_in", 3600)
        settings.token_expiry = now_datetime() + timedelta(seconds=expires_in - 300)
        
        settings.save(ignore_permissions=True)
        frappe.db.commit()
        frappe.clear_cache(doctype="Teams Settings")
        
        return settings.access_token
        
    except requests.exceptions.Timeout:
        frappe.log_error("Token refresh request timed out", "Teams Token Refresh Timeout")
        frappe.throw("Authentication request timed out. Please try again.")
    except requests.exceptions.RequestException as e:
        frappe.log_error(f"Network error during token refresh: {str(e)}", "Teams Token Refresh Network Error")
        frappe.throw("Network error occurred during authentication. Please check your connection.")
    except Exception as e:
        frappe.log_error(f"Unexpected error during token refresh: {str(e)}", "Teams Token Refresh Error")
        frappe.throw("An unexpected error occurred during authentication.")


@frappe.whitelist()
def get_azure_user_id_by_email(email):
    """Get Azure user ID by email address with caching"""
    if not email:
        return None
    
    try:
        # First check if we already have the Azure ID in our database
        user_doc = frappe.db.get_value("User", {"email": email}, ["name", "azure_object_id"], as_dict=True)
        if user_doc and user_doc.get("azure_object_id"):
            return user_doc.azure_object_id
        
        # Get access token
        token = get_access_token()
        if not token:
            frappe.log_error(f"No access token available to fetch Azure ID for {email}", "Teams API Error")
            return None
        
        headers = {"Authorization": f"Bearer {token}"}
        
        # Try to get user by email
        encoded_email = urllib.parse.quote(email, safe='')
        url = f"{GRAPH_API}/users/{encoded_email}"
        
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            azure_id = response.json().get("id")
            
            # Cache the Azure ID in our database
            if azure_id and user_doc:
                try:
                    frappe.db.set_value("User", user_doc.name, "azure_object_id", azure_id)
                    frappe.db.commit()
                except Exception as e:
                    frappe.log_error(f"Failed to cache Azure ID for {email}: {str(e)}", "Teams Cache Error")
            
            return azure_id
            
        elif response.status_code == 401:
            # Token might be expired, try to refresh
            try:
                token = refresh_access_token()
                headers["Authorization"] = f"Bearer {token}"
                
                response = requests.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    azure_id = response.json().get("id")
                    
                    # Cache the Azure ID
                    if azure_id and user_doc:
                        try:
                            frappe.db.set_value("User", user_doc.name, "azure_object_id", azure_id)
                            frappe.db.commit()
                        except Exception:
                            pass
                    
                    return azure_id
            except Exception as e:
                frappe.log_error(f"Failed to refresh token while fetching Azure ID for {email}: {str(e)}", "Teams Token Error")
        
        elif response.status_code == 404:
            frappe.log_error(f"User not found in Azure AD: {email}", "Teams User Not Found")
        else:
            frappe.log_error(f"Failed to fetch Azure ID for {email}: {response.status_code} - {response.text}", "Teams API Error")
        
        return None
        
    except requests.exceptions.Timeout:
        frappe.log_error(f"Timeout while fetching Azure ID for {email}", "Teams API Timeout")
        return None
    except requests.exceptions.RequestException as e:
        frappe.log_error(f"Network error while fetching Azure ID for {email}: {str(e)}", "Teams Network Error")
        return None
    except Exception as e:
        frappe.log_error(f"Unexpected error while fetching Azure ID for {email}: {str(e)}", "Teams API Error")
        return None


@frappe.whitelist()
def get_login_url(docname=None):
    """Generate Microsoft Teams OAuth login URL"""
    try:
        settings = get_settings()
        
        # Validate required settings
        if not all([settings.client_id, settings.tenant_id, settings.redirect_uri]):
            frappe.throw("Teams integration is not properly configured. Please check Client ID, Tenant ID, and Redirect URI.")
        
        # Required scopes for the integration
        scope = 'User.Read OnlineMeetings.ReadWrite offline_access Chat.ReadWrite Chat.Create Chat.ReadBasic User.ReadBasic.All ChannelMessage.Send'
        state = f'from_create_button::{docname}'
        login_url = (f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/authorize"
                    f"?client_id={settings.client_id}&response_type=code&redirect_uri={urllib.parse.quote(settings.redirect_uri, safe='')}&response_mode=query&scope={urllib.parse.quote(scope)}&state={urllib.parse.quote(state)}")
        return login_url
        
    except Exception as e:
        frappe.log_error(f"Failed to generate login URL: {str(e)}", "Teams Login URL Error")
        frappe.throw("Failed to generate authentication URL")


@frappe.whitelist()
def validate_settings():
    """Validate Teams Settings configuration"""
    try:
        settings = get_settings()
        errors = []
        
        # Check required fields
        required_fields = {
            'client_id': 'Client ID',
            'client_secret': 'Client Secret', 
            'tenant_id': 'Tenant ID',
            'redirect_uri': 'Redirect URI'
        }
        
        for field, label in required_fields.items():
            if not getattr(settings, field, None):
                errors.append(f"{label} is required")
        
        # Validate redirect URI format
        if settings.redirect_uri:
            if not settings.redirect_uri.startswith(('http://', 'https://')):
                errors.append("Redirect URI must start with http:// or https://")
        
        # Validate tenant ID format (should be a GUID)
        if settings.tenant_id:
            import re
            guid_pattern = r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
            if not re.match(guid_pattern, settings.tenant_id.lower()):
                errors.append("Tenant ID should be a valid GUID format")
        
        return {
            "valid": len(errors) == 0,
            "errors": errors
        }
        
    except Exception as e:
        frappe.log_error(f"Settings validation error: {str(e)}", "Teams Settings Validation Error")
        return {
            "valid": False,
            "errors": ["Failed to validate settings"]
        }


@frappe.whitelist()
def test_api_connection():
    """Test API connection with current tokens"""
    try:
        token = get_access_token()
        if not token:
            return {
                "success": False,
                "message": "No valid access token available"
            }
        
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(f"{GRAPH_API}/me", headers=headers, timeout=10)
        
        if response.status_code == 200:
            user_data = response.json()
            return {
                "success": True,
                "message": "API connection successful",
                "user": {
                    "name": user_data.get("displayName"),
                    "email": user_data.get("mail") or user_data.get("userPrincipalName")
                }
            }
        else:
            return {
                "success": False,
                "message": f"API connection failed: {response.status_code}"
            }
            
    except Exception as e:
        frappe.log_error(f"API connection test failed: {str(e)}", "Teams API Test Error")
        return {
            "success": False,
            "message": "API connection test failed"
        }