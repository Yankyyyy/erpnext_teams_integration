import frappe
import requests
from datetime import timedelta
from frappe.utils import now_datetime, cstr
from .helpers import get_settings
import json
import hashlib

@frappe.whitelist(allow_guest=True)
def callback(code=None, state=None, error=None, error_description=None):
    """Handle OAuth callback from Microsoft Teams"""
    
    # Check for OAuth errors first
    if error:
        frappe.log_error(f"OAuth Error: {error} - {error_description}", "Teams OAuth Error")
        frappe.local.response["type"] = "redirect"
        frappe.local.response["location"] = "/app/teams-settings?teams_authentication_status=error"
        return
    
    if not code:
        frappe.throw("Authorization code is missing from callback")
    
    try:
        settings = get_settings()
        
        # Validate required settings
        if not all([settings.client_id, settings.client_secret, settings.tenant_id, settings.redirect_uri]):
            frappe.throw("Teams integration is not properly configured. Please check your settings.")
        
        # Prepare token exchange request
        token_url = f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "client_id": settings.client_id,
            "client_secret": settings.client_secret,
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": settings.redirect_uri,
            "scope": "https://graph.microsoft.com/.default"
        }
        
        # Exchange code for tokens
        response = requests.post(token_url, headers=headers, data=data, timeout=30)
        
        if response.status_code != 200:
            error_data = response.json() if response.headers.get('content-type', '').startswith('application/json') else response.text
            frappe.log_error(f"Token exchange failed: {response.status_code} - {error_data}", "Teams Token Exchange Error")
            frappe.throw(f"Failed to authenticate with Microsoft Teams. Please try again.")
        
        token_data = response.json()
        
        # Update settings with new tokens
        settings.access_token = token_data.get("access_token")
        settings.refresh_token = token_data.get("refresh_token")
        
        # Calculate token expiry (subtract 5 minutes for safety buffer)
        expires_in = token_data.get("expires_in", 3600)
        settings.token_expiry = now_datetime() + timedelta(seconds=expires_in - 300)
        
        settings.save(ignore_permissions=True)
        frappe.db.commit()
        
        # Get user info and save Azure ID
        try:
            user_info_response = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {settings.access_token}"},
                timeout=30
            )
            
            if user_info_response.status_code == 200:
                user_info = user_info_response.json()
                azure_id = user_info.get("id")
                user_email = user_info.get("mail") or user_info.get("userPrincipalName")
                
                if azure_id:
                    # Update current user's Azure ID
                    if frappe.session.user != "Guest":
                        frappe.db.set_value("User", frappe.session.user, "azure_object_id", azure_id)
                    
                    # Also update based on email if available
                    if user_email and frappe.db.exists("User", {"email": user_email}):
                        frappe.db.set_value("User", {"email": user_email}, "azure_object_id", azure_id)
                    
                    # Update settings with owner info if not set
                    if not settings.azure_owner_email_id and user_email:
                        settings.azure_owner_email_id = user_email
                        settings.owner_azure_object_id = azure_id
                        settings.save(ignore_permissions=True)
                    
                    frappe.db.commit()
                    
        except Exception as e:
            # Log but don't fail the authentication process
            frappe.log_error(f"Failed to fetch user info: {str(e)}", "Teams User Info Error")
        
        # Successful authentication redirect
        redirect_url = "/app/teams-settings?teams_authentication_status=success"
        
        # If state parameter contains redirect info, use it
        if state and state.startswith('from_create_button::'):
            doc_name = state.replace('from_create_button::', '')
            if doc_name:
                redirect_url = f"/app/event/{doc_name}?teams_authentication_status=success"
        
        frappe.local.response["type"] = "redirect"
        frappe.local.response["location"] = redirect_url
        
    except Exception as e:
        frappe.log_error(f"Authentication callback error: {str(e)}", "Teams Authentication Error")
        frappe.local.response["type"] = "redirect"
        frappe.local.response["location"] = "/app/teams-settings?teams_authentication_status=error"


@frappe.whitelist()
def get_authentication_status():
    """Check if Teams integration is properly authenticated"""
    try:
        settings = get_settings()
        
        if not settings.access_token:
            return {"authenticated": False, "message": "No access token found"}
        
        # Check if token is expired
        if settings.token_expiry and settings.token_expiry < now_datetime():
            return {"authenticated": False, "message": "Token expired"}
        
        # Test the token by making a simple API call
        headers = {"Authorization": f"Bearer {settings.access_token}"}
        response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers, timeout=10)
        
        if response.status_code == 200:
            return {"authenticated": True, "message": "Authentication successful"}
        else:
            return {"authenticated": False, "message": "Token validation failed"}
            
    except Exception as e:
        frappe.log_error(f"Authentication status check failed: {str(e)}", "Teams Auth Status Error")
        return {"authenticated": False, "message": "Authentication check failed"}


@frappe.whitelist()
def revoke_authentication():
    """Revoke Teams authentication and clear tokens"""
    try:
        settings = get_settings()
        
        # Clear all authentication related fields
        settings.access_token = ""
        settings.refresh_token = ""
        settings.token_expiry = None
        settings.save(ignore_permissions=True)
        frappe.db.commit()
        
        return {"success": True, "message": "Authentication revoked successfully"}
        
    except Exception as e:
        frappe.log_error(f"Failed to revoke authentication: {str(e)}", "Teams Auth Revoke Error")
        frappe.throw("Failed to revoke authentication")