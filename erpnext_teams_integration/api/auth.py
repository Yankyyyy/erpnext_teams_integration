import frappe
import requests
from werkzeug.wrappers import Response
from datetime import datetime, timedelta
from frappe.utils import now_datetime, cstr
from .helpers import get_settings, get_access_token
from werkzeug.exceptions import HTTPException
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
                if doc_name != "Teams Settings":
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

#Webhook and Subscription Management
GRAPH_API = "https://graph.microsoft.com/v1.0"

class GraphValidationResponse(HTTPException):
    def __init__(self, token):
        super().__init__()
        self.response = Response(token, status=200, mimetype='text/plain')

class GraphAcceptedResponse(HTTPException):
    def __init__(self):
        super().__init__()
        self.response = Response("Accepted", status=202, mimetype='text/plain')

# ---------------------------------------------------------------------------
# 🎧 THE WEBHOOK LISTENER
# ---------------------------------------------------------------------------
@frappe.whitelist(allow_guest=True)
def handle_graph_webhook(**kwargs):
    """
    The main listener for Microsoft Graph API Subscriptions.
    Must be a public endpoint (allow_guest=True) and accept **kwargs.
    """
    _ = frappe.request.get_data()

    # Grab token safely from the URL query parameters
    token = frappe.request.args.get('validationToken') or frappe.form_dict.get('validationToken')
    
    if token:
        raise GraphValidationResponse(token)

    # --- Handling the Actual RSVPs Below ---
    try:
        # Prevent Frappe from blocking Microsoft's POST payload due to missing session cookies
        frappe.local.flags.ignore_csrf = True
        
        payload = frappe.request.get_json()
        
        if payload and "value" in payload:
            for notification in payload.get("value", []):
                resource_url = notification.get("resource")
                if resource_url:
                    frappe.enqueue(
                        "erpnext_teams_integration.api.auth.process_rsvp_change",
                        resource_url=resource_url,
                        queue="short"
                    )
    except Exception as e:
        frappe.log_error(message=str(e), title="Webhook Payload Error")

    # Acknowledge the RSVP to Microsoft using the exact same hijack trick
    raise GraphAcceptedResponse()


# ---------------------------------------------------------------------------
# SUBSCRIPTION & SYNC LOGIC
# ---------------------------------------------------------------------------
@frappe.whitelist()
def subscribe_to_calendar_events():
    """Tells Microsoft Graph to send webhooks when meetings are updated (RSVPs)."""
    token = get_access_token()
    if not token:
        frappe.throw("Authentication required.")

    # REMINDER: Hardcoded for Ngrok testing. 
    # Move this to frappe.db.get_single_value('Teams Settings', 'webhook_base_url') for production.
    # ngrok_base = "https://ae39-115-241-89-123.ngrok-free.app"
    notification_url = frappe.utils.get_url("/api/method/erpnext_teams_integration.api.auth.handle_graph_webhook")

    # Calendar subscriptions expire in maximum 4230 minutes. We use 2 days.
    expiration_time = datetime.utcnow() + timedelta(days=2)

    payload = {
        "changeType": "updated",
        "notificationUrl": notification_url,
        "resource": "/me/events", 
        "expirationDateTime": expiration_time.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "clientState": "FrappeTeamsSyncV1"
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    res = requests.post(f"{GRAPH_API}/subscriptions", headers=headers, json=payload)

    if res.status_code == 201:
        data = res.json()
        return {"success": True, "subscription_id": data.get("id")}
    else:
        # Safely named kwargs to avoid the 140-char title limit crash
        frappe.log_error(message=res.text, title="Graph Webhook Error")
        frappe.throw(f"Failed to subscribe: {res.status_code}")
        

def process_rsvp_change(resource_url):
    """
    Background job triggered by Graph API Webhook.
    Fetches the latest event details and updates the Frappe Event Participants table.
    """
    try:
        token = get_access_token()
        if not token: 
            return
            
        if not resource_url.startswith("https"):
            url = f"{GRAPH_API}/{resource_url.lstrip('/')}"
        else:
            url = resource_url
            
        headers = {
            "Authorization": f"Bearer {token}", 
            "Content-Type": "application/json"
        }
            
        res = requests.get(url, headers=headers, timeout=30)
        
        if res.status_code != 200:
            frappe.log_error(message=res.text, title="RSVP Sync Error")
            return
            
        event_data = res.json()
        outlook_event_id = event_data.get("id")
        attendees = event_data.get("attendees", [])
        
        if not outlook_event_id or not attendees:
            return
            
        event_name = frappe.db.get_value("Event", {"custom_outlook_event_id": outlook_event_id}, "name")
        if not event_name:
            return 
            
        status_map = {
            "accepted": "Yes",
            "declined": "No",
            "tentative": "Maybe" 
        }
        
        attendee_responses = {}
        for a in attendees:
            email = a.get("emailAddress", {}).get("address", "").lower()
            status = a.get("status", {}).get("response", "").lower()
            
            if email and status in status_map:
                attendee_responses[email] = status_map[status]
                
        doc = frappe.get_doc("Event", event_name)
        doc_updated = False
        
        for row in doc.event_participants:
            row_email = (row.email or "").lower()
            if row_email in attendee_responses:
                new_status = attendee_responses[row_email]
                if row.attending != new_status:
                    row.attending = new_status
                    doc_updated = True
                    
        if doc_updated:
            doc.flags.ignore_validate = True 
            doc.flags.ignore_permissions = True
            doc.save()
            frappe.db.commit()
            
    except Exception as e:
        frappe.log_error(message=str(e), title="RSVP Processing Error")
        

def renew_graph_subscriptions():
    """
    Cron job to keep the Microsoft Graph Webhook alive.
    Runs daily.
    """
    try:
        frappe.msgprint("Renewing Microsoft Graph Subscriptions...")
        token = get_access_token()
        if not token: 
            return
            
        settings = frappe.get_single("Teams Settings")
        sub_id = settings.get("custom_webhook_subscription_id")
        
        headers = {
            "Authorization": f"Bearer {token}", 
            "Content-Type": "application/json"
        }
        
        if sub_id:
            expiration_time = datetime.utcnow() + timedelta(days=2)
            payload = {
                "expirationDateTime": expiration_time.strftime("%Y-%m-%dT%H:%M:%SZ")
            }
            
            res = requests.patch(f"{GRAPH_API}/subscriptions/{sub_id}", headers=headers, json=payload)
            
            if res.status_code == 200:
                return
            else:
                frappe.log_error(message=res.text, title="Webhook Renewal Warning")
        
        # Recreate if it failed or didn't exist
        result = subscribe_to_calendar_events()
        
        if result and result.get("success"):
            settings.db_set("custom_webhook_subscription_id", result.get("subscription_id"))
            frappe.db.commit()
            
    except Exception as e:
        error_details = str(e)
        safe_title = f"Webhook Renewal Error: {error_details}"[:135]
        frappe.log_error(message=frappe.get_traceback(), title=safe_title)