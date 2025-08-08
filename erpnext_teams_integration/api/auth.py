import frappe, requests
from datetime import timedelta
from frappe.utils import now_datetime
from .helpers import get_settings

@frappe.whitelist(allow_guest=True)
def callback(code=None, state=None):
    if not code:
        frappe.throw("Missing code")
    settings = get_settings()
    token_url = f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": settings.client_id,
        "client_secret": settings.client_secret,
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": settings.redirect_uri
    }
    res = requests.post(token_url, data=data)
    if res.status_code != 200:
        frappe.throw(f"Token exchange failed: {res.text}")
    token_data = res.json()
    settings.access_token = token_data.get("access_token")
    settings.refresh_token = token_data.get("refresh_token")
    settings.token_expiry = now_datetime() + timedelta(seconds=token_data.get("expires_in",3600))
    settings.save(ignore_permissions=True)
    frappe.db.commit()
    # get /me and save azure id on user
    user_info = requests.get("https://graph.microsoft.com/v1.0/me", headers={"Authorization": f"Bearer {settings.access_token}"}).json()
    azure_id = user_info.get("id")
    if azure_id:
        try:
            frappe.db.set_value("User", frappe.session.user, "azure_object_id", azure_id)
        except Exception:
            pass
    # redirect back
    frappe.local.response["type"] = "redirect"
    frappe.local.response["location"] = "/app/event?teams_authentication_status=success"
