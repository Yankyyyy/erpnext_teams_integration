import frappe, requests, urllib.parse
from datetime import timedelta
from frappe.utils import now_datetime, get_datetime
GRAPH_API = "https://graph.microsoft.com/v1.0"


def get_settings():
    s = frappe.get_doc("Teams Settings")
    s.reload()
    return s

@frappe.whitelist()
def get_access_token():
    settings = get_settings()
    token = getattr(settings, "access_token", None)
    if not token or (getattr(settings, "token_expiry", None) and get_datetime(settings.token_expiry) < now_datetime()):
        try:
            token = refresh_access_token()
        except Exception as e:
            frappe.log_error(str(e), "Teams Token Refresh")
            token = None
    return token

@frappe.whitelist()
def refresh_access_token():
    settings = get_settings()
    data = {
        "client_id": settings.client_id,
        "client_secret": settings.client_secret,
        "grant_type": "refresh_token",
        "refresh_token": settings.refresh_token,
        "scope": "https://graph.microsoft.com/.default",
        "redirect_uri": settings.redirect_uri
    }
    res = requests.post(f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/token", data=data)
    if res.status_code != 200:
        frappe.log_error(res.text, "Teams Refresh Failed")
        raise Exception(res.text)
    token_data = res.json()
    settings.access_token = token_data.get("access_token")
    settings.refresh_token = token_data.get("refresh_token", settings.refresh_token)
    settings.token_expiry = now_datetime() + timedelta(seconds=token_data.get("expires_in", 3600))
    settings.save(ignore_permissions=True)
    frappe.db.commit()
    frappe.clear_cache(doctype="Teams Settings")
    return settings.access_token

@frappe.whitelist()
def get_azure_user_id_by_email(email):
    if not email:
        return None
    user_doc = frappe.db.get_value("User", {"email": email}, ["name", "azure_object_id"], as_dict=True)
    if user_doc and user_doc.get("azure_object_id"):
        return user_doc.azure_object_id
    token = get_access_token()
    if not token:
        return None
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(f"{GRAPH_API}/users/{email}", headers=headers)
    if res.status_code == 200:
        azure_id = res.json().get("id")
        try:
            if user_doc:
                frappe.db.set_value("User", user_doc.name, "azure_object_id", azure_id)
        except Exception:
            pass
        return azure_id
    elif res.status_code == 401:
        token = refresh_access_token()
        headers["Authorization"] = f"Bearer {token}"
        res = requests.get(f"{GRAPH_API}/users/{email}", headers=headers)
        if res.status_code == 200:
            azure_id = res.json().get("id")
            try:
                if user_doc:
                    frappe.db.set_value("User", user_doc.name, "azure_object_id", azure_id)
            except Exception:
                pass
            return azure_id
    frappe.log_error(res.text, f"Unable to fetch Object ID for {email}")
    return None

@frappe.whitelist()
def get_login_url(docname):
    settings = get_settings()
    scope = 'User.Read OnlineMeetings.ReadWrite offline_access Chat.ReadWrite Chat.Create Chat.ReadBasic User.ReadBasic.All ChannelMessage.Send'
    state = f'from_create_button::{docname}'
    login_url = (f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/authorize"
                 f"?client_id={settings.client_id}&response_type=code&redirect_uri={urllib.parse.quote(settings.redirect_uri, safe='')}&response_mode=query&scope={urllib.parse.quote(scope)}&state={urllib.parse.quote(state)}")
    return login_url