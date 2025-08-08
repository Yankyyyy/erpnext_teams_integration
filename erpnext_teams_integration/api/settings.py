import frappe, requests
from frappe import _

@frappe.whitelist()
def get_enabled_doctypes():
    settings = frappe.get_single('Teams Settings')
    enabled = []
    if getattr(settings, 'enabled_doctypes', None):
        for r in settings.enabled_doctypes:
            enabled.append(r.doctype_name)
    return enabled

@frappe.whitelist()
def bulk_sync_azure_ids():
    settings = frappe.get_single('Teams Settings')
    token = settings.access_token
    if not token:
        frappe.throw('Authenticate first.')
    res = requests.get('https://graph.microsoft.com/v1.0/users', headers={'Authorization':f'Bearer {token}'})
    if res.status_code != 200:
        frappe.throw('Failed to fetch users: ' + res.text)
    users = res.json().get('value', [])
    count = 0
    for u in users:
        email = u.get('mail') or u.get('userPrincipalName')
        azure = u.get('id')
        if email and azure and frappe.db.exists('User', {'email': email}):
            frappe.db.set_value('User', {'email': email}, 'azure_object_id', azure)
            count += 1
    frappe.db.commit()
    return f'Updated {count} users.'
