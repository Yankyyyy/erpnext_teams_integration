import frappe, requests, pytz
from .helpers import get_access_token, get_azure_user_id_by_email, get_login_url

def to_utc_isoformat(dt):
    local_tz = pytz.timezone('Asia/Kolkata')
    if dt.tzinfo is None:
        dt = local_tz.localize(dt)
    return dt.astimezone(pytz.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

@frappe.whitelist()
def create_meeting(docname):
    token = get_access_token()
    if not token:
        return get_login_url(docname)
    doc = frappe.get_doc('Event', docname)
    start = to_utc_isoformat(doc.starts_on)
    end = to_utc_isoformat(doc.ends_on)
    attendees = []
    for p in doc.event_participants:
        azure_id = None
        if p.user:
            azure_id = frappe.db.get_value('User', p.user, 'azure_object_id')
        if not azure_id and p.email:
            azure_id = get_azure_user_id_by_email(p.email)
        if azure_id:
            attendees.append({'identity': {'user': {'id': azure_id}}})
    if not attendees:
        frappe.throw('No valid participants with Azure ID found.')
    payload = {'subject': doc.subject, 'startDateTime': start, 'endDateTime': end, 'participants': {'attendees': attendees}, 'isOnlineMeeting': True}
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    res = requests.post('https://graph.microsoft.com/v1.0/me/onlineMeetings', headers=headers, json=payload)
    if res.status_code in (200,201):
        join = res.json().get('joinUrl') or res.json().get('joinWebUrl')
        doc.db_set('custom_teams_meeting_url', join)
        return 'Teams meeting link saved successfully.'
    elif res.status_code == 401:
        return get_login_url(docname)
    else:
        frappe.throw(f'Teams API Error: {res.status_code}: {res.text}')
