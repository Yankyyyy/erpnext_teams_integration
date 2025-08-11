import frappe, requests, pytz
from .helpers import get_access_token, get_azure_user_id_by_email, get_login_url

SUPPORTED_DOCTYPES = {
    "Event": {"participants_field": "event_participants", "email_field": "email", "subject_field": "subject"},
    "Project": {"participants_field": "users", "email_field": "email", "subject_field": "project_name"},
    # Add more doctypes as needed
}

def to_utc_isoformat(dt):
    local_tz = pytz.timezone('Asia/Kolkata')
    if dt.tzinfo is None:
        dt = local_tz.localize(dt)
    return dt.astimezone(pytz.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

from datetime import datetime, time

def ensure_datetime(value, default_hour=00, default_minute=00, default_second=00):
    """Ensure a value is a datetime; if it's a date, attach default time."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        try:
            # Try parsing datetime string first
            return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            try:
                # Try parsing date string only
                date_obj = datetime.strptime(value, "%Y-%m-%d").date()
                return datetime.combine(date_obj, time(default_hour, default_minute, default_second))
            except ValueError:
                frappe.throw(f"Invalid date format: {value}")
    # If it's already a date object, combine with default time
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return datetime.combine(value, time(default_hour, default_minute, default_second))
    return value

def get_participants(doc):
    """Fetch Azure user IDs for participants from any supported doctype."""
    doctype = doc.doctype
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"Doctype {doctype} is not supported for Teams meetings.")
    mapping = SUPPORTED_DOCTYPES[doctype]
    participants_field = mapping["participants_field"]
    email_field = mapping["email_field"]

    participants = []
    for p in doc.get(participants_field):
        azure_id = None
        if getattr(p, "user", None):
            azure_id = frappe.db.get_value("User", p.user, "azure_object_id")
        if not azure_id and getattr(p, email_field, None):
            azure_id = get_azure_user_id_by_email(getattr(p, email_field))
        if azure_id:
            participants.append({"identity": {"user": {"id": azure_id}}})
    return participants

@frappe.whitelist()
def create_meeting(docname, doctype):
    token = get_access_token()
    if not token:
        return get_login_url(docname)

    doc = frappe.get_doc(doctype, docname)
    participants = get_participants(doc)

    if not participants:
        frappe.throw("No valid participants with Azure ID found.")

    meeting_url = doc.get("custom_teams_meeting_url")

    # Case 1: Meeting already exists
    if meeting_url:
        # Extract meeting ID from join URL (last segment after /)
        meeting_id = meeting_url.split("/")[-1]

        # Get existing participants from Teams
        headers = {"Authorization": f"Bearer {token}"}
        existing_res = requests.get(
            f"https://graph.microsoft.com/v1.0/me/onlineMeetings/{meeting_id}",
            headers=headers
        )
        if existing_res.status_code == 200:
            existing_data = existing_res.json()
            existing_attendees = {
                a["identity"]["user"]["id"]
                for a in existing_data.get("participants", {}).get("attendees", [])
            }

            # Find new participants
            new_participants = [
                p for p in participants
                if p["identity"]["user"]["id"] not in existing_attendees
            ]

            if new_participants:
                patch_payload = {
                    "participants": {
                        "attendees": list(existing_data.get("participants", {}).get("attendees", [])) + new_participants
                    }
                }
                patch_res = requests.patch(
                    f"https://graph.microsoft.com/v1.0/me/onlineMeetings/{meeting_id}",
                    headers={**headers, "Content-Type": "application/json"},
                    json=patch_payload
                )
                if patch_res.status_code in (200, 204):
                    return f"Added {len(new_participants)} new participant(s) to the meeting."
                else:
                    frappe.throw(f"Failed to add participants: {patch_res.status_code} {patch_res.text}")
            else:
                return "No new participants to add."

        elif existing_res.status_code == 401:
            return get_login_url(docname)
        else:
            frappe.throw(f"Error fetching meeting: {existing_res.status_code} {existing_res.text}")

    # Case 2: Create new meeting
    if doctype == "Event":
        start = to_utc_isoformat(doc.starts_on)
        end = to_utc_isoformat(doc.ends_on)
    if doctype == "Project":
        start = to_utc_isoformat(ensure_datetime(doc.expected_start_date, 00, 00, 00))   # start of day 9 AM
        end = to_utc_isoformat(ensure_datetime(doc.expected_end_date, 00, 00, 00))     # end of day 5:45:06 PM
    
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"Doctype {doctype} is not supported for Teams meetings.")
    mapping = SUPPORTED_DOCTYPES[doctype]
    subject = getattr(doc, mapping["subject_field"])

    payload = {
        "subject": subject or f"{doctype} Meeting: {docname}",
        "startDateTime": start,
        "endDateTime": end,
        "participants": {"attendees": participants},
        "isOnlineMeeting": True
    }
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    res = requests.post(
        "https://graph.microsoft.com/v1.0/me/onlineMeetings",
        headers=headers,
        json=payload
    )
    if res.status_code in (200, 201):
        join = res.json().get("joinUrl") or res.json().get("joinWebUrl")
        doc.db_set("custom_teams_meeting_url", join)
        return "Teams meeting link saved successfully."
    elif res.status_code == 401:
        return get_login_url(docname)
    else:
        frappe.throw(f"Teams API Error: {res.status_code}: {res.text}")