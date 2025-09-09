# apps/erpnext_teams_integration/erpnext_teams_integration/api/meetings.py

import json
from datetime import datetime, time, timedelta

import frappe
import pytz
import requests
from frappe.utils import get_datetime, now_datetime

from .helpers import get_access_token, get_azure_user_id_by_email, get_login_url

GRAPH_API = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Supported doctypes configuration
# ---------------------------------------------------------------------------

SUPPORTED_DOCTYPES = {
    "Event": {
        "participants_field": "event_participants",
        "email_field": "email",
        "subject_field": "subject",
        "start_field": "starts_on",
        "end_field": "ends_on",
    },
    "Project": {
        "participants_field": "users",
        "email_field": "email",
        "subject_field": "project_name",
        "start_field": "expected_start_date",
        "end_field": "expected_end_date",
    },
    # Add more doctypes here if needed
}

# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _safe_str(obj) -> str:
    try:
        if isinstance(obj, (dict, list)):
            return json.dumps(obj, default=str, ensure_ascii=False)
        return str(obj)
    except Exception:
        return "<unprintable>"

def safe_log_error(message: str, title: str = "Teams Integration Error"):
    """
    Log errors without tripping Frappe's 140-char title limit.
    Put the detailed content in the message body.
    """
    MAX_TITLE = 140
    title = _safe_str(title)[:MAX_TITLE]
    message = _safe_str(message)
    try:
        # frappe.log_error signature is (message=None, title=None)
        frappe.log_error(message=message, title=title)
    except Exception:
        # Last-ditch: avoid cascading failures
        pass

def to_utc_isoformat(dt, timezone_str="Asia/Kolkata"):
    """
    Convert a datetime (or parsable string) to UTC ISO 8601 with 'Z'.
    Always returns a string; if conversion fails, returns current UTC.
    """
    try:
        if not dt:
            raise ValueError("no datetime provided")

        if not isinstance(dt, datetime):
            dt = get_datetime(dt)

        try:
            local_tz = pytz.timezone(timezone_str)
        except Exception:
            local_tz = pytz.utc

        if dt.tzinfo is None:
            dt = local_tz.localize(dt)

        utc_dt = dt.astimezone(pytz.utc)
        return utc_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    except Exception as e:
        safe_log_error(f"to_utc_isoformat failed: {e}\nvalue={dt}")
        return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

def ensure_datetime_with_time(value, default_hour=9, default_minute=0):
    """
    Ensure we have a datetime with a non-midnight time component.
    If input has time 00:00:00 (likely a date), attach default business time.
    """
    try:
        if not value:
            return None

        if isinstance(value, datetime):
            dt = value
        else:
            dt = get_datetime(value)

        if dt.time() == time(0, 0, 0):
            dt = dt.replace(hour=default_hour, minute=default_minute)

        return dt
    except Exception as e:
        safe_log_error(f"ensure_datetime_with_time failed: {e}\nvalue={value}")
        return None
def _extract_meeting_id_from_join_url(join_url: str, token: str) -> str | None:
    """
    Resolve a Teams meeting ID from its JoinWebUrl using Microsoft Graph search.
    Returns the canonical meeting ID if found, otherwise None.
    """
    try:
        if not join_url:
            return None

        headers = _headers_with_auth(token, json_content=False)
        search_url = f"{GRAPH_API}/me/onlineMeetings?$filter=JoinWebUrl eq '{join_url}'"

        res = requests.get(search_url, headers=headers, timeout=30)

        if res.status_code == 401:
            # Let the caller handle re-authentication
            return None

        if res.status_code != 200:
            safe_log_error(
                message=f"Meeting lookup failed {res.status_code}: {res.text}",
                title="Teams Meeting Lookup Error",
            )
            return None

        meetings = res.json().get("value", [])
        if not meetings:
            return None

        return meetings[0].get("id")

    except Exception as e:
        safe_log_error(
            message=f"Meeting lookup exception: {str(e)}",
            title="Teams Meeting Lookup Error",
        )
        return None


def _headers_with_auth(token: str, json_content=True):
    h = {"Authorization": f"Bearer {token}"}
    if json_content:
        h["Content-Type"] = "application/json"
    return h

def _build_default_times_for_doctype(doc, doctype: str):
    """
    Build start/end datetimes (naive) for different doctypes with safe fallbacks.
    Returns (start_dt, end_dt)
    """
    cfg = SUPPORTED_DOCTYPES.get(doctype) or {}
    start_field = cfg.get("start_field")
    end_field = cfg.get("end_field")

    start_val = getattr(doc, start_field, None) if start_field else None
    end_val = getattr(doc, end_field, None) if end_field else None

    if doctype == "Project":
        # For projects, use business hours for date-only fields
        start_dt = ensure_datetime_with_time(start_val, 9, 0)
        end_dt = ensure_datetime_with_time(end_val, 17, 30)
    else:
        start_dt = ensure_datetime_with_time(start_val)
        end_dt = ensure_datetime_with_time(end_val)

    # Fallbacks if doc has nothing
    if not start_dt:
        start_dt = now_datetime()
    if not end_dt or end_dt <= start_dt:
        end_dt = start_dt + timedelta(hours=1)

    return start_dt, end_dt

def _resolve_subject(doc, doctype: str, docname: str) -> str:
    cfg = SUPPORTED_DOCTYPES.get(doctype) or {}
    subject_field = cfg.get("subject_field")
    subject = (getattr(doc, subject_field, None) or "").strip() if subject_field else ""
    return subject or f"{doctype} Meeting: {docname}"

def _build_attendees_from_participants_list(participants_list):
    """
    Given a list of Azure IDs, build the Graph attendees payload.
    """
    attendees = []
    for azure_id in participants_list:
        attendees.append({"identity": {"user": {"id": azure_id}}})
    return attendees

def _collect_participants_azure_ids(doc):
    """
    From the supported doctype, collect participant Azure Object IDs.
    """
    doctype = doc.doctype
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"Doctype {doctype} is not supported for Teams meetings.")

    cfg = SUPPORTED_DOCTYPES[doctype]
    participants_field = cfg["participants_field"]
    email_field = cfg["email_field"]

    azure_ids = set()
    rows = getattr(doc, participants_field, []) or []

    for row in rows:
        azure = None
        # Prefer linked User if present
        if getattr(row, "user", None):
            azure = frappe.db.get_value("User", row.user, "azure_object_id")

        # Fallback: resolve by email
        if not azure:
            email_val = getattr(row, email_field, None)
            if email_val:
                # First try cached user field to avoid Graph lookup if already stored
                azure = frappe.db.get_value("User", email_val, "azure_object_id")
                if not azure:
                    azure = get_azure_user_id_by_email(email_val)

        if azure:
            azure_ids.add(azure)

    return list(azure_ids)

# ---------------------------------------------------------------------------
# API: Create or Update meeting
# ---------------------------------------------------------------------------

@frappe.whitelist()
def create_meeting(docname, doctype):
    """
    Create a Teams meeting if none exists on the doc; otherwise update attendees.
    Returns dict with either success info or an auth_required + login_url directive.
    """
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"Doctype {doctype} is not supported for Teams meetings.")

    try:
        token = get_access_token()
        if not token:
            return {"error": "auth_required", "login_url": get_login_url(docname)}

        doc = frappe.get_doc(doctype, docname)

        azure_ids = _collect_participants_azure_ids(doc)
        if not azure_ids:
            frappe.throw("No valid participants with Azure ID found for meeting creation.")

        existing_meeting_url = doc.get("custom_teams_meeting_url")
        if existing_meeting_url:
            return _update_existing_meeting(doc, azure_ids, existing_meeting_url, token)

        return _create_new_meeting(doc, doctype, docname, azure_ids, token)

    except frappe.ValidationError:
        # Let explicit business validation bubble up
        raise
    except Exception as e:
        # Log detailed and show short message
        safe_log_error(
            message=f"Error creating meeting for {doctype} {docname}: {e}\n\n{frappe.get_traceback()}",
            title="Teams Meeting Create Error",
        )
        frappe.throw("Failed to create Teams meeting. Please check the error logs.")

def _update_existing_meeting(doc, azure_ids, meeting_url, token):
    """
    Add any missing attendees to an existing meeting.
    """
    try:
        meeting_id = _extract_meeting_id_from_join_url(meeting_url, token)
        if not meeting_id:
            frappe.throw("Could not extract meeting ID from existing meeting URL.")

        headers = _headers_with_auth(token, json_content=False)
        get_url = f"{GRAPH_API}/me/onlineMeetings/{meeting_id}"
        res = requests.get(get_url, headers=headers, timeout=30)

        if res.status_code == 401:
            return {"error": "auth_required", "login_url": get_login_url(doc.name)}

        if res.status_code != 200:
            safe_log_error(
                message=f"Fetch existing meeting failed {res.status_code}: {res.text}",
                title="Teams Meeting Fetch Error",
            )
            frappe.throw("Failed to fetch existing meeting details from Teams.")

        data = res.json() or {}
        existing_attendees = data.get("participants", {}).get("attendees", []) or []
        existing_ids = {a.get("identity", {}).get("user", {}).get("id") for a in existing_attendees if a}

        new_ids = [i for i in azure_ids if i not in existing_ids]
        if not new_ids:
            return {"success": True, "message": "No new participants to add to the meeting."}

        updated_attendees = existing_attendees + _build_attendees_from_participants_list(new_ids)

        patch_payload = {"participants": {"attendees": updated_attendees}}
        patch = requests.patch(
            get_url, headers=_headers_with_auth(token), json=patch_payload, timeout=30
        )
        if patch.status_code in (200, 204):
            return {"success": True, "message": f"Added {len(new_ids)} new participant(s) to the meeting."}

        safe_log_error(
            message=f"Update meeting failed {patch.status_code}: {patch.text}",
            title="Teams Meeting Update Error",
        )
        frappe.throw("Failed to update meeting participants on Teams.")

    except frappe.ValidationError:
        raise
    except Exception as e:
        safe_log_error(
            message=f"Error updating existing meeting: {e}\n\n{frappe.get_traceback()}",
            title="Teams Meeting Update Error",
        )
        frappe.throw("Failed to update existing Teams meeting. Check error logs.")

def _create_new_meeting(doc, doctype, docname, azure_ids, token):
    """
    Create a new Teams meeting with safe defaults.
    """
    try:
        subject = _resolve_subject(doc, doctype, docname)

        start_dt, end_dt = _build_default_times_for_doctype(doc, doctype)
        if start_dt >= end_dt:
            end_dt = start_dt + timedelta(hours=1)

        start_iso = to_utc_isoformat(start_dt)
        end_iso = to_utc_isoformat(end_dt)

        payload = {
            "subject": subject,
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "participants": {"attendees": _build_attendees_from_participants_list(azure_ids)},
            "isOnlineMeeting": True,
        }

        res = requests.post(
            f"{GRAPH_API}/me/onlineMeetings",
            headers=_headers_with_auth(token),
            json=payload,
            timeout=30,
        )

        if res.status_code == 401:
            return {"error": "auth_required", "login_url": get_login_url(docname)}

        if res.status_code not in (200, 201):
            short = f"Teams API error {res.status_code}"
            safe_log_error(
                message=f"Create meeting failed for {doctype} {docname}\nPayload={_safe_str(payload)}\nResponse={res.text}",
                title="Teams Meeting Creation Error",
            )
            frappe.throw(short)

        data = res.json() or {}
        join_url = data.get("joinUrl") or data.get("joinWebUrl")
        if not join_url:
            safe_log_error(message=f"No joinUrl in response: {data}", title="Teams Meeting Creation Error")
            frappe.throw("Meeting created on Teams but no join URL returned.")

        doc.db_set("custom_teams_meeting_url", join_url)
        frappe.db.commit()

        return {
            "success": True,
            "message": "Teams meeting created and link saved successfully.",
            "meeting_url": join_url,
        }

    except frappe.ValidationError:
        raise
    except Exception as e:
        safe_log_error(
            message=f"Error creating new meeting: {e}\n\n{frappe.get_traceback()}",
            title="Teams Meeting Creation Error",
        )
        frappe.throw("Failed to create new Teams meeting. See error logs.")

# ---------------------------------------------------------------------------
# API: Details
# ---------------------------------------------------------------------------

@frappe.whitelist()
def get_meeting_details(docname, doctype):
    """
    Return a summary of meeting details stored on the doc and, if possible, enriched from Graph.
    """
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        if not meeting_url:
            return {"exists": False, "message": "No Teams meeting found for this document."}
        
        token = get_access_token()

        meeting_id = _extract_meeting_id_from_join_url(meeting_url, token)
        if not meeting_id:
            return {"exists": True, "url": meeting_url, "message": "Meeting URL exists but ID not extractable."}

        if not token:
            return {
                "exists": True,
                "url": meeting_url,
                "message": "Meeting exists but cannot fetch details (authentication required).",
            }

        res = requests.get(
            f"{GRAPH_API}/me/onlineMeetings/{meeting_id}",
            headers=_headers_with_auth(token, json_content=False),
            timeout=30,
        )
        if res.status_code != 200:
            return {"exists": True, "url": meeting_url, "message": "Meeting URL exists but details unavailable."}

        data = res.json() or {}
        attendees = data.get("participants", {}).get("attendees", []) or []
        return {
            "exists": True,
            "url": meeting_url,
            "details": {
                "subject": data.get("subject"),
                "startDateTime": data.get("startDateTime"),
                "endDateTime": data.get("endDateTime"),
                "participants": len(attendees),
            },
        }

    except Exception as e:
        safe_log_error(
            message=f"Error getting meeting details for {doctype} {docname}: {e}\n\n{frappe.get_traceback()}",
            title="Teams Meeting Details Error",
        )
        return {"exists": False, "message": "Error fetching meeting details."}

# ---------------------------------------------------------------------------
# API: Delete
# ---------------------------------------------------------------------------

@frappe.whitelist()
def delete_meeting(docname, doctype):
    """
    Delete the meeting from Teams (best-effort) and clear the URL on the doc.
    """
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        if not meeting_url:
            return {"success": True, "message": "No meeting to delete."}
        
        token = get_access_token()

        meeting_id = _extract_meeting_id_from_join_url(meeting_url, token)
        if not meeting_id:
            # Can't delete remotely; clear URL locally.
            doc.db_set("custom_teams_meeting_url", "")
            frappe.db.commit()
            return {"success": True, "message": "Meeting URL cleared (could not extract meeting ID)."}

        if not token:
            return {"error": "auth_required", "message": "Authentication required to delete meeting."}

        res = requests.delete(
            f"{GRAPH_API}/me/onlineMeetings/{meeting_id}",
            headers=_headers_with_auth(token, json_content=False),
            timeout=30,
        )

        if res.status_code in (200, 204, 404):
            # 404 means: already gone → still clear locally.
            doc.db_set("custom_teams_meeting_url", "")
            frappe.db.commit()
            return {
                "success": True,
                "message": "Teams meeting deleted successfully." if res.status_code != 404 else "Meeting not found on Teams; URL cleared.",
            }

        safe_log_error(
            message=f"Delete meeting failed {res.status_code}: {res.text}",
            title="Teams Meeting Delete Error",
        )
        return {"success": False, "message": "Failed to delete meeting on Teams."}

    except Exception as e:
        safe_log_error(
            message=f"Error deleting meeting for {doctype} {docname}: {e}\n\n{frappe.get_traceback()}",
            title="Teams Meeting Delete Error",
        )
        return {"success": False, "message": "Error deleting meeting. See error logs."}

# ---------------------------------------------------------------------------
# API: Reschedule
# ---------------------------------------------------------------------------

@frappe.whitelist()
def reschedule_meeting(docname, doctype, new_start_time=None, new_end_time=None):
    """
    Reschedule an existing Teams meeting's start/end date-time.
    """
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        if not meeting_url:
            frappe.throw("No Teams meeting found to reschedule.")
            
        token = get_access_token()

        meeting_id = _extract_meeting_id_from_join_url(meeting_url, token)
        if not meeting_id:
            frappe.throw("Could not extract meeting ID from URL.")

        if not token:
            return {"error": "auth_required", "login_url": get_login_url(docname)}

        # Use new_* params or fall back to document fields
        if not new_start_time or not new_end_time:
            start_dt, end_dt = _build_default_times_for_doctype(doc, doctype)
        else:
            if doctype == "Project":
                start_dt = ensure_datetime_with_time(new_start_time, 9, 0)
                end_dt = ensure_datetime_with_time(new_end_time, 17, 30)
            else:
                start_dt = ensure_datetime_with_time(new_start_time)
                end_dt = ensure_datetime_with_time(new_end_time)

        if not start_dt or not end_dt:
            frappe.throw("Start and end times are required for rescheduling.")

        if start_dt >= end_dt:
            end_dt = start_dt + timedelta(hours=1)

        payload = {
            "startDateTime": to_utc_isoformat(start_dt),
            "endDateTime": to_utc_isoformat(end_dt),
        }

        res = requests.patch(
            f"{GRAPH_API}/me/onlineMeetings/{meeting_id}",
            headers=_headers_with_auth(token),
            json=payload,
            timeout=30,
        )

        if res.status_code in (200, 204):
            return {"success": True, "message": "Meeting rescheduled successfully."}

        if res.status_code == 401:
            return {"error": "auth_required", "login_url": get_login_url(docname)}

        safe_log_error(
            message=f"Reschedule meeting failed {res.status_code}: {res.text}",
            title="Teams Meeting Reschedule Error",
        )
        frappe.throw("Failed to reschedule meeting on Teams.")

    except frappe.ValidationError:
        raise
    except Exception as e:
        safe_log_error(
            message=f"Error rescheduling meeting for {doctype} {docname}: {e}\n\n{frappe.get_traceback()}",
            title="Teams Meeting Reschedule Error",
        )
        frappe.throw("Failed to reschedule meeting. See error logs.")

# ---------------------------------------------------------------------------
# API: Attendees
# ---------------------------------------------------------------------------

@frappe.whitelist()
def get_meeting_attendees(docname, doctype):
    """
    Fetch attendee list for the meeting (best-effort).
    """
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        if not meeting_url:
            return {"attendees": [], "message": "No meeting found."}
        
        token = get_access_token()

        meeting_id = _extract_meeting_id_from_join_url(meeting_url, token)
        if not meeting_id:
            return {"attendees": [], "message": "Could not extract meeting ID."}

        if not token:
            return {"attendees": [], "message": "Authentication required."}

        res = requests.get(
            f"{GRAPH_API}/me/onlineMeetings/{meeting_id}",
            headers=_headers_with_auth(token, json_content=False),
            timeout=30,
        )
        if res.status_code != 200:
            return {"attendees": [], "message": f"Failed to fetch attendees: {res.status_code}"}

        data = res.json() or {}
        attendees = data.get("participants", {}).get("attendees", []) or []

        out = []
        for a in attendees:
            user = (a or {}).get("identity", {}).get("user", {}) or {}
            out.append(
                {
                    "id": user.get("id"),
                    "displayName": user.get("displayName") or "Unknown",
                    "email": user.get("email"),
                }
            )

        return {"attendees": out, "count": len(out)}

    except Exception as e:
        safe_log_error(
            message=f"Error getting meeting attendees for {doctype} {docname}: {e}\n\n{frappe.get_traceback()}",
            title="Teams Meeting Attendees Error",
        )
        return {"attendees": [], "message": "Error fetching attendees."}

# ---------------------------------------------------------------------------
# API: Validation helper (optional)
# ---------------------------------------------------------------------------

@frappe.whitelist()
def validate_meeting_time(start_time, end_time, timezone_str="Asia/Kolkata"):
    """
    Validate start/end times; returns {valid, errors, duration_hours}.
    """
    try:
        start_dt = get_datetime(start_time)
        end_dt = get_datetime(end_time)

        errors = []

        if start_dt >= end_dt:
            errors.append("End time must be after start time.")

        duration = end_dt - start_dt
        if duration.total_seconds() > 24 * 3600:
            errors.append("Meeting duration cannot exceed 24 hours.")
        if duration.total_seconds() < 15 * 60:
            errors.append("Meeting duration should be at least 15 minutes.")

        if start_dt < now_datetime():
            errors.append("Meeting cannot be scheduled in the past.")

        return {
            "valid": len(errors) == 0,
            "errors": errors,
            "duration_hours": round(duration.total_seconds() / 3600, 2),
        }
    except Exception as e:
        return {"valid": False, "errors": [f"Invalid date/time format: {e}"]}
