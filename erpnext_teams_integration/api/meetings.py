import frappe
import requests
import pytz
from .helpers import get_access_token, get_azure_user_id_by_email, get_login_url
from frappe.utils import get_datetime, now_datetime
from datetime import datetime, time, timedelta
import json

GRAPH_API = "https://graph.microsoft.com/v1.0"

# Supported doctypes configuration
SUPPORTED_DOCTYPES = {
    "Event": {
        "participants_field": "event_participants", 
        "email_field": "email", 
        "subject_field": "subject",
        "start_field": "starts_on",
        "end_field": "ends_on"
    },
    "Project": {
        "participants_field": "users", 
        "email_field": "email", 
        "subject_field": "project_name",
        "start_field": "expected_start_date",
        "end_field": "expected_end_date"
    },
}


def to_utc_isoformat(dt, timezone_str='Asia/Kolkata'):
    """Convert datetime to UTC ISO format string"""
    try:
        if not dt:
            return None
            
        # If it's already a datetime
        if isinstance(dt, datetime):
            local_dt = dt
        else:
            # Convert string or date to datetime
            local_dt = get_datetime(dt)
        
        # Get timezone
        try:
            local_tz = pytz.timezone(timezone_str)
        except:
            local_tz = pytz.timezone('UTC')
        
        # Localize if naive
        if local_dt.tzinfo is None:
            local_dt = local_tz.localize(local_dt)
        
        # Convert to UTC
        utc_dt = local_dt.astimezone(pytz.utc)
        
        # Return ISO format
        return utc_dt.strftime('%Y-%m-%dT%H:%M:%SZ')
        
    except Exception as e:
        frappe.log_error(f"Error converting datetime to UTC: {str(e)}", "Teams DateTime Error")
        # Return current time as fallback
        return datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')


def ensure_datetime_with_time(value, default_hour=9, default_minute=0):
    """Ensure a value is a datetime with time component"""
    try:
        if not value:
            return None
            
        # If it's already a datetime, return as-is
        if isinstance(value, datetime):
            return value
            
        # Convert to datetime
        dt = get_datetime(value)
        
        # If it's midnight (00:00:00), it was probably a date-only field
        if dt.time() == time(0, 0, 0):
            # Replace with business hours
            dt = dt.replace(hour=default_hour, minute=default_minute)
            
        return dt
        
    except Exception as e:
        frappe.log_error(f"Error converting to datetime: {str(e)}", "Teams DateTime Error")
        return None


def get_participants_azure_ids(doc):
    """Get Azure user IDs for document participants"""
    doctype = doc.doctype
    
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"Doctype {doctype} is not supported for Teams meetings.")
    
    config = SUPPORTED_DOCTYPES[doctype]
    participants_field = config["participants_field"]
    email_field = config["email_field"]
    
    participants = []
    participants_data = getattr(doc, participants_field, [])
    
    for participant in participants_data:
        azure_id = None
        
        # Try to get from user field first
        if hasattr(participant, 'user') and participant.user:
            azure_id = frappe.db.get_value("User", participant.user, "azure_object_id")
        
        # If not found, try email field
        if not azure_id:
            email = getattr(participant, email_field, None)
            if email:
                azure_id = get_azure_user_id_by_email(email)
        
        if azure_id:
            participants.append({
                "identity": {
                    "user": {
                        "id": azure_id
                    }
                }
            })
    
    return participants


@frappe.whitelist()
def create_meeting(docname, doctype):
    """Create or update Teams meeting for a document"""
    if doctype not in SUPPORTED_DOCTYPES:
        frappe.throw(f"Doctype {doctype} is not supported for Teams meetings.")
    
    try:
        # Get access token
        token = get_access_token()
        if not token:
            login_url = get_login_url(docname)
            return {"error": "auth_required", "login_url": login_url}
        
        # Get document
        doc = frappe.get_doc(doctype, docname)
        config = SUPPORTED_DOCTYPES[doctype]
        
        # Get participants
        participants = get_participants_azure_ids(doc)
        if not participants:
            frappe.throw("No valid participants with Azure ID found for meeting creation.")
        
        # Check if meeting already exists
        existing_meeting_url = doc.get("custom_teams_meeting_url")
        
        if existing_meeting_url:
            return update_existing_meeting(doc, config, participants, existing_meeting_url, token)
        else:
            return create_new_meeting(doc, config, participants, token)
            
    except Exception as e:
        frappe.log_error(f"Error creating meeting for {doctype} {docname}: {str(e)}", "Teams Meeting Error")
        frappe.throw(f"Failed to create Teams meeting: {str(e)}")


def update_existing_meeting(doc, config, participants, meeting_url, token):
    """Update existing Teams meeting with new participants"""
    try:
        # Extract meeting ID from URL
        meeting_id = None
        if "/meetup-join/" in meeting_url:
            meeting_id = meeting_url.split("/meetup-join/")[-1].split("?")[0]
        elif "/" in meeting_url:
            meeting_id = meeting_url.split("/")[-1].split("?")[0]
        
        if not meeting_id:
            frappe.log_error(f"Could not extract meeting ID from URL: {meeting_url}", "Teams Meeting ID Error")
            frappe.throw("Could not extract meeting ID from existing meeting URL")
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # Get existing meeting details
        get_url = f"{GRAPH_API}/me/onlineMeetings/{meeting_id}"
        response = requests.get(get_url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            existing_data = response.json()
            existing_attendees = existing_data.get("participants", {}).get("attendees", [])
            
            # Get existing participant IDs
            existing_ids = set()
            for attendee in existing_attendees:
                user_id = attendee.get("identity", {}).get("user", {}).get("id")
                if user_id:
                    existing_ids.add(user_id)
            
            # Find new participants
            new_participants = []
            for participant in participants:
                user_id = participant["identity"]["user"]["id"]
                if user_id not in existing_ids:
                    new_participants.append(participant)
            
            if new_participants:
                # Update meeting with new participants
                all_participants = existing_attendees + new_participants
                
                patch_data = {
                    "participants": {
                        "attendees": all_participants
                    }
                }
                
                patch_response = requests.patch(
                    get_url, 
                    headers=headers, 
                    json=patch_data,
                    timeout=30
                )
                
                if patch_response.status_code in (200, 204):
                    return {
                        "success": True,
                        "message": f"Added {len(new_participants)} new participant(s) to the meeting."
                    }
                else:
                    frappe.log_error(f"Failed to update meeting: {patch_response.status_code} - {patch_response.text}", "Teams Meeting Update Error")
                    frappe.throw(f"Failed to update meeting participants: {patch_response.status_code}")
            else:
                return {
                    "success": True,
                    "message": "No new participants to add to the meeting."
                }
                
        elif response.status_code == 401:
            login_url = get_login_url(doc.name)
            return {"error": "auth_required", "login_url": login_url}
        else:
            frappe.log_error(f"Failed to fetch existing meeting: {response.status_code} - {response.text}", "Teams Meeting Fetch Error")
            frappe.throw(f"Failed to fetch existing meeting details: {response.status_code}")
            
    except Exception as e:
        frappe.log_error(f"Error updating existing meeting: {str(e)}", "Teams Meeting Update Error")
        frappe.throw(f"Failed to update existing meeting: {str(e)}")


def create_new_meeting(doc, config, participants, token):
    """Create a new Teams meeting"""
    try:
        # Prepare meeting times
        start_field = config["start_field"]
        end_field = config["end_field"]
        subject_field = config["subject_field"]
        
        # Get start and end times
        start_value = getattr(doc, start_field, None)
        end_value = getattr(doc, end_field, None)
        
        # Handle different field types
        if doc.doctype == "Event":
            start_dt = ensure_datetime_with_time(start_value)
            end_dt = ensure_datetime_with_time(end_value)
        elif doc.doctype == "Project":
            # For projects, use start of day and end of day
            start_dt = ensure_datetime_with_time(start_value, 9, 0)  # 9 AM
            end_dt = ensure_datetime_with_time(end_value, 17, 30)    # 5:30 PM
        else:
            start_dt = ensure_datetime_with_time(start_value)
            end_dt = ensure_datetime_with_time(end_value)
        
        # Validate times
        if not start_dt or not end_dt:
            frappe.throw("Meeting start and end times are required")
        
        if start_dt >= end_dt:
            # If end is not after start, make it 1 hour after start
            end_dt = start_dt + timedelta(hours=1)
        
        # Convert to UTC ISO format
        start_iso = to_utc_isoformat(start_dt)
        end_iso = to_utc_isoformat(end_dt)
        
        # Prepare meeting payload
        subject = getattr(doc, subject_field, f"{doc.doctype} Meeting: {doc.name}")
        
        payload = {
            "subject": subject,
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "participants": {
                "attendees": participants
            },
            "isOnlineMeeting": True,
            "allowMeetingChat": True,
            "allowTeamworkReactions": True
        }
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # Create meeting
        response = requests.post(
            f"{GRAPH_API}/me/onlineMeetings",
            headers=headers,
            json=payload,
            timeout=30
        )
        
        if response.status_code in (200, 201):
            meeting_data = response.json()
            join_url = meeting_data.get("joinUrl") or meeting_data.get("joinWebUrl")
            
            if join_url:
                # Save meeting URL to document
                doc.db_set("custom_teams_meeting_url", join_url)
                frappe.db.commit()
                
                return {
                    "success": True,
                    "message": "Teams meeting created and link saved successfully.",
                    "meeting_url": join_url
                }
            else:
                frappe.throw("Meeting created but no join URL received")
                
        elif response.status_code == 401:
            login_url = get_login_url(doc.name)
            return {"error": "auth_required", "login_url": login_url}
        else:
            error_text = response.text
            try:
                error_data = response.json()
                error_text = error_data.get("error", {}).get("message", error_text)
            except:
                pass
            
            frappe.log_error(f"Failed to create meeting: {response.status_code} - {error_text}", "Teams Meeting Creation Error")
            frappe.throw(f"Failed to create Teams meeting: {response.status_code} - {error_text}")
            
    except Exception as e:
        frappe.log_error(f"Error creating new meeting: {str(e)}", "Teams New Meeting Error")
        frappe.throw(f"Failed to create new Teams meeting: {str(e)}")


@frappe.whitelist()
def get_meeting_details(docname, doctype):
    """Get Teams meeting details for a document"""
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        
        if not meeting_url:
            return {"exists": False, "message": "No Teams meeting found for this document"}
        
        # Extract meeting ID
        meeting_id = None
        if "/meetup-join/" in meeting_url:
            meeting_id = meeting_url.split("/meetup-join/")[-1].split("?")[0]
        elif "/" in meeting_url:
            meeting_id = meeting_url.split("/")[-1].split("?")[0]
        
        if not meeting_id:
            return {"exists": True, "url": meeting_url, "message": "Meeting URL exists but ID not extractable"}
        
        # Get access token
        token = get_access_token()
        if not token:
            return {"exists": True, "url": meeting_url, "message": "Meeting exists but cannot fetch details (auth required)"}
        
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(f"{GRAPH_API}/me/onlineMeetings/{meeting_id}", headers=headers, timeout=30)
        
        if response.status_code == 200:
            meeting_data = response.json()
            return {
                "exists": True,
                "url": meeting_url,
                "details": {
                    "subject": meeting_data.get("subject"),
                    "startDateTime": meeting_data.get("startDateTime"),
                    "endDateTime": meeting_data.get("endDateTime"),
                    "participants": len(meeting_data.get("participants", {}).get("attendees", []))
                }
            }
        else:
            return {"exists": True, "url": meeting_url, "message": "Meeting URL exists but details unavailable"}
            
    except Exception as e:
        frappe.log_error(f"Error getting meeting details: {str(e)}", "Teams Meeting Details Error")
        return {"exists": False, "message": "Error fetching meeting details"}


@frappe.whitelist()
def delete_meeting(docname, doctype):
    """Delete Teams meeting for a document"""
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        
        if not meeting_url:
            return {"success": True, "message": "No meeting to delete"}
        
        # Extract meeting ID
        meeting_id = None
        if "/meetup-join/" in meeting_url:
            meeting_id = meeting_url.split("/meetup-join/")[-1].split("?")[0]
        elif "/" in meeting_url:
            meeting_id = meeting_url.split("/")[-1].split("?")[0]
        
        if not meeting_id:
            # Just clear the URL if we can't extract ID
            doc.db_set("custom_teams_meeting_url", "")
            frappe.db.commit()
            return {"success": True, "message": "Meeting URL cleared (could not extract meeting ID)"}
        
        # Get access token
        token = get_access_token()
        if not token:
            return {"error": "auth_required", "message": "Authentication required to delete meeting"}
        
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.delete(f"{GRAPH_API}/me/onlineMeetings/{meeting_id}", headers=headers, timeout=30)
        
        if response.status_code in (200, 204):
            # Clear the meeting URL from document
            doc.db_set("custom_teams_meeting_url", "")
            frappe.db.commit()
            return {"success": True, "message": "Teams meeting deleted successfully"}
        elif response.status_code == 404:
            # Meeting doesn't exist on Teams, just clear URL
            doc.db_set("custom_teams_meeting_url", "")
            frappe.db.commit()
            return {"success": True, "message": "Meeting not found on Teams, URL cleared"}
        else:
            frappe.log_error(f"Failed to delete meeting: {response.status_code} - {response.text}", "Teams Meeting Delete Error")
            return {"success": False, "message": f"Failed to delete meeting: {response.status_code}"}
            
    except Exception as e:
        frappe.log_error(f"Error deleting meeting: {str(e)}", "Teams Meeting Delete Error")
        return {"success": False, "message": f"Error deleting meeting: {str(e)}"}


@frappe.whitelist()
def reschedule_meeting(docname, doctype, new_start_time=None, new_end_time=None):
    """Reschedule an existing Teams meeting"""
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        
        if not meeting_url:
            frappe.throw("No Teams meeting found to reschedule")
        
        # Extract meeting ID
        meeting_id = None
        if "/meetup-join/" in meeting_url:
            meeting_id = meeting_url.split("/meetup-join/")[-1].split("?")[0]
        elif "/" in meeting_url:
            meeting_id = meeting_url.split("/")[-1].split("?")[0]
        
        if not meeting_id:
            frappe.throw("Could not extract meeting ID from URL")
        
        # Get access token
        token = get_access_token()
        if not token:
            return {"error": "auth_required", "login_url": get_login_url(docname)}
        
        # Use provided times or get from document
        if doctype in SUPPORTED_DOCTYPES:
            config = SUPPORTED_DOCTYPES[doctype]
            
            if not new_start_time:
                new_start_time = getattr(doc, config["start_field"], None)
            if not new_end_time:
                new_end_time = getattr(doc, config["end_field"], None)
        
        if not new_start_time or not new_end_time:
            frappe.throw("Start and end times are required for rescheduling")
        
        # Convert times
        if doctype == "Project":
            start_dt = ensure_datetime_with_time(new_start_time, 9, 0)
            end_dt = ensure_datetime_with_time(new_end_time, 17, 30)
        else:
            start_dt = ensure_datetime_with_time(new_start_time)
            end_dt = ensure_datetime_with_time(new_end_time)
        
        # Validate times
        if start_dt >= end_dt:
            end_dt = start_dt + timedelta(hours=1)
        
        start_iso = to_utc_isoformat(start_dt)
        end_iso = to_utc_isoformat(end_dt)
        
        # Update meeting
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        patch_data = {
            "startDateTime": start_iso,
            "endDateTime": end_iso
        }
        
        response = requests.patch(
            f"{GRAPH_API}/me/onlineMeetings/{meeting_id}",
            headers=headers,
            json=patch_data,
            timeout=30
        )
        
        if response.status_code in (200, 204):
            return {"success": True, "message": "Meeting rescheduled successfully"}
        else:
            frappe.log_error(f"Failed to reschedule meeting: {response.status_code} - {response.text}", "Teams Meeting Reschedule Error")
            frappe.throw(f"Failed to reschedule meeting: {response.status_code}")
            
    except Exception as e:
        frappe.log_error(f"Error rescheduling meeting: {str(e)}", "Teams Meeting Reschedule Error")
        frappe.throw(f"Failed to reschedule meeting: {str(e)}")


@frappe.whitelist()
def get_meeting_attendees(docname, doctype):
    """Get list of meeting attendees"""
    try:
        doc = frappe.get_doc(doctype, docname)
        meeting_url = doc.get("custom_teams_meeting_url")
        
        if not meeting_url:
            return {"attendees": [], "message": "No meeting found"}
        
        # Extract meeting ID
        meeting_id = None
        if "/meetup-join/" in meeting_url:
            meeting_id = meeting_url.split("/meetup-join/")[-1].split("?")[0]
        elif "/" in meeting_url:
            meeting_id = meeting_url.split("/")[-1].split("?")[0]
        
        if not meeting_id:
            return {"attendees": [], "message": "Could not extract meeting ID"}
        
        # Get access token
        token = get_access_token()
        if not token:
            return {"attendees": [], "message": "Authentication required"}
        
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(f"{GRAPH_API}/me/onlineMeetings/{meeting_id}", headers=headers, timeout=30)
        
        if response.status_code == 200:
            meeting_data = response.json()
            attendees = meeting_data.get("participants", {}).get("attendees", [])
            
            attendee_list = []
            for attendee in attendees:
                user_info = attendee.get("identity", {}).get("user", {})
                attendee_list.append({
                    "id": user_info.get("id"),
                    "displayName": user_info.get("displayName", "Unknown"),
                    "email": user_info.get("email")
                })
            
            return {
                "attendees": attendee_list,
                "count": len(attendee_list)
            }
        else:
            return {"attendees": [], "message": f"Failed to fetch attendees: {response.status_code}"}
            
    except Exception as e:
        frappe.log_error(f"Error getting meeting attendees: {str(e)}", "Teams Meeting Attendees Error")
        return {"attendees": [], "message": "Error fetching attendees"}


@frappe.whitelist()
def validate_meeting_time(start_time, end_time, timezone_str='Asia/Kolkata'):
    """Validate meeting start and end times"""
    try:
        start_dt = get_datetime(start_time)
        end_dt = get_datetime(end_time)
        
        errors = []
        
        # Check if end is after start
        if start_dt >= end_dt:
            errors.append("End time must be after start time")
        
        # Check if meeting is too long (more than 24 hours)
        duration = end_dt - start_dt
        if duration.total_seconds() > 24 * 3600:
            errors.append("Meeting duration cannot exceed 24 hours")
        
        # Check if meeting is too short (less than 15 minutes)
        if duration.total_seconds() < 15 * 60:
            errors.append("Meeting duration should be at least 15 minutes")
        
        # Check if meeting is in the past
        current_time = now_datetime()
        if start_dt < current_time:
            errors.append("Meeting cannot be scheduled in the past")
        
        return {
            "valid": len(errors) == 0,
            "errors": errors,
            "duration_hours": round(duration.total_seconds() / 3600, 2)
        }
        
    except Exception as e:
        return {
            "valid": False,
            "errors": [f"Invalid date/time format: {str(e)}"]
        }