# ERPNext & Microsoft Teams Integration

Seamlessly connect ERPNext with Microsoft Teams to enhance collaboration, streamline communication, and bring your business operations closer to your team chats.

---

## Features

### **Sync ERPNext Data with Teams**
- Connect supported ERPNext doctypes (Events, Projects, etc.) directly to Teams
- Keep participants synchronized between ERPNext and Teams automatically
- Real-time bidirectional data synchronization

### **Create Teams Group Chats from ERPNext**
- Instantly create Teams group chats for any supported doctype with its participants
- Automatically add new members when they are added later in ERPNext
- Prevent duplicate chats by reusing existing chat IDs when possible
- Smart participant management with Azure AD integration

### **Send & Receive Messages**
- Post messages directly to Teams chats from ERPNext
- Post messages to specific Teams channels from ERPNext
- View all chat history inside ERPNext with inbound & outbound messages
- HTML message formatting support with XSS protection

### **Two-Way Message Sync**
- Fetch and store recent Teams chat messages into ERPNext
- Maintain a searchable log of all Teams communications
- Automatic message deduplication
- Performance-optimized message storage

### **Microsoft Teams Meeting Creation**
- Create Teams meetings directly from ERPNext records (Events & Projects)
- Automatically share meeting links with all relevant participants
- Reschedule meetings by updating document dates
- Add/remove participants from existing meetings
- "Join Teams Meeting" button auto-appears on the form when a meeting exists
- Support for meeting deletion and attendee management

### **Calendar Event Creation & Blocking**
- Creating a Teams meeting from an Event automatically blocks the time slot on Microsoft Outlook/Teams calendar
- Meeting start/end times are derived directly from the ERPNext Event's `starts_on` and `ends_on` fields
- Project meetings default to business hours (9:00 AM–5:30 PM) when only dates are provided
- Graceful fallback: if no time is set on the event, a sensible default is applied rather than failing
- Meeting times are converted to UTC before being sent to Microsoft Graph, ensuring correct timezone handling across regions
- Attendees are resolved from event participants via their Azure AD Object IDs

### **Advanced Authentication & Security**
- OAuth 2.0 integration with Microsoft Graph API
- Secure token storage with automatic refresh handling (5-minute safety buffer before expiry)
- Comprehensive error handling and retry mechanisms with token refresh on 401 responses
- Tenant ID and redirect URI format validation in the settings form

### **Monitoring & Analytics**
- Clear error messages for authentication, permission, or API failures
- Comprehensive server-side logging with safe title truncation (Frappe 140-char limit handled)
- Usage statistics dashboard: total messages, unique chats, user engagement
- Recent activity breakdown (last 7 days)
- Data export capabilities (JSON & CSV) for compliance
- Bulk message cleanup with configurable retention period

---

## Prerequisites

- **ERPNext v14+** (tested with v14 and v15)
- **Microsoft 365 Business/Enterprise subscription** with Teams access
- **Azure Active Directory** tenant with app registration permissions
- **System Manager** role in ERPNext for configuration
- **`pytz`** Python package (for timezone-aware meeting time conversion)

---

## Installation

### Step 1: Install the App

```bash
cd $PATH_TO_YOUR_BENCH
bench get-app https://github.com/your-repo/erpnext_teams_integration --branch master
bench --site your-site-name install-app erpnext_teams_integration
bench --site your-site-name migrate
bench restart
```

After install, the app automatically:
- Creates the `azure_object_id` custom field on the `User` doctype
- Creates the `Teams Settings` singleton with a pre-populated redirect URI
- Sets up default permissions for all Teams doctypes
- Creates database indexes for optimal query performance

### Step 2: Azure App Registration

1. **Go to Azure Portal** → Azure Active Directory → App registrations
2. **Create a new registration:**
   - Name: "ERPNext Teams Integration"
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: `https://your-erpnext-site.com/api/method/erpnext_teams_integration.api.auth.callback`

3. **Configure API Permissions** (Microsoft Graph):

   | Permission | Type |
   |---|---|
   | `User.Read` | Delegated |
   | `User.ReadBasic.All` | Delegated |
   | `Chat.ReadWrite` | Delegated |
   | `Chat.Create` | Delegated |
   | `ChannelMessage.Send` | Delegated |
   | `OnlineMeetings.ReadWrite` | Delegated |
   | `offline_access` | Delegated |
   | `Calendar.ReadWrite` | Delegated |

   **Grant admin consent** for your organization.

4. **Create a client secret** under "Certificates & secrets" and save the value securely.

5. **Note down:**
   - Application (client) ID
   - Directory (tenant) ID
   - Client secret value

### Step 3: ERPNext Configuration

1. Go to **Teams Settings** in ERPNext
2. Fill in your Azure app credentials (Client ID, Client Secret, Tenant ID)
3. Verify the Redirect URI is auto-populated correctly
4. Click **"Authenticate with Teams"** and complete the OAuth flow
5. Click **"Test Connection"** to verify chat and meetings access
6. Click **"Sync Azure IDs"** to link Frappe users with their Microsoft accounts

---

## Configuration

### Teams Settings Fields

| Field | Description |
|---|---|
| Client ID | Azure app's Application (client) ID |
| Client Secret | Azure app client secret value |
| Tenant ID | Azure Directory (tenant) ID — must be a valid GUID |
| Redirect URI | Auto-populated; must match exactly what is set in Azure |
| Access Token | Managed automatically; do not edit manually |
| Refresh Token | Managed automatically; do not edit manually |
| Token Expiry | Managed automatically; tokens refresh 5 minutes before expiry |
| Azure Owner Email ID | Email of the authenticated Microsoft account |
| Owner Azure Object ID | Read-only; auto-populated on authentication |
| Enabled Doctypes | Child table of doctypes to enable for Teams integration |

### Supported Doctypes

| Doctype | Participants Field | Email Field | Subject Field | Start Field | End Field |
|---|---|---|---|---|---|
| Event | `event_participants` | `email` | `subject` | `starts_on` | `ends_on` |
| Project | `users` | `email` | `project_name` | `expected_start_date` | `expected_end_date` |

### Adding Custom Doctypes

Modify `SUPPORTED_DOCTYPES` in both `api/chat.py` and `api/meetings.py`:

```python
SUPPORTED_DOCTYPES = {
    "Task": {
        "participants_field": "assigned_users",
        "email_field": "user",
        "subject_field": "subject",
        "start_field": "exp_start_date",
        "end_field": "exp_end_date"
    }
}
```

Then add the corresponding frontend button group in a new `public/js/task_teams_chat.js` and register it in `hooks.py`:

```python
doctype_js = {
    "Project": "public/js/project_teams_chat.js",
    "Event": "public/js/event_teams_chat.js",
    "Task": "public/js/task_teams_chat.js"
}
```

### Custom Fields Added to Doctypes

**Event & Project:**

| Field | Type | Description |
|---|---|---|
| `custom_teams_chat_id` | Data (Read Only) | Linked Teams group chat ID |
| `custom_teams_meeting_url` | Small Text (Read Only) | Teams meeting join URL |
| `custom_join_teams_meeting` | Button (conditional) | Opens meeting URL in new tab; visible only when meeting URL is set |
| `custom_outlook_event_id` | Small Text (Read Only) | Outlook Event ID to identify the event |

**User:**

| Field | Type | Description |
|---|---|---|
| `azure_object_id` | Data (Read Only) | Microsoft Azure AD Object ID |

---

## Usage Guide

### Teams Dropdown (available on Event & Project forms)

| Button | Action |
|---|---|
| Create Teams Chat | Creates a new group chat with all document participants |
| Open Teams Chat | Shows local message history in a modal |
| Send Teams Message | Prompts for a message and sends it to the linked chat |
| Post to Channel | Posts a message to a specific Team/Channel by ID |
| Create Teams Meeting | Creates an online meeting and saves the join URL on the document |
| Sync Now | Fetches latest messages from the linked chat |

### Meeting Creation & Calendar Blocking

When "Create Teams Meeting" is clicked on an Event:

1. The app resolves all `event_participants` to their Azure Object IDs
2. A meeting is created via Microsoft Graph with the event's `starts_on`/`ends_on` as the schedule
3. The meeting is added to the organizer's Outlook/Teams calendar, blocking the time slot
4. All attendees receive a calendar invite through Teams
5. The `custom_teams_meeting_url` field is populated with the join URL
6. The "Join Teams Meeting" button becomes visible on the form

**Time handling logic:**
- If `starts_on`/`ends_on` have no time component (midnight), a sensible default is applied: 9:00 AM start for Projects, and current time for Events
- All times are converted from `Asia/Kolkata` (IST) to UTC before being sent to the API
- If end time is missing or earlier than start, it defaults to `start + 1 hour`

### Syncing Conversations

**Manual (per document):** Teams → Sync Now  
**Manual (all chats):** Teams Settings → Sync Actions → Sync All Conversations  
**Automatic (scheduled):** Configured in `hooks.py` — runs hourly by default:

```python
scheduler_events = {
    "hourly": [
        "erpnext_teams_integration.api.chat.sync_all_conversations"
    ]
}
```

---

## 🔧 API Reference

### Authentication (`api/auth.py`)

| Method | Description |
|---|---|
| `erpnext_teams_integration.api.auth.callback` | OAuth callback handler (guest-accessible) |
| `erpnext_teams_integration.api.auth.get_authentication_status` | Check current token validity |
| `erpnext_teams_integration.api.auth.revoke_authentication` | Clear all tokens |

### Chat (`api/chat.py`)

| Method | Key Args | Description |
|---|---|---|
| `create_group_chat_for_doc` | `docname`, `doctype` | Create or update group chat |
| `send_message_to_chat` | `chat_id`, `message`, `docname`, `doctype` | Send a message |
| `get_local_chat_messages` | `chat_id`, `limit` | Fetch stored messages (max 500) |
| `fetch_and_store_chat_messages` | `chat_id`, `top` | Pull from Graph API and store locally |
| `post_message_to_channel` | `team_id`, `channel_id`, `message` | Post to a Teams channel |
| `sync_all_conversations` | `chat_id` (optional) | Sync one or all conversations |
| `get_chat_statistics` | `chat_id` (optional) | Message counts and stats |

### Meetings (`api/meetings.py`)

| Method | Key Args | Description |
|---|---|---|
| `create_meeting` | `docname`, `doctype` | Create meeting or update attendees |
| `get_meeting_details` | `docname`, `doctype` | Fetch meeting info from Graph |
| `delete_meeting` | `docname`, `doctype` | Delete meeting and clear URL |
| `reschedule_meeting` | `docname`, `doctype`, `new_start_time`, `new_end_time` | Update meeting schedule |
| `get_meeting_attendees` | `docname`, `doctype` | List current attendees |
| `validate_meeting_time` | `start_time`, `end_time`, `timezone_str` | Validate time range |

### Settings (`api/settings.py`)

| Method | Description |
|---|---|
| `bulk_sync_azure_ids` | Pull all Azure AD users and update local `azure_object_id` |
| `test_teams_connection` | Verify API access and permission scopes |
| `get_teams_statistics` | Usage stats + recent activity + top chats |
| `cleanup_old_messages` | Delete messages older than N days |
| `export_chat_history` | Export messages as JSON or CSV |
| `validate_configuration` | Check all required settings and token validity |
| `reset_integration` | Clear all tokens and auth data |

### Helpers (`api/helpers.py`)

| Method | Description |
|---|---|
| `get_access_token` | Returns valid token, auto-refreshes if near expiry |
| `refresh_access_token` | Manually trigger a token refresh |
| `get_azure_user_id_by_email` | Resolve email → Azure Object ID (with local caching) |
| `get_login_url` | Generate the Microsoft OAuth authorization URL |
| `validate_settings` | Validate GUID format, redirect URI, required fields |
| `test_api_connection` | Quick `/me` call to verify token |

---

## File Structure

```
erpnext_teams_integration/
├── hooks.py                              # App hooks, scheduled jobs, doctype JS
├── install.py                            # Post-install setup & pre-uninstall cleanup
├── modules.txt
├── api/
│   ├── auth.py                           # OAuth callback & token management
│   ├── chat.py                           # Group chat, messaging, sync
│   ├── meetings.py                       # Meeting CRUD, scheduling, attendees
│   ├── helpers.py                        # Token utils, Azure ID resolution
│   └── settings.py                       # Bulk sync, stats, export, cleanup
├── erpnext_teams_integration/
│   ├── custom/
│   │   ├── event.json                    # Custom fields for Event doctype
│   │   └── project.json                  # Custom fields for Project doctype
│   └── doctype/
│       ├── teams_settings/               # Singleton settings DocType
│       ├── teams_conversation/           # Tracks linked chat per document
│       ├── teams_chat_message/           # Stores all inbound/outbound messages
│       └── teams_enabled_doctype/        # Child table for enabled doctypes
└── public/
    └── js/
        ├── event_teams_chat.js           # Teams dropdown for Event form
        └── project_teams_chat.js         # Teams dropdown for Project form
```

---

## Security & Permissions

- Access tokens and refresh tokens are stored in the `Teams Settings` singleton — accessible only to System Manager
- Tokens are never logged or included in error messages
- Automatic refresh prevents stale token issues; if the refresh token itself is invalid, all tokens are cleared and re-auth is required
- `sanitize_html` is applied to all inbound message bodies before storage
- Outbound messages are `html.escape()`d before being sent to the Graph API
- All API routes are `@frappe.whitelist()` — enforcing standard Frappe session authentication

---

## Troubleshooting

**Authentication fails:**
- Ensure Azure app permissions have admin consent granted
- Check that the Redirect URI in Azure matches exactly (including protocol)
- Verify the client secret hasn't expired in Azure Portal

**Meeting times are wrong:**
- Check that `starts_on`/`ends_on` are set on the Event with a time, not just a date
- The app assumes `Asia/Kolkata` (IST) as the local timezone — if your server is in a different timezone, update `timezone_str` in `to_utc_isoformat()`

**Users not found / added to chat:**
- Run "Sync Azure IDs" from Teams Settings to populate `azure_object_id` on all users
- Verify user emails match exactly between ERPNext and Microsoft 365
- Confirm users have active Teams licenses

**API rate limits:**
- Microsoft Graph throttles heavily under bulk operations — stagger sync jobs
- The hourly scheduler is conservative by design; increase frequency cautiously

### Debug Logging

```json
// sites/[site]/site_config.json
{
    "developer_mode": 1,
    "log_level": "DEBUG"
}
```

Check: `logs/web.log`, `logs/worker.log`, and the **Error Log** DocType in ERPNext.

---

## Performance Notes

- Database indexes are auto-created on install for `chat_id`, `message_id`, `direction`, and `azure_object_id`
- `message_id` has a unique index — duplicate messages are silently skipped
- `get_local_chat_messages` caps results at 500; use `limit` parameter to paginate
- Azure Object IDs are cached in `User.azure_object_id` to avoid redundant Graph API calls
- Use `cleanup_old_messages(days=30)` periodically to prevent unbounded table growth

---

## Testing

```bash
# Run all app tests
bench --site test_site run-tests --app erpnext_teams_integration

# Run specific module
bench --site test_site run-tests --app erpnext_teams_integration --module erpnext_teams_integration.api.chat
```

**Manual checklist:**
- [ ] Full OAuth flow (authenticate → test connection → revoke → re-authenticate)
- [ ] Chat creation with 1, 2, and 5+ participants
- [ ] Adding a participant to an existing chat
- [ ] Message send/receive and local storage
- [ ] Meeting creation from Event (with time set)
- [ ] Meeting creation from Event (without time — fallback behavior)
- [ ] Calendar block visible in Outlook/Teams after meeting creation
- [ ] Meeting reschedule reflects updated time in Teams calendar
- [ ] Meeting deletion clears URL from document
- [ ] Token auto-refresh when near expiry
- [ ] `sync_all_conversations` scheduler job

---

## Maintenance

1. **Monitor token expiry** — the refresh is automatic, but watch for expired client secrets in Azure Portal (they have a max lifetime of 24 months)
2. **Cleanup messages** — use the "Cleanup Old Messages" button in Teams Settings periodically
3. **Re-sync Azure IDs** — run "Sync Azure IDs" whenever new users are onboarded or emails change
4. **Check error logs weekly** — especially after Microsoft Graph API changes

---

## Contributing

1. Fork and clone the repo
2. `pip install -e .` in the app directory
3. Follow PEP 8 for Python; ES6+ for JavaScript
4. Add docstrings to new API methods
5. Open a PR with a clear description of what changed and why

---

## License

MIT — see [license.txt](license.txt) for details.

---