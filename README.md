# ERPNext & Microsoft Teams Integration

Seamlessly connect ERPNext with Microsoft Teams to enhance collaboration, streamline communication, and bring your business operations closer to your team chats.

## 🚀 Features

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
- Create Teams meetings directly from ERPNext records
- Automatically share meeting links with all relevant participants
- Support for recurring meetings and meeting updates
- Meeting rescheduling and participant management

### **Advanced Authentication & Security**
- OAuth 2.0 integration with Microsoft Graph API
- Secure token storage with automatic refresh handling
- Comprehensive error handling and retry mechanisms
- Rate limiting and API quota management

### **Monitoring & Analytics**
- Clear error messages for authentication, permission, or API failures
- Comprehensive server-side logging for debugging
- Usage statistics and activity monitoring
- Data export capabilities for compliance

## 📋 Prerequisites

- **ERPNext v14+** (tested with v14 and v15)
- **Microsoft 365 Business/Enterprise subscription** with Teams access
- **Azure Active Directory** tenant with app registration permissions
- **System Manager** role in ERPNext for configuration

## 📦 Installation

### Step 1: Install the App

```bash
# Navigate to your Frappe bench directory
cd $PATH_TO_YOUR_BENCH

# Get the app from repository
bench get-app https://github.com/your-repo/erpnext_teams_integration --branch master

# Install on your site
bench --site your-site-name install-app erpnext_teams_integration

# Restart your site
bench --site your-site-name migrate
bench restart
```

### Step 2: Azure App Registration

1. **Go to Azure Portal** → Azure Active Directory → App registrations
2. **Create a new registration:**
   - Name: "ERPNext Teams Integration"
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: `https://your-erpnext-site.com/api/method/erpnext_teams_integration.api.auth.callback`

3. **Configure API Permissions:**
   - Microsoft Graph API permissions:
     - `User.Read` (Delegated)
     - `User.ReadBasic.All` (Delegated)  
     - `Chat.ReadWrite` (Delegated)
     - `Chat.Create` (Delegated)
     - `ChannelMessage.Send` (Delegated)
     - `OnlineMeetings.ReadWrite` (Delegated)
   - **Grant admin consent** for your organization

4. **Create a client secret:**
   - Go to "Certificates & secrets"
   - Create a new client secret (save this value securely)

5. **Note down these values:**
   - Application (client) ID
   - Directory (tenant) ID  
   - Client secret value

### Step 3: ERPNext Configuration

1. **Go to Teams Settings** in ERPNext
2. **Fill in the Azure app details:**
   - Client ID: Your Azure app's Application ID
   - Client Secret: The secret value you created
   - Tenant ID: Your Azure directory tenant ID
   - Redirect URI: Should be auto-populated with your site URL

3. **Enable Doctypes:**
   - Add the doctypes you want to integrate (Event, Project, etc.)
   - Save the settings

4. **Authenticate with Teams:**
   - Click "Authenticate with Teams"
   - Complete the OAuth flow in the popup window
   - Verify authentication with "Test Connection"

5. **Sync User Data:**
   - Click "Sync Azure IDs" to link Frappe users with Microsoft accounts
   - This enables automatic participant detection

## 🛠️ Configuration

### Supported Doctypes

Currently supported doctypes with their participant fields:

| Doctype | Participants Field | Email Field | Subject Field |
|---------|-------------------|-------------|---------------|
| Event | event_participants | email | subject |
| Project | users | email | project_name |

### Adding Custom Doctypes

To add support for additional doctypes, modify the `SUPPORTED_DOCTYPES` configuration in:
- `erpnext_teams_integration/api/chat.py`
- `erpnext_teams_integration/api/meetings.py`

Example:
```python
SUPPORTED_DOCTYPES = {
    "Task": {
        "participants_field": "assigned_users", 
        "email_field": "user",
        "subject_field": "subject"
    }
}
```

### Required Custom Fields

The app automatically creates these custom fields:

**Event & Project:**
- `custom_teams_chat_id` - Stores the Teams chat ID
- `custom_teams_meeting_url` - Stores the Teams meeting join URL
- `custom_join_teams_meeting` - Button to open meeting

**User:**
- `azure_object_id` - Microsoft Azure user ID for API calls

## 🎯 Usage Guide

### Creating Team Chats

1. **From Event/Project form:**
   - Click "Teams" dropdown → "Create Teams Chat"
   - All participants with Microsoft accounts will be added
   - Chat ID is automatically stored in the document

2. **Programmatically:**
   ```python
   frappe.call({
       method: "erpnext_teams_integration.api.chat.create_group_chat_for_doc",
       args: { docname: "EVT-001", doctype: "Event" }
   })
   ```

### Sending Messages

1. **From ERPNext:**
   - Use "Send Teams Message" button in document
   - Messages are stored locally for history

2. **API Method:**
   ```python
   frappe.call({
       method: "erpnext_teams_integration.api.chat.send_message_to_chat",
       args: { 
           chat_id: "chat_id_here", 
           message: "Hello from ERPNext!",
           docname: "EVT-001",
           doctype: "Event"
       }
   })
   ```

### Creating Meetings

1. **From Document:**
   - Click "Create Teams Meeting" in Teams dropdown
   - Meeting is scheduled based on document dates
   - Join URL is automatically saved

2. **Meeting Management:**
   - Add/remove participants by updating document participants
   - Reschedule by updating document dates and recreating meeting
   - Use "Join Teams Meeting" button to open meeting

### Syncing Conversations

1. **Manual Sync:**
   - Use "Sync Now" button in Teams dropdown
   - Or "Sync All Conversations" in Teams Settings

2. **Automatic Sync:**
   - Set up a scheduled job in hooks.py:
   ```python
   scheduler_events = {
       "hourly": [
           "erpnext_teams_integration.api.chat.sync_all_conversations"
       ]
   }
   ```

## 🔧 API Reference

### Authentication Methods

```python
# Get authentication status
frappe.call("erpnext_teams_integration.api.auth.get_authentication_status")

# Revoke authentication
frappe.call("erpnext_teams_integration.api.auth.revoke_authentication")
```

### Chat Methods

```python
# Create group chat
frappe.call("erpnext_teams_integration.api.chat.create_group_chat_for_doc", {
    "docname": "DOC-001",
    "doctype": "Event"
})

# Send message
frappe.call("erpnext_teams_integration.api.chat.send_message_to_chat", {
    "chat_id": "19:xxx@thread.v2",
    "message": "Hello World",
    "docname": "DOC-001",
    "doctype": "Event"
})

# Get chat messages
frappe.call("erpnext_teams_integration.api.chat.get_local_chat_messages", {
    "chat_id": "19:xxx@thread.v2",
    "limit": 50
})

# Sync conversations
frappe.call("erpnext_teams_integration.api.chat.sync_all_conversations", {
    "chat_id": "19:xxx@thread.v2"  # Optional: sync specific chat
})
```

### Meeting Methods

```python
# Create meeting
frappe.call("erpnext_teams_integration.api.meetings.create_meeting", {
    "docname": "EVT-001",
    "doctype": "Event"
})

# Get meeting details
frappe.call("erpnext_teams_integration.api.meetings.get_meeting_details", {
    "docname": "EVT-001",
    "doctype": "Event"
})

# Delete meeting
frappe.call("erpnext_teams_integration.api.meetings.delete_meeting", {
    "docname": "EVT-001", 
    "doctype": "Event"
})
```

## 🛡️ Security & Permissions

### Token Security
- Access tokens are stored securely in the database
- Automatic token refresh prevents expiration issues
- Tokens are never logged or exposed in error messages

### User Permissions
- Only users with appropriate ERPNext permissions can access Teams features
- Azure AD controls which users can join chats and meetings
- All API calls are made with delegated permissions

### Data Privacy
- Chat messages are stored locally for performance and offline access
- No sensitive data is transmitted unnecessarily
- Supports data cleanup and export for compliance

## 🔍 Troubleshooting

### Common Issues

**Authentication Fails:**
- Verify Azure app permissions are granted admin consent
- Check redirect URI matches exactly (including https)
- Ensure client secret hasn't expired

**Users Not Found:**
- Run "Sync Azure IDs" to link Frappe users with Microsoft accounts
- Verify user emails match between systems
- Check user has Teams license in Microsoft 365

**API Rate Limits:**
- Microsoft Graph API has throttling limits
- Implement delays between bulk operations
- Monitor usage in Teams Settings statistics

**Chat Creation Fails:**
- Ensure all participants have Teams access
- Verify OAuth scopes include chat permissions
- Check error logs for specific API errors

### Debug Mode

Enable debug logging by adding to your site config:

```python
# In sites/[site]/site_config.json
{
    "developer_mode": 1,
    "log_level": "DEBUG"
}
```

### Log Files

Check these log files for debugging:
- `logs/web.log` - General application logs
- `logs/worker.log` - Background job logs  
- Error Log doctype in ERPNext - Teams-specific errors

### Getting Help

1. **Check Error Logs** in ERPNext first
2. **Verify Configuration** using "Validate Configuration" button
3. **Test Connection** to confirm API access
4. **Review Microsoft Graph API documentation** for permission issues

## 🚀 Performance Optimization

### Database Optimization
- Indexes are automatically created for faster queries
- Regular cleanup of old messages recommended
- Consider archiving old conversation data

### API Optimization
- Implement caching for frequently accessed data
- Use batch operations where possible
- Monitor rate limits and implement backoff

### Best Practices
- Limit message history sync to recent data
- Use specific chat sync instead of bulk sync when possible
- Regular maintenance of Azure AD user mappings

## 🤝 Contributing

We welcome contributions from developers of all skill levels!

### Development Setup

1. **Fork the repository** and clone locally
2. **Set up development environment:**
   ```bash
   cd apps/erpnext_teams_integration
   pip install -e .
   pre-commit install
   ```

3. **Make changes** and test thoroughly
4. **Run quality checks:**
   ```bash
   pre-commit run --all-files
   ```

### Code Quality Standards

- **Python**: Follow PEP 8, use type hints
- **JavaScript**: ES6+, consistent formatting with Prettier
- **Documentation**: Update README and docstrings
- **Testing**: Add unit tests for new features

### Submitting Changes

1. **Create feature branch:** `git checkout -b feature/amazing-feature`
2. **Commit changes:** `git commit -m "Add amazing feature"`
3. **Push to branch:** `git push origin feature/amazing-feature`
4. **Open Pull Request** with detailed description

## 🧪 Testing

### Automated Testing

```bash
# Run all tests
bench --site test_site run-tests --app erpnext_teams_integration

# Run specific test file
bench --site test_site run-tests --app erpnext_teams_integration --module path.to.test
```

### Manual Testing Checklist

- [ ] Authentication flow works end-to-end
- [ ] Chat creation with various participant combinations
- [ ] Message sending and receiving
- [ ] Meeting creation and management
- [ ] Error handling for network issues
- [ ] Token refresh on expiration

## 📊 Monitoring & Analytics

### Usage Statistics
Access comprehensive statistics in Teams Settings:
- Total messages sent/received
- Active conversations count
- User engagement metrics
- API usage patterns

### Health Monitoring
- Token expiration alerts
- Failed API call tracking
- Performance metrics
- Error rate monitoring

## 🔄 Maintenance

### Regular Tasks
1. **Monitor token expiration** - tokens refresh automatically but watch for issues
2. **Clean up old messages** - use cleanup function to manage database size
3. **Review error logs** - identify patterns and optimize accordingly
4. **Update user mappings** - sync Azure IDs when users are added/changed

### Version Updates
1. **Backup data** before updating
2. **Test in staging** environment first
3. **Review breaking changes** in release notes
4. **Update API permissions** if new scopes are required

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](license.txt) file for details.

## 💬 Support & Community

- **Documentation**: This README and inline code documentation
- **Issues**: Use GitHub Issues for bug reports and feature requests
- **Discussions**: GitHub Discussions for questions and community support
- **Microsoft Graph**: [Official Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)

## 🙏 Acknowledgments

- **Frappe Framework** team for the excellent foundation
- **Microsoft Graph API** for comprehensive Teams integration
- **ERPNext Community** for feedback and contributions
- **Open Source Contributors** who make this project better

---

**Made with ❤️ for the ERPNext community**

*Let's build something awesome together! ✨*