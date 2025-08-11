### ERPNext & Microsoft Teams Integration

Seamlessly connect ERPNext with Microsoft Teams to enhance collaboration, streamline communication, and bring your business operations closer to your team chats.

### 🚀 Features
With the ERPNext + Microsoft Teams Integration App, you can:

Sync ERPNext Data with Teams

Connect supported ERPNext doctypes (e.g., Events, Projects) directly to Teams.

Keep participants in sync between ERPNext and Teams automatically.

Create Teams Group Chats from ERPNext

Instantly create Teams group chats for any supported doctype with its participants.

Automatically add new members if they are added later in ERPNext.

Prevent duplicate chats by reusing existing chat IDs when possible.

Send & Receive Messages

Post messages directly to a Teams chat from ERPNext.

Post messages to a specific Teams channel from ERPNext.

View all chat history inside ERPNext with inbound & outbound messages stored in Teams Chat Message doctype.

Two-Way Message Sync

Fetch and store recent Teams chat messages into ERPNext.

Maintain a searchable log of all Teams communications for linked ERPNext records.

Conversation Management

Sync all Teams conversations or a specific one from within ERPNext.

Store conversations in ERPNext for reporting, auditing, and historical context.

Microsoft Teams Meeting Creation

Create Teams meetings directly from ERPNext records (Events, Projects, or other supported doctypes).

Automatically share meeting links with all relevant participants.

Authentication & Integration

OAuth 2.0 integration with Microsoft Graph API.

Secure token storage and automatic refresh handling.

Error Handling & Logging

Clear error messages for authentication, permission, or API failures.

Server-side logging of failed requests for debugging.


### 📦 Installation
Install the app using the bench CLI:

cd $PATH_TO_YOUR_BENCH
bench get-app $URL_OF_THIS_REPO --branch master
bench install-app erpnext_teams_integration


### 🤝 Contributing
We welcome contributions from developers of all skill levels! Whether you’ve found a bug, want to add a new feature, or improve documentation — we’d love to have your input.

Here’s how you can get started:

Fork the repo and create your branch:

git checkout -b feature/amazing-feature
Install development tools (we use pre-commit to maintain code quality):

cd apps/erpnext_teams_integration
pre-commit install
Commit with style:
Pre-commit runs the following tools before each commit:

ruff — Python linter

eslint — JavaScript linter

prettier — Code formatter

pyupgrade — Python syntax upgrades

Push to your branch and open a Pull Request.
Be descriptive — tell us what problem you’re solving and how you tested it.


### 🛠 Continuous Integration (CI)
This repository is equipped with GitHub Actions:

CI Workflow — Installs the app and runs unit tests on every push to master.

Linters — Runs:

Frappe Semgrep Rules

pip-audit
to check dependencies for known vulnerabilities.


### 📄 License
This project is licensed under the MIT License.


### 💬 Get Involved
We believe open-source thrives when people collaborate.
If you’ve got ideas, feedback, or just want to say hi — open an issue or start a discussion.

Let’s build something awesome together! ✨