### ERPNext & Microsoft Teams Integration

Seamlessly connect ERPNext with Microsoft Teams to enhance collaboration, streamline communication, and bring your business operations closer to your team chats.

### 🚀 Features
Sync ERPNext events, tasks, and notifications with Microsoft Teams.

Enable instant updates and alerts directly in Teams.

Improve productivity by bridging ERP data with your communication hub.


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