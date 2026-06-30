# The Power People — Email Automation

## Install

1. Copy the `EmailAutomationDeployment` folder to the Windows PC.
2. Double-click **Email Automation.exe**.
3. Use **Mail Account Setup** to add Gmail accounts and Google App Passwords.
4. Use the single **Email Automation** Desktop shortcut after installation. No Python, pip, terminal, or package installation is required.

## Excel files

Place `mail_list.xlsx` in `F:\CODEX\Email_automation`.

`mail_list.xlsx` supports this structure: `First_Name`, `Last_Name`, `Email`, `Company`, `Designation`, `Country`, `Subject`, `Body`, `Sender_Name`, `Sender_Email`, `Status`, `Result`, `SentDate`.

Each row supplies its own Subject, Body, Sender_Name, and Sender_Email. Sender_Email selects the enabled Gmail account used for SMTP. App Passwords are stored only in Windows Credential Manager. If an account is missing or disabled, that row fails with `Sender Account Not Configured` and processing continues.

## Dynamic placeholders

Use any Excel column name inside double braces in Subject or Body:

```text
Subject: Opportunity for {{Company}}
Body: Hi {{First_Name}} {{Last_Name}},

{{Designation}} of {{Company}}
```

Placeholder matching is case-insensitive and spaces inside the braces are allowed. For example, `{{first_name}}`, `{{ FIRST_NAME }}`, and `{{First_Name}}` are identical. Empty cells and unknown placeholders become empty strings without stopping the send.

All current and future columns are detected automatically. Adding columns such as `Phone`, `Website`, `LinkedIn`, or `Industry` immediately makes `{{Phone}}`, `{{Website}}`, `{{LinkedIn}}`, or `{{Industry}}` available without a software update.

## Unified application

The single **Email Automation** application provides Dashboard, Send Mail Now, Daily Scheduling, Mail Account Setup, Reports, Settings, Logs, and Exit. Unlimited weekly schedules are stored independently and configured to run missed starts as soon as Windows permits.

Application data, logs, backups, reports, credentials, and configuration are installed under `%LOCALAPPDATA%\ThePowerPeople\EmailAutomation`.

## Scheduling note

Missed starts are configured to run as soon as Windows permits. A task that runs while the user is logged off requires Windows credentials and may be restricted by administrator policy; the configuration window states this limitation.

## Gmail note

Use a Google App Password, not the normal Gmail password. Gmail sending limits and anti-abuse controls still apply.
"# desktopEmailAutomationSystem" 
