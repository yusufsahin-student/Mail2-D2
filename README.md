# FRC Sponsorship Email Bot

Reads each row in an Excel or CSV file as a recipient, fills the `(column name)` fields in the template, and sends emails through SMTP in a controlled manner.

## Quick Start (Windows)

1. Double-click the `run.bat` file in this folder.
2. On the first launch, the application creates its own virtual environment and installs the `openpyxl` package.
3. Select your Excel/CSV file, choose the worksheet and the email column.
4. Edit the subject and message; use Excel column names in the format `(name)`, `(surname)`, `(company)`.
5. Enter your SMTP settings.
6. Create a preview, fix problematic rows, and send a test email to your own address first.
7. If the test is successful, start the bulk sending process.

If you want to run it from the command line:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

## Excel Format

The first row must contain the column headers. Example:

| name | surname | company | mission | email |
|---|---|---|---|---|
| John | Doe | Example Inc. | Bringing young people together with technology | john@example.com |

- Each filled row represents one recipient.
- Header matching is tolerant of uppercase/lowercase differences and Turkish character variations.
- To display actual parentheses in the template, use double parentheses: `((not))` → `(not)`.
- If a field used in the template is empty or does not exist in Excel, that row will not be sent.
- Invalid and duplicate email addresses are skipped by default.

## SMTP Examples

| Provider | Server | Port | Security |
|---|---|---:|---|
| Gmail | smtp.gmail.com | 587 | STARTTLS |
| Microsoft 365 | smtp.office365.com | 587 | STARTTLS |

Your organization's administrator may need to enable SMTP AUTH access. Gmail and some other providers require an app password instead of your normal account password. Official resources: [Google Workspace SMTP settings](https://support.google.com/a/answer/176600) and [Microsoft 365 client SMTP submission](https://learn.microsoft.com/en-gb/Exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365).

This version uses SMTP username/password authentication. Microsoft plans to disable this method by default in Exchange Online by the end of December 2026; if you use a Microsoft 365 account, plan the transition to OAuth with your administrator. Current timeline: [Microsoft announcement](https://techcommunity.microsoft.com/blog/exchange/updated-exchange-online-smtp-auth-basic-authentication-deprecation-timeline/4489835).

### Turkish Character / ASCII Error

The application supports Unicode usernames and passwords. If your provider still rejects the login:

- Enter your full email address in the **Username** field, not your display name.
- If you use Gmail, use an app password instead of your normal account password.
- For Microsoft 365 accounts, consult your administrator regarding SMTP AUTH or OAuth requirements.

## Security and Sending Control

- Passwords are not saved to any file.
- A confirmation window appears before real bulk sending.
- Between each email, a new random delay is selected between the minimum and maximum limits you specify. The range can be `0.5–3600` seconds, and the process can be paused while waiting.
- Results are saved as CSV files in the `mail_bot_logs` folder next to your selected Excel file.
- Email content and passwords are not stored in the logs.
- Follow your email provider's daily sending limits and personal data/anti-spam communication rules.

## Supported File Types

- `.xlsx`
- `.xlsm` (data reading only)
- `.csv` (UTF-8 or Windows Turkish encoding)

The old `.xls` format is not supported; save it as `.xlsx` in Excel.