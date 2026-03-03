# Office 365 Calendar Cleaner

A Python script to bulk-delete all calendar events for a specific user in an Office 365 / Microsoft 365 tenant using the Microsoft Graph API.

## Use Case

Useful for IT admins who need to wipe a user's calendar — for example when offboarding staff, resetting a shared mailbox, or cleaning up test accounts.

---

## Prerequisites

- Python 3.7+
- An Azure App Registration with the correct permissions (see below)
- Admin access to your Microsoft 365 tenant

---

## Azure App Setup

1. Go to [portal.azure.com](https://portal.azure.com) and sign in
2. Search for **App registrations** → click **New registration**
3. Give it a name (e.g. `Calendar Cleaner`) → click **Register**
4. Note down your **Application (client) ID** and **Directory (tenant) ID**
5. Go to **Certificates & secrets** → **New client secret** → copy the value immediately
6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
7. Search for and add: `Calendars.ReadWrite`
8. Click **Grant admin consent**

> **Important:** This script uses **Application permissions** (not Delegated), so it can act on any user's calendar without requiring them to log in.

---

## Installation

```bash
python3 -m venv calendar-cleaner
source calendar-cleaner/bin/activate
pip install requests msal
```

---

## Configuration

Open `delete_calendar_events.py` and fill in the CONFIG section at the top:

```python
CLIENT_ID     = "YOUR_CLIENT_ID_HERE"       # Application (client) ID from Azure
TENANT_ID     = "YOUR_TENANT_ID_HERE"       # Directory (tenant) ID from Azure
CLIENT_SECRET = "YOUR_CLIENT_SECRET_HERE"   # Client secret value from Azure
TARGET_USER   = "user@yourdomain.com"       # Email of the user whose calendar to clear
```

> **Never commit your real credentials to version control.** Consider using environment variables or a `.env` file for sensitive values.

---

## Usage

```bash
python delete_calendar_events.py
```

The script will:
1. Authenticate with Microsoft Graph
2. Fetch all calendar events for the target user
3. Show you the total count
4. Ask you to type `YES` to confirm before deleting anything
5. Delete all events and print a summary

---

## Security Notes

- Add your credentials to `.gitignore` if storing them in a separate config file
- The client secret should be rotated regularly via the Azure Portal
- Limit the app registration's permissions to only what is needed

---

## License

MIT
