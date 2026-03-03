# Office 365 Calendar Cleaner

A Python script to bulk-delete calendar events from your Office 365 / Microsoft 365 calendar using the Microsoft Graph API. Supports deleting all events or filtering by a search term.

Authenticates via an interactive browser login — no admin credentials or client secrets required.

## Use Cases

- Delete all events from your calendar
- Delete only events matching a keyword (e.g. all "Zoom" meetings, all "Standup" entries)

---

## Prerequisites

- Python 3.7+
- An Azure App Registration (free, one-time setup — see below)
- A Microsoft 365 account

---

## Azure App Setup

1. Go to [portal.azure.com](https://portal.azure.com) and sign in
2. Search for **App registrations** → click **New registration**
3. Give it a name (e.g. `Calendar Cleaner`) → click **Register**
4. Note down your **Application (client) ID** and **Directory (tenant) ID**
5. Go to **Authentication** → **Add a platform** → **Mobile and desktop applications**
6. Tick the redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient` → click **Configure**
7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
8. Search for and add: `Calendars.ReadWrite`

> No client secret needed — the user logs in via their own Microsoft credentials in the browser.

---

## Installation

```bash
python3 -m venv calendar-cleaner
source calendar-cleaner/bin/activate
pip install requests msal
```

---

## Configuration

Open `delete_calendar_events.py` and fill in the CONFIG section:

```python
CLIENT_ID   = "YOUR_CLIENT_ID_HERE"     # Application (client) ID from Azure
TENANT_ID   = "YOUR_TENANT_ID_HERE"     # Directory (tenant) ID from Azure

SEARCH_TERM = ""                         # Optional: filter by keyword (e.g. "zoom")
```

### SEARCH_TERM examples

| Value | Behaviour |
|---|---|
| `""` | Deletes **all** events |
| `"zoom"` | Deletes only events with "zoom" in the subject (case-insensitive) |
| `"standup"` | Deletes only events with "standup" in the subject |

---

## Usage

```bash
python delete_calendar_events.py
```

The script will:
1. Open a browser window for Microsoft login
2. Authenticate as the logged-in user (token is cached for future runs)
3. Fetch all calendar events
4. Filter by search term (if set)
5. Show you the matching count
6. Ask you to type `YES` to confirm before deleting anything
7. Delete matching events and print a summary

---

## Security Notes

- No admin credentials or client secrets are stored — authentication is handled entirely via Microsoft's login flow
- The access token is cached locally by MSAL for convenience on repeat runs
- Only the calendar of the authenticated user is affected

---

## License

MIT
