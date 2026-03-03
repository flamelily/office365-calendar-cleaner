"""
Office 365 Calendar Cleaner - Microsoft Graph API
Deletes calendar events for a specific user in your tenant.
Optionally filter by a search term to only delete matching events.

SETUP:
1. pip install requests msal
2. In Azure Portal, ensure your app has:
   - API Permission: Microsoft Graph > Application > Calendars.ReadWrite
   - Admin consent granted
3. Fill in your credentials below
4. Run: python delete_calendar_events.py
"""

import requests
import msal
import sys

# ============================================================
# CONFIG - Fill these in with your Azure app details
# ============================================================
CLIENT_ID     = "YOUR_CLIENT_ID_HERE"
TENANT_ID     = "YOUR_TENANT_ID_HERE"
CLIENT_SECRET = "YOUR_CLIENT_SECRET_HERE"
TARGET_USER   = "user@yourdomain.com"   # Email of the user whose calendar to clear

# Optional: only delete events whose subject contains this word/phrase
# Leave as empty string "" to delete ALL events
SEARCH_TERM   = ""   # e.g. "zoom" or "standup" or ""
# ============================================================

AUTHORITY  = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES     = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        print("Authentication failed:")
        print(result.get("error_description", result))
        sys.exit(1)
    print("Authenticated successfully")
    return result["access_token"]


def get_all_events(headers):
    events = []
    url = f"{GRAPH_BASE}/users/{TARGET_USER}/events?$select=id,subject&$top=100"
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 404:
            print(f"User '{TARGET_USER}' not found. Check the email address.")
            sys.exit(1)
        response.raise_for_status()
        data = response.json()
        events.extend(data.get("value", []))
        print(f"  Fetched {len(events)} events so far...", end="\r")
        url = data.get("@odata.nextLink")
    print(f"\nTotal events found: {len(events)}")
    return events


def filter_events(events, search_term):
    if not search_term:
        return events
    term = search_term.lower()
    matched = [e for e in events if term in e.get("subject", "").lower()]
    print(f"Events matching \"{search_term}\": {len(matched)}")
    return matched


def delete_events(headers, events):
    total = len(events)
    deleted = 0
    failed = 0
    for i, event in enumerate(events, 1):
        event_id = event["id"]
        subject = event.get("subject", "(no title)")
        url = f"{GRAPH_BASE}/users/{TARGET_USER}/events/{event_id}"
        response = requests.delete(url, headers=headers)
        if response.status_code == 204:
            deleted += 1
            print(f"  [{i}/{total}] Deleted: {subject[:60]}")
        else:
            failed += 1
            print(f"  [{i}/{total}] Failed ({response.status_code}): {subject[:60]}")
    return deleted, failed


def main():
    print("=" * 55)
    print("  Office 365 Calendar Cleaner")
    print("=" * 55)

    if "YOUR_" in CLIENT_ID or "YOUR_" in TENANT_ID or "YOUR_" in CLIENT_SECRET:
        print("Please fill in CLIENT_ID, TENANT_ID, and CLIENT_SECRET in the CONFIG section.")
        sys.exit(1)

    if TARGET_USER == "user@yourdomain.com":
        print("Please set TARGET_USER to the email address of the user.")
        sys.exit(1)

    print(f"\nTarget user: {TARGET_USER}")
    if SEARCH_TERM:
        print(f"Filter:      Only events containing \"{SEARCH_TERM}\"")
    else:
        print(f"Filter:      None (all events will be deleted)")

    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    print("\nFetching all calendar events...")
    all_events = get_all_events(headers)

    events = filter_events(all_events, SEARCH_TERM)

    if not events:
        if SEARCH_TERM:
            print(f"No events found matching \"{SEARCH_TERM}\".")
        else:
            print("No events found - calendar is already empty!")
        return

    if SEARCH_TERM:
        print(f"\nWARNING: This will permanently delete {len(events)} events matching \"{SEARCH_TERM}\" for {TARGET_USER}.")
    else:
        print(f"\nWARNING: This will permanently delete ALL {len(events)} events for {TARGET_USER}.")

    confirm = input("Type YES to confirm: ").strip()
    if confirm != "YES":
        print("Aborted. No events were deleted.")
        return

    print(f"\nDeleting {len(events)} events...\n")
    deleted, failed = delete_events(headers, events)

    print("\n" + "=" * 55)
    print(f"  Deleted: {deleted}")
    if failed:
        print(f"  Failed:  {failed}")
    print("  Done!")
    print("=" * 55)


if __name__ == "__main__":
    main()
