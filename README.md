# Outlook to Google Calendar Sync

A Python-based bridge that syncs accepted Outlook meetings from a work laptop to a specific Google Calendar ("Work").

## Features
- **Smart Sync:** Only syncs meetings you have accepted or organized.
- **Auto-Cleanup:** Removes meetings from Google if they are cancelled or deleted in Outlook.
- **Timezone Aware:** Automatically detects system timezone and handles Pacific/Chicago/etc.
- **Recurring Support:** Correctly expands recurring series into individual instances.

## Initial Setup
1. **Google Cloud Console:**
   - Create a project.
   - Enable "Google Calendar API".
   - Create "OAuth Client ID" (Desktop App).
   - Download the JSON file, rename it to `credentials.json`, and place it in this folder.
2. **Google Calendar:**
   - Create a new calendar in your Gmail account named **Work**.
3. **Libraries:**
   - Run `pip install -r requirements.txt` or `py -m pip install -r requirements.txt`.

## How to Run
- **Manual:** Run `python calendarsync.py` or `py .\calendarsync.py`.
- **Automatic:** Use the provided `run_sync.bat` with Windows Task Scheduler.

## Troubleshooting
- **Logs:** Check `sync_log.txt` for errors.
- **Auth:** If sync stops, delete `token.json` and run manually to re-authenticate.
- **Outlook:** Outlook must be open/running for the sync to access the MAPI interface.

