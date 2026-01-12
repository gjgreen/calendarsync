import datetime
import os.path
import win32com.client
import pytz
from tzlocal import get_localzone 
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/calendar']

# Timezone Detection
try:
    LOCAL_TZ = get_localzone()
except:
    LOCAL_TZ = pytz.timezone("America/Chicago")

# Outlook Constants
olResponseAccepted = 3
olResponseOrganized = 1 

def get_google_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)

def get_calendar_id(service, calendar_name):
    calendar_list = service.calendarList().list().execute()
    for entry in calendar_list.get('items', []):
        if entry['summary'] == calendar_name:
            return entry['id']
    return 'primary'

def sync_outlook_to_google():
    service = get_google_service()
    target_cal_id = get_calendar_id(service, "Work")
    print(f"Syncing using timezone: {LOCAL_TZ}")
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    start_search = datetime.datetime.now() - datetime.timedelta(days=1)
    end_search = datetime.datetime.now() + datetime.timedelta(days=14)
    filter_str = f"[Start] >= '{start_search.strftime('%m/%d/%Y %I:%M %p')}' AND [End] <= '{end_search.strftime('%m/%d/%Y %I:%M %p')}'"
    items = items.Restrict(filter_str)

    outlook_instance_ids = []

    for appt in items:
        try:
            if appt.ResponseStatus not in [olResponseAccepted, olResponseOrganized]:
                continue

            unique_id = f"{appt.EntryID}_{appt.Start.strftime('%Y%m%dT%H%M%S')}"
            outlook_instance_ids.append(unique_id)
            
            # FIXED TIMEZONE LOGIC: Works for both pytz and ZoneInfo
            if hasattr(LOCAL_TZ, 'localize'):
                start_dt = LOCAL_TZ.localize(appt.Start.replace(tzinfo=None))
                end_dt = LOCAL_TZ.localize(appt.End.replace(tzinfo=None))
            else:
                start_dt = appt.Start.replace(tzinfo=None).replace(tzinfo=LOCAL_TZ)
                end_dt = appt.End.replace(tzinfo=None).replace(tzinfo=LOCAL_TZ)

            event_body = {
                'summary': appt.Subject,
                'location': appt.Location,
                'description': f"Outlook Sync\n\n{appt.Body[:500]}",
                'start': {'dateTime': start_dt.isoformat(), 'timeZone': str(LOCAL_TZ)},
                'end': {'dateTime': end_dt.isoformat(), 'timeZone': str(LOCAL_TZ)},
                'extendedProperties': {'private': {'outlook_instance_id': unique_id}}
            }

            existing = service.events().list(calendarId=target_cal_id, privateExtendedProperty=f"outlook_instance_id={unique_id}").execute()
            
            if not existing.get('items'):
                print(f"Adding: {appt.Subject}")
                service.events().insert(calendarId=target_cal_id, body=event_body).execute()
            else:
                print(f"Verified: {appt.Subject}")
        except Exception as e:
            print(f"Error processing event: {e}")

    # --- CLEANUP ---
    print("Checking for removals...")
    # Properly handle UTC conversion for the search window
    g_start = start_search.replace(tzinfo=None).replace(tzinfo=LOCAL_TZ).astimezone(pytz.utc).isoformat().replace('+00:00', 'Z')
    g_end = end_search.replace(tzinfo=None).replace(tzinfo=LOCAL_TZ).astimezone(pytz.utc).isoformat().replace('+00:00', 'Z')

    g_events = service.events().list(calendarId=target_cal_id, timeMin=g_start, timeMax=g_end, singleEvents=True).execute()
    
    for g_event in g_events.get('items', []):
        g_outlook_id = g_event.get('extendedProperties', {}).get('private', {}).get('outlook_instance_id')
        if g_outlook_id and g_outlook_id not in outlook_instance_ids:
            print(f"Deleting cancelled event: {g_event.get('summary')}")
            service.events().delete(calendarId=target_cal_id, eventId=g_event['id']).execute()

if __name__ == '__main__':
    sync_outlook_to_google()
