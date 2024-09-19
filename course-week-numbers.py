import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import date, datetime, timedelta

# If modifying these scopes, delete the file token.json.
SCOPES = [
	"https://www.googleapis.com/auth/calendar.events",
	"https://www.googleapis.com/auth/calendar.settings.readonly",
	"https://www.googleapis.com/auth/calendar.readonly",
	]

EVENT_COLORS = {
    '1': 'Lavender',
    '2': 'Sage',
    '3': 'Grape',
    '4': 'Flamingo',
    '5': 'Banana',
    '6': 'Tangerine',
    '7': 'Peacock',
    '8': 'Graphite',
    '9': 'Blueberry',
    '10': 'Basil',
    '11': 'Tomato'
}

def main():
	"""Shows basic usage of the Google Calendar API.
	Prints the start and name of the next 10 events on the user's calendar.
	"""
  
	# Authenticate user
	creds = authenticate()

	try:
		# Create calendar service object
		service = build("calendar", "v3", credentials=creds)

		# Get default values
		today_date = date.today().strftime("%Y-%m-%d")
		default_tz = get_user_setting(service, "timezone")
		num_events = "2"

		# Prompt user for event details
		print("***** Course Week Numbers *****")
		print("")
		user_title_prefix = input("Enter the calendar event title. Events will be created as ""{title} Week 1"" <Press Enter for none>: ")
		event_body = input("Enter a description (body) for the event: <None> ")
		event_start_date = input(f"What is the date of the event for Week 1 (YYYY-MM-DD)? <{today_date}>: ") or today_date
		num_events = input(f"How many weeks to create? <{num_events}> ") or num_events
		event_tz = input(f"What timezone to use? <{default_tz}> ") or default_tz

		# Prompt for event color
		print("\nChoose an event color:\n")
		for i, color_name in EVENT_COLORS.items():
			print(f" {i}. {color_name}")
		event_color = input(f"\nColor choice: <1> ") or "1"

		# Validate inputs
		# Date
		event_start_date = event_start_date.rstrip()
		if not validate_date_format(event_start_date):
			print(f"*** The date '{event_start_date}' was not entered in YYYY-MM-DD format!")
			return
		else:
			event_start_date_obj = datetime.strptime(event_start_date, "%Y-%m-%d")
			event_date_full = event_start_date_obj.strftime("%A, %B %d, %Y")
		
		# Number of events
		if not num_events.isdigit():
			print(f"*** The number of events '{num_events}' is not a number!")
			return
		else:
			num_events = int(num_events)

		# Color
		if event_color not in EVENT_COLORS:
			print(f"*** The color selection '{event_color}' is not valid!")
			return

		# Space title prefix if required
		event_title_prefix = ""
		if user_title_prefix != "":
			event_title_prefix = user_title_prefix.rstrip() + " "

		# Prompt for confirmation
		print("\nEvents will use the following:\n")
		print("  Title prefix:    %s" % ("<None>" if user_title_prefix == "" else user_title_prefix))
		print(f"  Description:     {event_body}")
		print(f"  Start Date:      {event_date_full}")
		print(f"  Number of Weeks: {num_events}")
		print(f"  Timezone:        {event_tz}")
		print(f"  Color:           {EVENT_COLORS.get(event_color)}")
		print("\n  First event will have the title: %s" % (event_title_prefix + "Week 1"))
		confirm = input("\nProceed? <Y> ") or "Y"
		if confirm != "Y":
			print("Aborted")
			return

		# Insert events
		for i in range(num_events):

			# Set event parameters
			event_title = event_title_prefix + "Week " + str(i + 1)
			event_date = (event_start_date_obj + timedelta(days=(i * 7))).strftime("%Y-%m-%d")

			#print(event_title, event_body, event_date, event_tz, event_color)

			# Insert event
			insert_event(service, event_date, event_title, event_body, event_tz, event_color)

	except HttpError as error:
		print(f"An error occurred: {error}")

#
# Get calendar colors
#
def get_calendar_colors(service):

	# Get available colors
	colors = service.colors().get().execute()

	return colors.get("event", {})

#
# Insert test event
#
def insert_event(service, event_date, event_title, event_body, event_tz, event_color):

	# Create event object
	event = {
		"summary": event_title,
		"description": event_body,
		"start": {
			"date": event_date,
			"timeZone": event_tz
		},
		"end": {
			"date": event_date,
			"timeZone": event_tz
		},
		"colorId": event_color,
		"reminders": {
			"useDefault": False,
			"overrides": []
		},
		"transparency": "transparent", # Sets event to Free (rather than Busy)
	}

	# Insert the event
	result = service.events().insert(calendarId="primary", body=event).execute()

	# Return event link
	return event.get("htmlLink")

#
# Authenticate
#
def authenticate():

	creds = None

  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
	if os.path.exists("token.json"):
		creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  # If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				"credentials.json", SCOPES
		)
			creds = flow.run_local_server()

	# Save the credentials for the next run
		with open("token.json", "w") as token:
			token.write(creds.to_json())
	return creds

#
# Get named user setting
#
def get_user_setting(service, setting):

	# Get named setting
	settings = service.settings().get(setting=setting).execute()

	return settings["value"]

#
# Validate date format
#
def validate_date_format(date_str):
	try:
		# Try to parse the string into a date
		datetime.strptime(date_str, '%Y-%m-%d')
		return True
	except ValueError:
		# If parsing fails, it's an invalid date format
		return False

if __name__ == "__main__":
  main()