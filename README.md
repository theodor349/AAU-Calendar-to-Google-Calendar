# AAU-Calendar-to-Google-Calendar
Retrieves calendar information about Lectors at AAU and creates an event in Google Calendar

---
**How to setup:**
 * Navigate to \AAU Calendar\AAU Calendar\bin\Debug\netcoreapp2.1 and open properties.txt
 * Here you'll have to add a calendarUrl and a calendarId 
  - The callendarUrl is the url of the callendar e.g. https://www.moodle.aau.dk/calmoodle/public/?sid=6119 
  - The calendarId is the id of the calendar you want to post the events to. Found by navigating to https://calendar.google.com and then navigating to the calendar settings and scrolling down to Calendar ID

* Also youl'll have to generate a Google OAuth 2.0 client ID which you can do here https://console.developers.google.com/apis
* Download the OAuth 2.0 client ID file and make sure to rename it to "credentials.json" 
* Then put the file at \AAU Calendar\AAU Calendar\AAU Calendar
