# ms-teams-calendar-to-csv
Convert your MS Teams calendar events to GCal compatible CSV file

Requirements:

1. Install Python requirements via Pip - `pip3 install pytz python-dateutil httpx`
2. Find your MS Teams Calendar details by navigating to the Teams web app, inspect page, network, and look for `calendar` in the network calls. Extract the start times, the gmt offset, bearer token authentication, and cookie. Replace values in script.
3. Run script and find generated weeklyevents.csv
4. Go to Google Calendar -> Gear -> Settings -> Import/Export -> Load CSV into calendar


Events should now be front and backloaded for that current work week!