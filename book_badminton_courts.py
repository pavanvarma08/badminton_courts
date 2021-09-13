#!/usr/bin/env python3

"""
Book badminton courts at sportsarenan

PREREQUISTE:
    1) Set your cookies as environmental variable (COOKIES) which is used for fetching user's booked courts and to book new courts
    You can find the cookies in the network tab when you login to https://www.sportarenan.se/
    2) Store your calendar API OAUTH credentails info in credentials.json file, in order to create an event.

Usage:
    book_badminton_courts.py [--start-time <name>] [--courts <num>] [--duration <num>] [--book-courts] [--fill-courts]
                             [--email-list <str>]

Options:
    --start-time <name>  start time of courts [default: 17:00]
    --courts <num>       number of courts [default: 2]
    --duration <num>     duration of courts [default: 2]
    --book-courts        Will book courts
    --fill-courts        will fill courts
    --email-list <str>   comma separated emails, to send notifications of booked courts
"""

import requests
import html
from datetime import datetime, timedelta
import os
import json
import logging
from docopt import docopt
from bs4 import BeautifulSoup

# Google calender API
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


logging.basicConfig(format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

try:
    COOKIES = json.loads(os.environ.get("COOKIES"))
except json.decoder.JSONDecodeError as err:
    logger.error(f"Failed to parse Cookies: {err}")
    exit(1)


def get_available_slots(date: str, time_slots):
    url = "https://www.sportarenan.se/wp-content/themes/sportarenan/sportarenan-functions/wb/list_free_courts.php"

    data = {
        "containingDiv": "reserve_lanes_1",
        "activity": 2,
        "date": date,
        "searchFreeLanes": "",
        "psDate": "",
        "psActId": "",
        "psHour": "",
    }
    res = requests.post(url, data=data)
    html_doc = html.unescape(res.text)

    soup = BeautifulSoup(html_doc, "html.parser")
    available_items_html = soup.find_all("div", attrs={"class": "book-item"})
    available_time_slots = {}
    for available_item_html in available_items_html:
        time_slot = available_item_html.find("input")["name"]
        if time_slot not in time_slots:
            continue

        if "last-free-lane" in available_item_html["class"]:
            time_slot = f"{time_slot} (last)"

        if available_time_slots.get(date):
            available_time_slots[date].append(time_slot)
        else:
            available_time_slots[date] = [time_slot]

    return available_time_slots


def get_booking_times():
    logger.info("Fetching booked time slots")
    bookings_url = "https://www.sportarenan.se/min-sida/?kategori=bokningar"

    res = requests.get(bookings_url, cookies=COOKIES)

    html_doc = html.unescape(res.text)
    soup1 = BeautifulSoup(html_doc, "html.parser")
    year = datetime.now().year

    booked_slots_html = soup1.find_all("tr", attrs={"class": "values"})

    booked_times = {}
    for booked_slot_html in booked_slots_html:

        if booked_slot_html.find_all("td")[0].contents[0] == "Idag":
            date = datetime.now().strftime("%Y-%m-%d")
        else:
            date = booked_slot_html.find_all("td")[0].contents[2]
            date = f"{year}/{date}"
            date = datetime.strptime(date, "%Y/%d/%m").strftime("%Y-%m-%d")

        if not booked_times.get(date):
            booked_times[date] = {}

        start_time = booked_slot_html.find_all("td")[1].contents[0]
        duration = booked_slot_html.find_all("td")[1].contents[2][0]
        start_time_obj = datetime.strptime(start_time, "%H:%M")

        for i in range(int(duration)):
            next_start_time_obj = start_time_obj + timedelta(hours=i)
            next_start_time = next_start_time_obj.strftime("%H:%M")

            if not booked_times[date].get(next_start_time):
                booked_times[date][next_start_time] = 0

            booked_times[date][next_start_time] += 1

    logger.info(f"Booked slots: {booked_times}")
    return booked_times


def check_book_time_slots(availability, booked, num_of_courts, email_list):
    logger.info("Checking if they are any bookable slots")
    for date, time_slots in booked.items():
        bookable_time_slots = [
            key for key, val in time_slots.items() if val < num_of_courts
        ]
        if not availability.get(date):
            continue
        if any(
            time_slot in availability.get(date) for time_slot in bookable_time_slots
        ):
            book_time_slots(date, bookable_time_slots)
            generate_calendar_event(date, email_list)


def book_time_slots(date, time_slots):
    logger.info(f"Booking {date}: {time_slots}")
    if date == datetime.now().strftime("%Y-%m-%d"):
        logger.warning(f"Skipped booking slots for the same date:{date}")
        return

    url = "https://www.sportarenan.se/wp-content/themes/sportarenan/sportarenan-functions/wb/list_free_courts.php"
    data = {
        "submitBookings": "",
        "date": date.replace("-", ""),
        "act": "2",
        "act_text": "Badminton",
        "noOfLunches": "0",
        "containingDiv": "reserve_lanes_1",
    }

    for time_slot in time_slots:
        data[time_slot] = f"2-{time_slot}"

    try:
        response = requests.post(url, data=data, cookies=COOKIES)
        response.raise_for_status()
    except (requests.exceptions.HTTPError, requests.exceptions.Timeout) as e:
        logger.error(
            f"Failed to book time slots:{time_slots}. on {date} \n Error: {str(e)}"
        )


def get_available_slots_upto_given_days(days: int, time_slots):
    logger.info(f"Fetch available time slots for the next {days} days")
    today = datetime.now()
    availability = {}
    for i in range(days):
        req_date = today + timedelta(days=i)

        if req_date.weekday() > 4:
            continue

        req_date = req_date.strftime("%Y-%m-%d")
        availability.update(get_available_slots(req_date, time_slots))
    logger.info(f"Available slots: {availability}")
    return availability


def book_from_booked_times(time_slots, num_of_courts, email_list):
    booked = get_booking_times()
    availability = get_available_slots_upto_given_days(16, time_slots)
    check_book_time_slots(availability, booked, num_of_courts, email_list)


def book_courts_on_desired_date(days, time_slots, num_of_courts, email_list):
    today = datetime.now()
    req_date = today + timedelta(days=days)
    req_date = req_date.strftime("%Y-%m-%d")

    for _ in range(num_of_courts):
        book_time_slots(req_date, time_slots)

    generate_calendar_event(req_date, email_list)


def get_time_slots(start_time, duration):
    timeslots = [start_time]
    for _ in range(1, duration):
        start_time_obj = datetime.strptime(start_time, "%H:%M")
        next_time_obj = start_time_obj + timedelta(hours=1)
        next_time = next_time_obj.strftime("%H:%M")
        timeslots.append(next_time)
        start_time = next_time
    return timeslots


def get_calendar_service():
    # code from quickstart page: https://developers.google.com/calendar/api/quickstart/python

    current_dir = os.path.dirname(os.path.abspath(__file__))
    scopes = ["https://www.googleapis.com/auth/calendar"]
    credentails_file = os.path.join(current_dir, "credentials.json")

    if not credentails_file:
        logger.error(
            f"cannot create calendar event without credentials file: {credentails_file}"
        )

    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentails_file, scopes)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    service = build("calendar", "v3", credentials=creds)
    return service


def get_booked_time_slot(date):
    booked_times = get_booking_times()
    return booked_times.get(date)


def get_badminton_event(date):
    service = get_calendar_service()

    date_obj = datetime.strptime(date, "%Y-%m-%d")
    day_start_time = (
        datetime(
            date_obj.year, date_obj.month, date_obj.day, hour=0, minute=0
        ).isoformat()
        + "Z"
    )
    day_end_time = (
        datetime(
            date_obj.year, date_obj.month, date_obj.day, hour=23, minute=59
        ).isoformat()
        + "Z"
    )
    events_result = (
        service.events()
        .list(
            calendarId="primary",
            timeMin=day_start_time,
            timeMax=day_end_time,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])

    for event in events:
        if "Badminton" in event.get("summary"):
            event_id = event.get("id")
            return event_id


def update_calendar_event(service, event_id, event_body):
    event_result = (
        service.events()
        .update(
            calendarId="primary",
            eventId=event_id,
            body=event_body,
        )
        .execute()
    )
    logger.info(
        f"Updated event: {event_result['summary']}. Time: {event_result['start']['dateTime']} - {event_result['end']['dateTime']}"
    )


def create_calender_event(service, event_body):
    event_result = (
        service.events()
        .insert(
            calendarId="primary",
            body=event_body,
        )
        .execute()
    )

    logger.info(
        f"created event: {event_result['summary']}. Time: {event_result['start']['dateTime']} - {event_result['end']['dateTime']}"
    )


def generate_calendar_event(date, email_list=[]):
    service = get_calendar_service()
    time_slots = get_booked_time_slot(date)

    if date == datetime.now().strftime("%Y-%m-%d"):
        logger.warning(f"Skipped creating event for the same date:{date}")
        return

    if not time_slots:
        logger.error("Nothing booked for today, skipped creating event")
        return
    date_obj = datetime.strptime(date, "%Y-%m-%d")
    start_time_str = list(time_slots.keys())[0]
    start_time_hour = int(start_time_str.split(":")[0])
    duration = len(time_slots)
    start_time = datetime(date_obj.year, date_obj.month, date_obj.day, start_time_hour)
    end_time = start_time + timedelta(hours=duration)
    attendees = [{"email": email} for email in email_list]

    event_body = {
        "summary": f"Badminton {list(time_slots.values())} (A)",
        "description": "This is an automated event",
        "start": {
            "dateTime": start_time.isoformat(),
            "timeZone": "Europe/Stockholm",
        },
        "end": {
            "dateTime": end_time.isoformat(),
            "timeZone": "Europe/Stockholm",
        },
        "location": "Sportarenan, Bergsätersgatan 21, 421 66 Västra Frölunda, Sweden",
        "sendNotifications": "true",
        "sendUpdates": "all",
        "reminders": {
            "useDefault": "false",
            "overrides": [{"method": "popup", "minutes": 540}],
        },
        "attendees": attendees,
    }

    event_id = get_badminton_event(date)
    if event_id:
        update_calendar_event(service, event_id, event_body)
    else:
        create_calender_event(service, event_body)
    logger.info(f"Attendes: {str(email_list)}")


def setup_logger(_logger):
    _logger.setLevel(logging.INFO)

    current_dir = os.path.dirname(os.path.abspath(__file__))
    fh = logging.FileHandler(os.path.join(current_dir, "booking.log"), mode="a")
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    fh.setFormatter(formatter)
    _logger.addHandler(fh)


def main():
    args = docopt(__doc__)

    setup_logger(logger)
    start_time = args["--start-time"]
    duration = int(args["--duration"])
    courts = int(args["--courts"])
    email_list = args["--email-list"].split(",") if args["--email-list"] else []
    time_slots = get_time_slots(start_time, duration)

    if args["--book-courts"]:
        logger.info(
            f"Book courts with start time {start_time}, duration: {duration} & courts: {courts}"
        )
        days = 14  # 2 weeks
        book_courts_on_desired_date(days, time_slots, courts, email_list)

    if args["--fill-courts"]:
        logger.info(
            f"Fill courts with start time {start_time}, duration: {duration} & courts: {courts}"
        )
        book_from_booked_times(time_slots, courts, email_list)


if __name__ == "__main__":
    main()
