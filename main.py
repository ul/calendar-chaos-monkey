#!/usr/bin/env python
# -*- coding: utf-8 -*-


"""Randomly decline a meeting with 4+ attendees on a week starting from tomorrow.
"""


from __future__ import print_function
import datetime
import os.path
import pytz
import random
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

### User preferences ###
MIN_ATTENDEES = 4
MUST_ATTEND = ["Debrief"]

SCOPES = ["https://www.googleapis.com/auth/calendar"]

ONE_DAY = datetime.timedelta(hours=24)


def get_service():
    """Initialize GCal API. Auth if necessary."""

    configs = os.path.dirname(os.path.realpath(__file__))
    token = configs + "/token.json"
    secrets = configs + "/client_secrets.json"

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token):
        creds = Credentials.from_authorized_user_file(token, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(secrets, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token, "w") as token:
            token.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


def datetime_to_gdate(t):
    """Format datetime to GCal API friendly string RFC3339."""

    return (
        t.astimezone(pytz.utc).replace(tzinfo=None).isoformat() + "Z"
    )  # 'Z' indicates UTC time


def gdate_to_datetime(s):
    """Parse GCal API datetime."""

    return datetime.datetime.fromisoformat(s[:-1]).replace(tzinfo=pytz.utc).astimezone()


def main():

    service = get_service()

    now = datetime.datetime.now().astimezone()
    # Do not surprise-decline today's meetings.
    start = (now + datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0)
    end = (start + datetime.timedelta(days=6)).replace(hour=23, minute=59, second=59)
    time_min = datetime_to_gdate(start)
    time_max = datetime_to_gdate(end)

    candidates = []
    page_token = None
    while True:
        events = (
            service.events()
            .list(
                calendarId="primary",
                pageToken=page_token,
                timeMin=time_min,
                timeMax=time_max,
                maxResults=2500,
                singleEvents=True,
            )
            .execute()
        )
        for event in events["items"]:
            if (
                "attendees" in event
                and len(event["attendees"]) >= MIN_ATTENDEES
                and not any(x in event["summary"] for x in MUST_ATTEND)
                and not any(
                    x.get("self") and x["responseStatus"] == "declined"
                    for x in event["attendees"]
                )
            ):
                candidates.append(event)
        page_token = events.get("nextPageToken")
        if not page_token:
            break

    winner = random.choice(candidates)

    for attendee in event["attendees"]:
        if attendee.get("self"):
            me = attendee
            me["responseStatus"] = "declined"
            me[
                "comment"
            ] = "My calendar monkey dropped this event. If you think I should attend it anyway, please ping me directly. Sorry for the inconvenience."
            service.events().patch(
                calendarId="primary",
                eventId=winner["id"],
                body={"attendees": [me]},
            ).execute()
            print(
                "Declined '{}' on {}".format(
                    winner["summary"],
                    winner["start"].get("dateTime") or winner["start"].get("date"),
                )
            )
            break


if __name__ == "__main__":
    main()
