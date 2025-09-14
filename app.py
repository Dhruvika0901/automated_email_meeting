import streamlit as st
import pandas as pd
import base64
import uuid
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
from datetime import datetime, timedelta, timezone

# --------------------------------------------------
# Google API Setup
# --------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]
TIMEZONE = "Asia/Kolkata"


def get_google_service(api_name, api_version):
    creds = None
    if os.path.exists("token.pkl"):
        with open("token.pkl", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pkl", "wb") as token:
            pickle.dump(creds, token)
    return build(api_name, api_version, credentials=creds)


def get_sender_email():
    gmail = get_google_service("gmail", "v1")
    profile = gmail.users().getProfile(userId="me").execute()
    return profile.get("emailAddress", "me")


# --------------------------------------------------
# Calendar Event Creator
# --------------------------------------------------
def schedule_meeting(date, time_str, duration_min, attendees, summary, description=None, location=None, recurrence=None):
    service = get_google_service("calendar", "v3")

    start_dt_local = datetime.strptime(f"{date} {time_str}", "%Y-%m-%d %H:%M")
    end_dt_local = start_dt_local + timedelta(minutes=int(duration_min))

    event = {
        "summary": summary,
        "description": description or "",
        "location": location or "",
        "start": {"dateTime": start_dt_local.isoformat(), "timeZone": TIMEZONE},
        "end": {"dateTime": end_dt_local.isoformat(), "timeZone": TIMEZONE},
        "attendees": [{"email": e} for e in attendees],
        "conferenceData": {
            "createRequest": {
                "requestId": str(uuid.uuid4()),
                "conferenceSolutionKey": {"type": "hangoutsMeet"},
            }
        },
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "email", "minutes": 30},
                {"method": "popup", "minutes": 10},
            ],
        },
    }

    if recurrence:
        event["recurrence"] = [recurrence]

    created = service.events().insert(
        calendarId="primary",
        body=event,
        conferenceDataVersion=1,
        sendUpdates="all",
    ).execute()

    meet_link = ""
    try:
        entry_points = created.get("conferenceData", {}).get("entryPoints", [])
        for ep in entry_points:
            if ep.get("entryPointType") == "video":
                meet_link = ep.get("uri", "")
                break
        if not meet_link and entry_points:
            meet_link = entry_points[0].get("uri", "")
    except Exception:
        pass

    return meet_link, created.get("id", str(uuid.uuid4())), start_dt_local, end_dt_local


# --------------------------------------------------
# ICS Generator
# --------------------------------------------------
def to_utc(dt_local, tz_offset_minutes=330):
    offset = timedelta(minutes=tz_offset_minutes)
    return (dt_local - offset).replace(tzinfo=timezone.utc)


def build_ics(summary, description, location, organizer_email, attendees, dt_start_local, dt_end_local, uid=None, meet_link=None):
    uid = uid or str(uuid.uuid4())
    dtstamp_utc = datetime.utcnow().replace(tzinfo=timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    dtstart_utc = to_utc(dt_start_local).strftime("%Y%m%dT%H%M%SZ")
    dtend_utc = to_utc(dt_end_local).strftime("%Y%m%dT%H%M%SZ")

    desc_lines = [description or ""]
    if meet_link:
        desc_lines.append(f"Google Meet: {meet_link}")
    desc_lines.append(f"Local Timezone: {TIMEZONE} ({dt_start_local.strftime('%Y-%m-%d %H:%M')})")
    desc = "\\n".join(desc_lines)

    attendee_lines = [f"ATTENDEE;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:{a}" for a in attendees]

    ics = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "METHOD:REQUEST",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{dtstamp_utc}",
        f"DTSTART:{dtstart_utc}",
        f"DTEND:{dtend_utc}",
        f"SUMMARY:{summary}",
        f"DESCRIPTION:{desc}",
        f"LOCATION:{location or ''}",
        f"ORGANIZER:mailto:{organizer_email}",
        *attendee_lines,
        "END:VEVENT",
        "END:VCALENDAR",
    ]
    return "\r\n".join(ics)


# --------------------------------------------------
# Gmail Sender (HTML + ICS)
# --------------------------------------------------
def send_custom_email_with_ics(to_email, subject, html_body, ics_text):
    gmail = get_google_service("gmail", "v1")

    msg = MIMEMultipart("mixed")
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(html_body, "html"))

    ics_part = MIMEApplication(ics_text.encode("utf-8"), _subtype="ics")
    ics_part.add_header("Content-Disposition", 'attachment; filename="invite.ics"')
    ics_part.add_header("Content-Type", 'text/calendar; method=REQUEST; name="invite.ics"')
    msg.attach(ics_part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    gmail.users().messages().send(userId="me", body={"raw": raw}).execute()


# --------------------------------------------------
# Streamlit UI
# --------------------------------------------------
st.title("📅 Advanced Meeting Scheduler")

uploaded_file = st.file_uploader("Upload CSV with attendees' emails", type=["csv"])
attendees = []
if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    if "email" not in df.columns:
        st.error("CSV must have an 'email' column")
    else:
        attendees = df["email"].dropna().tolist()
        st.success(f"Loaded {len(attendees)} attendees.")

with st.form("meeting_form"):
    meeting_date = st.date_input("Meeting Date")
    meeting_time = st.time_input("Meeting Time")
    duration = st.number_input("Duration (minutes)", min_value=15, value=30, step=15)
    topic = st.text_input("Meeting Topic", "Team Meeting")
    

    recurrence = st.selectbox("Recurrence", ["None", "Daily", "Weekly", "Monthly"])
    send_custom = st.checkbox("Send custom branded email with .ics", True)

    submitted = st.form_submit_button("Schedule Meeting")

if submitted:
    if not attendees:
        st.error("No attendees uploaded.")
    else:
        recur_rule = None
        if recurrence == "Daily":
            recur_rule = "RRULE:FREQ=DAILY;COUNT=5"
        elif recurrence == "Weekly":
            recur_rule = "RRULE:FREQ=WEEKLY;COUNT=5"
        elif recurrence == "Monthly":
            recur_rule = "RRULE:FREQ=MONTHLY;COUNT=5"

        meet_link, event_id, start_dt, end_dt = schedule_meeting(
            meeting_date.strftime("%Y-%m-%d"),
            meeting_time.strftime("%H:%M"),
            duration,
            attendees,
            topic,
            
            recurrence=recur_rule,
        )

        st.success(f"Meeting Scheduled ✅\nGoogle Meet: {meet_link}")

        if send_custom:
            organizer_email = get_sender_email()
            ics_text = build_ics(topic, organizer_email, attendees, start_dt, end_dt, uid=event_id, meet_link=meet_link)

            html_template = f"""
            <html>
            <body style="font-family:Arial, sans-serif; color:#333;">
              <h2 style="color:#2d89ef;">📅 {topic}</h2>
              <p><b>Date:</b> {meeting_date}<br>
              <b>Time:</b> {meeting_time}<br>
              <b>Duration:</b> {duration} minutes<br>
              <b>Google Meet:</b> <a href="{meet_link}">{meet_link}</a></p>
              
              <hr>
              <p style="font-size:12px;color:gray;">This is an automated invite.</p>
            </body>
            </html>
            """

            for email in attendees:
                send_custom_email_with_ics(email, f"Meeting Invite: {topic}", html_template, ics_text)
                st.info(f"Custom email sent to {email}")
