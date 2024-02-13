import unittest
from w32obj import W32Event
import icalendar
import datetime
import w32a_cal
import pytz

class ConvertWin32ToIcalTest(unittest.TestCase):

    def __init__(self, *args, **kwargs):
        super(ConvertWin32ToIcalTest, self).__init__(*args, **kwargs)

        self.events: list[W32Event] = []

        start_dt = datetime.datetime(year=2024, month=2, day=13, hour=12, minute=30, tzinfo=pytz.utc)

        self.event_with_end_args = {
            "id": "123",
            "subject": "Test",
            "start": start_dt,
            "end": start_dt + datetime.timedelta(hours=1),
            ### "duration": 60,
            "all_day": False,
            "body": "",
            "organizer": "",
            "busy_status": w32a_cal.BusyStatus.BUSY,
            "meeting_status": w32a_cal.MeetingStatus.RECEIVED,
            "importance": w32a_cal.Importance.NORMAL,
            "categories": "",
            "req_attendees": [],
            "opt_attendees": [],
            "recurring": False
        }
        self.event_with_end: W32Event = W32Event(**self.event_with_end_args)
        self.events.append(self.event_with_end)

        self.event_with_duration_args = {
            "id": "123",
            "subject": "Test",
            "start": start_dt,
            ### "end": start_dt + datetime.timedelta(hours=1),
            "duration": 60,
            "all_day": False,
            "body": "",
            "organizer": "",
            "busy_status": w32a_cal.BusyStatus.BUSY,
            "meeting_status": w32a_cal.MeetingStatus.RECEIVED,
            "importance": w32a_cal.Importance.NORMAL,
            "categories": "",
            "req_attendees": [],
            "opt_attendees": [],
            "recurring": False
        }
        self.event_with_duration: W32Event = W32Event(**self.event_with_duration_args)
        self.events.append(self.event_with_duration)

    def test_event_with_end(self):
        event_args = self.event_with_end_args.copy()
        event = W32Event(**event_args)

        ical_event: icalendar.Event = w32a_cal.win32_event_to_ical(event)[0]
        self.assertEqual(ical_event.get('UID'), event_args['id'])
        self.assertEqual(ical_event.get('SUMMARY'), event_args['subject'])
        self.assertEqual(ical_event.get('DTSTART').dt, event_args['start'])

        self.assertEqual(ical_event.get('DTEND').dt, self.event_with_end_args['end'])

        # AllDayEvent is False
        self.assertEqual(ical_event.get('DESCRIPTION'), event_args['body'])
        self.assertEqual(ical_event.get('ORGANIZER'), event_args['organizer'])

        # self.assertEqual(ical_event.get('CATEGORIES'), self.event_with_end_args['categories'])
        # self.assertEqual(ical_event.get('ATTENDEE'), self.event_with_end_args['req_attendees'])

        self.assertEqual(ical_event.get('RRULE').to_ical(), b'')

    def test_event_with_duration(self):
        event_args = self.event_with_duration_args.copy()
        event = W32Event(**event_args)

        ical_event: icalendar.Event = w32a_cal.win32_event_to_ical(event)[0]
        self.assertEqual(ical_event.get('UID'), event_args['id'])
        self.assertEqual(ical_event.get('SUMMARY'), event_args['subject'])
        self.assertEqual(ical_event.get('DTSTART').dt, event_args['start'])

        self.assertEqual(ical_event.get('DURATION').dt.total_seconds(), event_args['duration'] * 60)

        # AllDayEvent is False
        self.assertEqual(ical_event.get('DESCRIPTION'), event_args['body'])
        self.assertEqual(ical_event.get('ORGANIZER'), event_args['organizer'])

        # self.assertEqual(ical_event.get('CATEGORIES'), self.event_with_end_args['categories'])
        # self.assertEqual(ical_event.get('ATTENDEE'), self.event_with_end_args['req_attendees'])

        self.assertEqual(ical_event.get('RRULE').to_ical(), b'')

    def test_event_with_end_and_duration(self):
        event_args = self.event_with_duration_args.copy()
        event_args['end'] = event_args['start'] + datetime.timedelta(hours=1)

        event = W32Event(**event_args)

        ical_event: icalendar.Event = w32a_cal.win32_event_to_ical(event)[0]
        self.assertEqual(ical_event.get('UID'), event_args['id'])
        self.assertEqual(ical_event.get('SUMMARY'), event_args['subject'])
        self.assertEqual(ical_event.get('DTSTART').dt, event_args['start'])

        self.assertEqual(ical_event.get('DURATION').dt.total_seconds(), event_args['duration'] * 60)

        # AllDayEvent is False
        self.assertEqual(ical_event.get('DESCRIPTION'), event_args['body'])
        self.assertEqual(ical_event.get('ORGANIZER'), event_args['organizer'])

        # self.assertEqual(ical_event.get('CATEGORIES'), self.event_with_end_args['categories'])
        # self.assertEqual(ical_event.get('ATTENDEE'), self.event_with_end_args['req_attendees'])

        self.assertEqual(ical_event.get('RRULE').to_ical(), b'')



