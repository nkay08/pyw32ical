import unittest
from w32obj import W32Event, W32RecurrencePattern, W32Exception
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
            "body": "TestBody",
            "organizer": "TestOrganizer",
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

        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        self.assertEqual(ical_event.get('UID'), event_args['id'])
        self.assertEqual(ical_event.get('SUMMARY'), event_args['subject'])
        self.assertEqual(ical_event.get('DTSTART').dt, event_args['start'])
        self.assertIsInstance(ical_event.get('DTSTART').dt, datetime.datetime)
        self.assertTrue(ical_event.get('DTSTART').dt.hour != 0 and ical_event.get('DTSTART').dt.minute != 0)

        self.assertEqual(ical_event.get('DTEND').dt, self.event_with_end_args['end'])
        self.assertIsInstance(ical_event.get('DTEND').dt, datetime.datetime)
        self.assertTrue(ical_event.get('DTEND').dt.hour != 0 and ical_event.get('DTEND').dt.minute != 0)

        self.assertIsNone(ical_event.get('DURATION'))

        # AllDayEvent is False
        self.assertEqual(ical_event.get('DESCRIPTION'), event_args['body'])
        self.assertEqual(ical_event.get('ORGANIZER'), event_args['organizer'])

        # self.assertEqual(ical_event.get('CATEGORIES'), self.event_with_end_args['categories'])
        # self.assertEqual(ical_event.get('ATTENDEE'), self.event_with_end_args['req_attendees'])

        self.assertEqual(ical_event.get('RRULE').to_ical(), b'')

    def test_event_with_duration(self):
        event_args = self.event_with_duration_args.copy()
        event = W32Event(**event_args)

        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        self.assertEqual(ical_event.get('UID'), event_args['id'])
        self.assertEqual(ical_event.get('SUMMARY'), event_args['subject'])
        self.assertEqual(ical_event.get('DTSTART').dt, event_args['start'])
        self.assertIsInstance(ical_event.get('DTSTART').dt, datetime.datetime)
        self.assertIsNone(ical_event.get('DTEND'))

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

        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        self.assertEqual(ical_event.get('UID'), event_args['id'])
        self.assertEqual(ical_event.get('SUMMARY'), event_args['subject'])
        self.assertEqual(ical_event.get('DTSTART').dt, event_args['start'])
        self.assertIsInstance(ical_event.get('DTSTART').dt, datetime.datetime)
        self.assertIsNone(ical_event.get('DTEND'))

        self.assertEqual(ical_event.get('DURATION').dt.total_seconds(), event_args['duration'] * 60)

        # AllDayEvent is False
        self.assertEqual(ical_event.get('DESCRIPTION'), event_args['body'])
        self.assertEqual(ical_event.get('ORGANIZER'), event_args['organizer'])

        # self.assertEqual(ical_event.get('CATEGORIES'), self.event_with_end_args['categories'])
        # self.assertEqual(ical_event.get('ATTENDEE'), self.event_with_end_args['req_attendees'])

        self.assertEqual(ical_event.get('RRULE').to_ical(), b'')


    def test_all_day(self):
        event_args = self.event_with_end_args.copy()
        event_args['all_day'] = True
        event = W32Event(**event_args)

        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        self.assertIsInstance(ical_event.get('DTSTART').dt, datetime.date)
        self.assertIsInstance(ical_event.get('DTEND').dt, datetime.date)
        self.assertIsNone(ical_event.get('DURATION'))

    def test_categories(self):
        pass

    def test_attendees(self):
        pass

    def test_recurring_daily(self):
        event_args = self.event_with_end_args.copy()
        event_args['recurring'] = True

        event_args['recurrence_state'] = w32a_cal.RecurrenceState.MASTER

        recurrence_interval = 1
        recurrence_count = 10
        recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.DAILY,
                                                  recurrence_interval,
                                                  recurrence_count)
        event_args['recurrence_pattern'] = recurrence_pattern

        event = W32Event(**event_args)
        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        # Check basic event properties
        self.assertEqual(ical_event.get('UID'), event_args['id'])
        self.assertEqual(ical_event.get('SUMMARY'), event_args['subject'])
        self.assertEqual(ical_event.get('DTSTART').dt, event_args['start'])
        self.assertIsInstance(ical_event.get('DTSTART').dt, datetime.datetime)
        self.assertTrue(ical_event.get('DTSTART').dt.hour != 0 and ical_event.get('DTSTART').dt.minute != 0)

        self.assertEqual(ical_event.get('DTEND').dt, self.event_with_end_args['end'])
        self.assertIsInstance(ical_event.get('DTEND').dt, datetime.datetime)
        self.assertTrue(ical_event.get('DTEND').dt.hour != 0 and ical_event.get('DTEND').dt.minute != 0)

        self.assertIsNone(ical_event.get('DURATION'))

        self.assertEqual(ical_event.get('DESCRIPTION'), event_args['body'])
        self.assertEqual(ical_event.get('ORGANIZER'), event_args['organizer'])

        # Check recurrence properties
        self.assertIsNotNone(ical_event.get('RRULE'))
        self.assertEqual(ical_event.get('RRULE').get('FREQ'), 'DAILY')
        self.assertEqual(ical_event.get('RRULE').get('INTERVAL'), recurrence_interval)
        self.assertEqual(ical_event.get('RRULE').get('COUNT'), recurrence_count)

    def test_recurring_weekly(self):
        event_args = self.event_with_end_args.copy()
        event_args['recurring'] = True

        event_args['recurrence_state'] = w32a_cal.RecurrenceState.MASTER

        recurrence_interval = 1
        recurrence_count = 10
        dayofweekmask = w32a_cal.DayOfWeekMaskEnum.MONDAY | w32a_cal.DayOfWeekMaskEnum.WEDNESDAY | w32a_cal.DayOfWeekMaskEnum.FRIDAY

        with self.assertRaises(ValueError):
            recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.WEEKLY,
                                                    recurrence_interval,
                                                    occurrences=recurrence_count,
                                                    day_of_week_mask=None)
            event_args['recurrence_pattern'] = recurrence_pattern
            event = W32Event(**event_args)
            ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)

        recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.WEEKLY,
                                                  recurrence_interval,
                                                  occurrences=recurrence_count,
                                                  day_of_week_mask=dayofweekmask)
        event_args['recurrence_pattern'] = recurrence_pattern
        event = W32Event(**event_args)
        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        self.assertIsNotNone(ical_event.get('RRULE'))
        self.assertEqual(len(ical_event.get('RRULE').items()), 4)
        self.assertEqual(ical_event.get('RRULE').get('FREQ'), 'WEEKLY')
        self.assertEqual(ical_event.get('RRULE').get('INTERVAL'), recurrence_interval)
        self.assertEqual(ical_event.get('RRULE').get('COUNT'), recurrence_count)
        self.assertEqual(ical_event.get('RRULE').get('byday'), ['MO', 'WE', 'FR'])

    def test_recurring_monthly(self):
        event_args = self.event_with_end_args.copy()
        event_args['recurring'] = True

        event_args['recurrence_state'] = w32a_cal.RecurrenceState.MASTER

        recurrence_interval = 1
        recurrence_count = 10
        bymonthday = 13

        with self.assertRaises(ValueError):
            recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.MONTHLY,
                                                    recurrence_interval,
                                                    occurrences=recurrence_count,
                                                    day_of_month=None)
            event_args['recurrence_pattern'] = recurrence_pattern
            event = W32Event(**event_args)
            ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)

        recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.MONTHLY,
                                                  recurrence_interval,
                                                  occurrences=recurrence_count,
                                                  day_of_month=bymonthday)
        event_args['recurrence_pattern'] = recurrence_pattern
        event = W32Event(**event_args)
        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        self.assertIsNotNone(ical_event.get('RRULE'))
        self.assertEqual(len(ical_event.get('RRULE').items()), 4)
        self.assertEqual(ical_event.get('RRULE').get('FREQ'), 'MONTHLY')
        self.assertEqual(ical_event.get('RRULE').get('INTERVAL'), recurrence_interval)
        self.assertEqual(ical_event.get('RRULE').get('COUNT'), recurrence_count)
        self.assertEqual(ical_event.get('RRULE').get('BYMONTHDAY'), bymonthday)

    def test_recurring_yearly(self):
        event_args = self.event_with_end_args.copy()
        event_args['recurring'] = True

        event_args['recurrence_state'] = w32a_cal.RecurrenceState.MASTER

        recurrence_interval = 1
        recurrence_count = 10
        bymonthday = 13
        bymonth = 2

        recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.YEARLY,
                                                    recurrence_interval,
                                                    occurrences=recurrence_count,
                                                    day_of_month=bymonthday,
                                                    month_of_year=None)
        event_args['recurrence_pattern'] = recurrence_pattern
        event = W32Event(**event_args)
        with self.assertRaises(ValueError):
            ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)

        recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.YEARLY,
                                                    recurrence_interval,
                                                    occurrences=recurrence_count,
                                                    day_of_month=None,
                                                    month_of_year=bymonth)
        event_args['recurrence_pattern'] = recurrence_pattern
        event = W32Event(**event_args)
        with self.assertRaises(ValueError):
            ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)

        recurrence_pattern = W32RecurrencePattern(w32a_cal.RecurrenceType.YEARLY,
                                                    recurrence_interval,
                                                    occurrences=recurrence_count,
                                                    day_of_month=bymonthday,
                                                    month_of_year=bymonth)
        event_args['recurrence_pattern'] = recurrence_pattern
        event = W32Event(**event_args)
        ical_events: list[icalendar.Event] = w32a_cal.win32_event_to_ical(event)
        self.assertEqual(len(ical_events), 1)
        ical_event: icalendar.Event = ical_events[0]

        self.assertIsNotNone(ical_event.get('RRULE'))
        self.assertEqual(len(ical_event.get('RRULE').items()), 5)
        self.assertEqual(ical_event.get('RRULE').get('FREQ'), 'YEARLY')
        self.assertEqual(ical_event.get('RRULE').get('INTERVAL'), recurrence_interval)
        self.assertEqual(ical_event.get('RRULE').get('COUNT'), recurrence_count)
        self.assertEqual(ical_event.get('RRULE').get('BYMONTHDAY'), bymonthday)
        self.assertEqual(ical_event.get('RRULE').get('BYMONTH'), bymonth)

    def test_recurring_monthly_nth(self):
        pass

    def test_recurring_yearly_nth(self):
        pass