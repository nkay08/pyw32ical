from enum import IntEnum, IntFlag
import datetime
import dateutil.parser
import pytz
import icalendar

import logging

from typing import Optional

ICAL_FILTER_FULL={
  "summary": True, # or "subject"
  "description": True, # or "body"
  "busy": True, # or "transp"
  "organizer": True,
  "status": True, # or "meetingstatus",
  "categories": True,
  "importance": True, # or "priority"
  "location": True,
  "attendees": True
  }

ICAL_FILTER_SAFE={
  "busy": True,
  "status": True
  }

OUTLOOK_DATE_FORMAT = '%m/%d/%Y %H:%M'

# https://learn.microsoft.com/en-us/office/vba/api/outlook.olimportance
class Importance(IntEnum):
  LOW = 0
  NORMAL = 1
  HIGH = 2

  # https://docs.microsoft.com/en-us/office/vba/api/outlook.olbusystatus
class BusyStatus(IntEnum):
  FREE = 0
  TENTATIVE = 1
  BUSY = 2
  OUT_OF_OFFICE = 3
  WORKING_ELSEWHERE = 4

  # https://learn.microsoft.com/en-us/office/vba/api/outlook.olmeetingstatus
class MeetingStatus(IntEnum):
  NON_MEETING = 0
  MEETING = 1
  RECEIVED = 3
  CANCELED = 5
  RECEIVED_AND_CANCELED = 7

  # https://learn.microsoft.com/en-us/office/vba/api/outlook.olrecurrencetype
class RecurrenceType(IntEnum):
  DAILY = 0
  WEEKLY = 1
  MONTHLY = 2
  MONTHLY_NTH = 3
  YEARLY = 5
  YEARLY_NTH = 6

# https://learn.microsoft.com/en-us/office/vba/api/outlook.olrecurrencestate
class RecurrenceState(IntEnum):
  NOT_RECURRING = 0
  MASTER = 1
  OCCURRENCE = 2
  EXCEPTION = 3

class DayOfWeekMaskEnum(IntFlag):
  MONDAY = 2
  TUESDAY = 4
  WEDNESDAY = 8
  THURSDAY = 16
  FRIDAY = 32
  SATURDAY = 64
  SUNDAY = 1


def _win32_busystatus_to_ical(status: MeetingStatus) -> Optional[str]:
  BusyStatus2Ical = {
    BusyStatus.FREE: "TRANSPARENT",
    BusyStatus.TENTATIVE: "TRANSPARENT",
    BusyStatus.BUSY: "OPAQUE",
    BusyStatus.OUT_OF_OFFICE: "OPAQUE",
    BusyStatus.WORKING_ELSEWHERE: "OPAQUE",
  }
  return BusyStatus2Ical.get(status, None)

def _win32_meetingstatus_to_ical(status: MeetingStatus) -> Optional[str]:
  MeetingStatus2Ical = {
    MeetingStatus.NON_MEETING:   "TENTATIVE",
    MeetingStatus.MEETING:  "CONFIRMED",
    MeetingStatus.RECEIVED: "TENTATIVE",
    MeetingStatus.CANCELED: "CANCELLED",
    MeetingStatus.RECEIVED_AND_CANCELED: "CANCELLED"
  }
  return MeetingStatus2Ical.get(status, None)

def _win32_recurrence_type_to_ical(rec_type: RecurrenceType) -> Optional[str]:
  RecurrenceType2Ical = {
    RecurrenceType.DAILY.value: "DAILY",
    RecurrenceType.WEEKLY.value: "WEEKLY",
    RecurrenceType.MONTHLY.value: "MONTHLY",
    RecurrenceType.MONTHLY_NTH.value: "MONTHLY",
    RecurrenceType.YEARLY.value: "YEARLY",
    RecurrenceType.YEARLY_NTH.value: "YEARLY"
  }
  return RecurrenceType2Ical.get(rec_type, None)

def win32_date_to_datetime(d: str, utc: bool = False, tz: Optional[datetime.tzinfo] = None) -> datetime.datetime:
  dt = dateutil.parser.parse(str(d))
  if tz is not None:
    dt = dt.replace(tzinfo=tz)
  elif utc:
    dt = dt.replace(tzinfo=pytz.utc)

  return dt

def _win32_day_of_week_mask_to_ical_str(win32_mask: DayOfWeekMaskEnum) -> list[int]:
  rrule_weekday = []
  if (win32_mask & DayOfWeekMaskEnum.MONDAY):
    rrule_weekday.append("MO")
  if (win32_mask & DayOfWeekMaskEnum.TUESDAY):
    rrule_weekday.append("TU")
  if (win32_mask & DayOfWeekMaskEnum.WEDNESDAY):
    rrule_weekday.append("WE")
  if (win32_mask & DayOfWeekMaskEnum.THURSDAY):
    rrule_weekday.append("TH")
  if (win32_mask & DayOfWeekMaskEnum.FRIDAY):
    rrule_weekday.append("FR")
  if (win32_mask & DayOfWeekMaskEnum.SATURDAY):
    rrule_weekday.append("SA")
  if (win32_mask & DayOfWeekMaskEnum.SUNDAY):
    rrule_weekday.append("SU")
  return rrule_weekday

def _win32_day_of_week_mask_to_ical_int(win32_mask) -> list[int]:
  rrule_weekday = []
  if (win32_mask & DayOfWeekMaskEnum.MONDAY):
    rrule_weekday.append(0)
  if (win32_mask & DayOfWeekMaskEnum.TUESDAY):
    rrule_weekday.append(1)
  if (win32_mask & DayOfWeekMaskEnum.WEDNESDAY):
    rrule_weekday.append(2)
  if (win32_mask & DayOfWeekMaskEnum.THURSDAY):
    rrule_weekday.append(3)
  if (win32_mask & DayOfWeekMaskEnum.FRIDAY):
    rrule_weekday.append(4)
  if (win32_mask & DayOfWeekMaskEnum.SATURDAY):
    rrule_weekday.append(5)
  if (win32_mask & DayOfWeekMaskEnum.SUNDAY):
    rrule_weekday.append(6)
  return rrule_weekday

def _win32_day_of_week_mask_valid_for_type(rtype):
  if rtype is None:
    return False

  if rtype == RecurrenceType.WEEKLY or rtype == RecurrenceType.MONTHLY_NTH or rtype == RecurrenceType.YEARLY_NTH:
    return True

  return False

def _win32_importance_to_ical(win32_importance):
  if win32_importance == Importance.HIGH:
    return 4
  elif win32_importance == Importance.NORMAL:
    return 5
  else:
    return 6

def _win32_event_recurrence_to_rrule_dict(win32_event) -> dict:
  # https://icalendar.org/rrule-tool.html
  # DTSTART is defined in the Event, so we do not need it here

  if not win32_event.IsRecurring or win32_event.RecurrenceState != RecurrenceState.MASTER:
    return {}

  win32_recurrence = win32_event.GetRecurrencePattern()

  rrule_dict = {
    'freq': _win32_recurrence_type_to_ical(win32_recurrence.RecurrenceType),
    'interval': win32_recurrence.Interval,
  }

  if not win32_recurrence.NoEndDate:
    if win32_recurrence.PatternEndDate is not None:
      end_date = win32_date_to_datetime(win32_recurrence.PatternEndDate)
      if win32_recurrence.EndTime is not None:
        end_date = datetime.datetime.combine(end_date.date(), win32_date_to_datetime(win32_recurrence.EndTime).time())
      # Needs timezone
      end_date = end_date.replace(tzinfo=pytz.utc)
      rrule_dict['until'] = end_date
    elif win32_recurrence.Occurrences > 0:
      rrule_dict['count'] = win32_recurrence.Occurrences

  rtype = win32_recurrence.RecurrenceType
  day_of_week_mask = None
  if _win32_day_of_week_mask_valid_for_type(win32_recurrence.RecurrenceType):
    day_of_week_mask = _win32_day_of_week_mask_to_ical_str(win32_recurrence.DayOfWeekMask)

  if rtype == RecurrenceType.WEEKLY:
    if day_of_week_mask is not None:
      rrule_dict['byday'] = day_of_week_mask
    else:
      raise ValueError("DayOfWeekMask must be set for WEEKLY recurrence")
  if rtype == RecurrenceType.MONTHLY:
    # rrule_dict['byweekday'] = self.day_of_week_mask.to_rrule_weekday()
    if getattr(win32_recurrence, 'DayOfMonth', None) is not None:
      rrule_dict['bymonthday'] = win32_recurrence.DayOfMonth
    else:
      raise ValueError("DayOfMonth must be set for MONTHLY recurrence")
  if rtype == RecurrenceType.MONTHLY_NTH:
    # TODO: INSTANCE https://learn.microsoft.com/en-us/office/vba/api/outlook.recurrencepattern
    if getattr(win32_recurrence, 'DayOfMonth', None) is not None:
      rrule_dict['bymonthday'] = win32_recurrence.DayOfMonth
    if day_of_week_mask is not None:
      rrule_dict['byday'] = day_of_week_mask
    if not rrule_dict.get('byday') and not rrule_dict.get('bymonthday'):
      raise ValueError("DayOfMonth or DayOfWeekMask must be set for MONTHLY_NTH recurrence")
  if rtype == RecurrenceType.YEARLY:
    if getattr(win32_recurrence, 'DayOfMonth', None) is not None:
      rrule_dict['bymonthday'] = win32_recurrence.DayOfMonth
    if getattr(win32_recurrence, 'MonthOfYear', None) is not None:
      rrule_dict['bymonth'] = win32_recurrence.MonthOfYear
    if not rrule_dict.get('bymonthday') and not rrule_dict.get('bymonth'):
      raise ValueError("DayOfMonth or MonthOfYear must be set for YEARLY recurrence")
  if rtype == RecurrenceType.YEARLY_NTH:
    # TODO INSTANCE
    if day_of_week_mask is not None:
      rrule_dict['byday'] = day_of_week_mask
    if getattr(win32_recurrence, 'MonthOfYear', None) is not None:
      rrule_dict['bymonth'] = win32_recurrence.MonthOfYear
    if not rrule_dict.get('byday') and not rrule_dict.get('bymonth'):
      raise ValueError("DayOfWeekMask or MonthOfYear must be set for YEARLY_NTH recurrence")

  return rrule_dict


def win32_event_to_ical(win32_event, parse_recurrence: bool = True, filter: Optional[dict] = None) -> list[icalendar.Event]:
  import pytz
  import icalendar
  event_list: list[icalendar.Event] = []
  ical_event:icalendar.Event = icalendar.Event()

  # Recurrences should not have different UID
  # with GlobalAppointmentId recurrence exceptions may have different UID!
  # GlobalAppointmentId # https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.globalappointmentid
  # https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.entryid
  ical_event.add('UID', win32_event.EntryID)

  start = win32_date_to_datetime(win32_event.StartUTC, utc=True)
  if getattr(win32_event, "AllDayEvent", False):
    ical_event.add("DTSTART", start.date())
  else:
    ical_event.add("DTSTART", start)

  # https://icalendar.org/iCalendar-RFC-5545/3-8-7-1-date-time-created.html
  # https://icalendar.org/iCalendar-RFC-5545/3-8-7-2-date-time-stamp.html
  if getattr(win32_event, "CreationTime", None) is not None:
    ical_event.add('DTSTAMP', win32_date_to_datetime(win32_event.CreationTime, utc=True))
    ical_event.add('CREATED', win32_date_to_datetime(win32_event.CreationTime, utc=True))

  # https://icalendar.org/iCalendar-RFC-5545/3-8-7-3-last-modified.html
  ical_event.add('LAST-MODIFIED', win32_date_to_datetime(win32_event.LastModificationTime, utc=True))

  # DTEND and DURATION properties must not occur in the same VEVENT Reference: RFC 5545 3.6.1. Event Component
  # http://icalendar.org/iCalendar-RFC-5545/3-6-1-event-component.html
  if getattr(win32_event, "Duration", 0) is not None and win32_event.Duration > 0:
    ical_event.add('DURATION', datetime.timedelta(minutes = win32_event.Duration))
  else:
    end = win32_date_to_datetime(win32_event.EndUTC, utc=True)
    if getattr(win32_event, "AllDayEvent", False):
      # TODO: set it https://github.com/icalendar/icalendar/issues/71
      ical_event.add("DTEND", end.date())
    else:
      ical_event.add("DTEND", end)


  if filter is None or filter.get("summary", False) or filter.get("subject", False):
    # string
    ical_event.add('SUMMARY', win32_event.Subject)
  else:
    ical_event.add('SUMMARY', "Event")

  if filter is None or filter.get("description", False) or filter.get("body", False):
    # string
    if getattr(win32_event, "Body", None) is not None:
      ical_event.add('DESCRIPTION', win32_event.Body)

  if filter is None or filter.get("organizer", False):
    # string
    if getattr(win32_event, "Organizer", None) is not None:
      ical_event.add('ORGANIZER', win32_event.Organizer)

  if filter is None or filter.get("transp", False) or filter.get("busy", False):
    # https://docs.microsoft.com/en-us/office/vba/api/outlook.olbusystatus
    if getattr(win32_event, "BusyStatus", None) is not None:
      ical_event.add('TRANSP', _win32_busystatus_to_ical(win32_event.BusyStatus))

  if filter is None or filter.get("status", False) or filter.get("meetingstatus", False):
    # https://docs.microsoft.com/en-us/office/vba/api/outlook.olmeetingstatus
    if getattr(win32_event, "MeetingStatus", None) is not None:
      ical_event.add('STATUS', _win32_meetingstatus_to_ical(win32_event.MeetingStatus))

  if filter is None or filter.get("location", False):
    # string
    if getattr(win32_event, "Location", None) is not None:
      ical_event.add('LOCATION', win32_event.Location)

  if filter is None or filter.get("categories", False):
    # string
    if getattr(win32_event, "Categories", None) is not None:
      ical_event.add('CATEGORIES', win32_event.Categories)

  if filter is None or filter.get("attendees", False):
    # str, semicolon delimited
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.requiredattendees
    required_attendees_str = getattr(win32_event, "RequiredAttendees", "")
    optional_attendees_str = getattr(win32_event, "OptionalAttendees", "")

    for attendee in required_attendees_str.split(";"):
      if attendee:
        ical_event.add('ATTENDEE', attendee, parameters={'ROLE':'REQ-PARTICIPANT'})
    for attendee in optional_attendees_str.split(";"):
      if attendee:
        ical_event.add('ATTENDEE', attendee)

  if filter is None or filter.get("priority", False) or filter.get("importance", False):
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.importance
    if getattr(win32_event, "Importance", None) is not None:
      ical_event.add('PRIORITY', _win32_importance_to_ical(win32_event.Importance))

  # recurrence
  sequence = 1

  # To avoid recursion do not parse recurrence when parsing recurrence exceptions
  if parse_recurrence:

    logging.debug("Parse recurrence for event %s - %s", win32_event.EntryID, win32_event.Subject)

    if win32_event.IsRecurring and win32_event.RecurrenceState == RecurrenceState.MASTER:
      win32_recurrence = win32_event.GetRecurrencePattern()
      if win32_recurrence is not None:
        exdate_list: list[datetime.datetime] = []
        for ex in win32_recurrence.Exceptions:
          exdate_datetime: datetime.datetime = datetime.datetime.combine(win32_date_to_datetime(ex.OriginalDate).date(),
                                                                        win32_date_to_datetime(win32_event.StartUTC).time())
          # We have to add the timezone or else, the recurrence-id does not match with the original ical date
          # -> without tz UTC, this would result in missing "Z" at the end of the datetime string
          exdate_datetime = exdate_datetime.replace(tzinfo=pytz.utc)
          exdate_vdate = icalendar.vDatetime(exdate_datetime)
          if not ex.Deleted:
            logging.debug("Parsing recurrence exception event")
            if ex.AppointmentItem is not None:
              # parse_recurrence must be False to avoid potential recursion!
              ex_ical_event = win32_event_to_ical(ex.AppointmentItem, parse_recurrence=False, filter=filter)[0]
              ex_ical_event.add("RECURRENCE-ID", exdate_vdate)
              if ex_ical_event.get('UID') != ical_event.get('UID'):
                logging.warning("Event and recurrence exception have different UID: %s <> %s", ical_event.decoded('UID').decode(), ex_ical_event.decoded('UID').decode())
                ex_ical_event['UID'] = win32_event.EntryID
              event_list.append(ex_ical_event)
            else: # Deleted
              exdate_list.append(exdate_datetime)
          sequence += 1

        if len(exdate_list) > 0:
          ical_event.add("EXDATE", exdate_list, parameters={'VALUE':'DATE-TIME'})

  ical_event.add("RRULE", _win32_event_recurrence_to_rrule_dict(win32_event))
  ical_event.add('SEQUENCE', sequence)

  event_list.insert(0, ical_event)

  return event_list