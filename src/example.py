from typing import Optional
import datetime
from tabulate import tabulate
import w32a_cal
import icalendar
import logging

logging.basicConfig(level=logging.DEBUG)

def get_outlook_events(start: Optional[datetime.datetime] = None, end: Optional[datetime.datetime] = None):
  import win32com.client

  Outlook = win32com.client.Dispatch("Outlook.Application")
  ns = Outlook.GetNamespace("MAPI")
  appts = ns.GetDefaultFolder(9).Items
  appts.Sort("[Start]")

  if start or end:
    restriction: str = ""
    if start:
      restriction += "[Start] >= '" + start.strftime("%d/%m/%Y") + "'"

    if end:
      end_str = "[End] <= '" + end.strftime("%d/%m/%Y") + "'"
      if restriction:
        restriction += " AND "
        restriction += end_str
    appts = appts.Restrict(restriction)

  return appts

def get_outlook_month_events():
  start:datetime.datetime = datetime.datetime.now()
  end:datetime.datetime = start + datetime.timedelta(days = 34)
  appts = get_outlook_events(start, end)
  return appts

def print_outlook_month_events():
  import logging
  appts = get_outlook_month_events()
  calcTableHeader: list[str] = ['Subject', 'Organizer', 'Start', 'Duration(Minutes)', 'Recurring', 'Master', 'UID']
  calcTableBody: list[list[str]] = []
  for e in appts:
    row: list[str] = []
    row.append(e.Subject)
    row.append(e.Organizer)
    row.append(e.Start.Format(w32a_cal.OUTLOOK_DATE_FORMAT))
    row.append(e.Duration)
    row.append(e.IsRecurring)
    row.append(e.RecurrenceState == w32a_cal.RecurrenceState.MASTER)
    row.append(e.EntryID)
    calcTableBody.append(row)
  logging.info("\n%s",tabulate(calcTableBody, headers=calcTableHeader))

def outlook_events_to_ical(appts):
  ical_events = []
  for a in appts:
    ical_events.extend(w32a_cal.win32_event_to_ical(a))
  return ical_events


def print_outlook_month_events_to_ical():
  import logging
  appts = get_outlook_month_events()
  ical_events = outlook_events_to_ical(appts)
  calcTableHeader: list[str] = ['Title', 'Organizer', 'Start', 'Duration', 'Recurring', 'Master', 'UID']
  calcTableBody: list[list[str]] = []
  for e in ical_events:
    dtstart: icalendar.vDDDTypes = e.get('DTSTART', None)
    start_str: str = dtstart.dt.strftime("%m/%d/%Y, %H:%M:%S") if dtstart is not None else ""
    duration = e.get('DURATION', None)
    duration_str: str = str(duration.dt) if dtstart is not None else ""
    recurring = e.get('RRULE', None) is not None

    row: list[str] = []
    row.append(e.get('SUMMARY', ""))
    row.append(start_str)
    row.append(duration_str)
    row.append(recurring)
    row.append(recurring and e.get('RECURRENCE-ID', None) is None)
    row.append(e.get('UID'))
    calcTableBody.append(row)

  logging.info("\n%s",tabulate(calcTableBody, headers=calcTableHeader))
