from typing import Optional
import datetime
from tabulate import tabulate
import w32a_cal
import icalendar
import logging

logging.basicConfig(level=logging.DEBUG)


def get_outlook_calendar_folder(start: Optional[datetime.datetime] = None, end: Optional[datetime.datetime] = None, name: Optional[str] = None) -> object:
  import win32com.client

  Outlook = win32com.client.Dispatch("Outlook.Application")
  # https://learn.microsoft.com/en-us/office/vba/api/outlook.application.getnamespace
  # https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace
  # Use GetNameSpace ("MAPI") to return the Outlook NameSpace object from the Application object.
  ns = Outlook.GetNamespace("MAPI")

  folder = None

  # You can get another folder by name
  # for x in ns.Folders:
  #   logging.debug("ns.Folders: %s", x.Name)
  #   for y in x.Folders:
  #     logging.debug(y.Name)
  #     if y.Name == "Calendar":
  #       for z in y.Folders:
  #         logging.debug("z.Name: %s", z.Name)

  # for x in ns.GetDefaultFolder(9).Folders:
  #   logging.debug("default.Folders: %s", x.Name)


  # https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getdefaultfolder
  # obtains the default Calendar folder for the user who is currently logged on.
  # https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
  # olFolderCalendar 	9 	The Calendar folder.
  folder = ns.GetDefaultFolder(9)

  # https://learn.microsoft.com/en-us/office/vba/api/outlook.folder
  if folder is None:
    raise ValueError("No Outlok folder found")

  # https://learn.microsoft.com/en-us/office/vba/api/outlook.items
  if name:
    logging.debug("Looking for calendar: %s", name)
    for cal in folder.Folders:
      if cal.Name == name:
        folder = cal
        logging.debug("Found calendar: %s", cal.Name)
        break

  return folder

def get_outlook_events(start: Optional[datetime.datetime] = None, end: Optional[datetime.datetime] = None, name: Optional[str] = None) -> list[object]:
  cal_folder = get_outlook_calendar_folder(start, end, name)

  if cal_folder is None:
    raise ValueError("No Outlook calendar folder found")

  appts = cal_folder.Items

  if appts is None:
    raise ValueError("No Outlook calendar found")

  appts.Sort("[Start]")

  if start or end:
    restriction: str = ""
    if start:
      restriction += "[Start] >= '" + start.strftime(w32a_cal.OUTLOOK_DATE_FORMAT2) + "'"

    if end:
      end_str = "[End] <= '" + end.strftime(w32a_cal.OUTLOOK_DATE_FORMAT2) + "'"
      if restriction:
        restriction += " AND "
        restriction += end_str
    appts = appts.Restrict(restriction)

  logging.debug("appts.Count: %s", appts.Count)

  import w32obj
  # for e in appts:
  ae = w32obj.make_anonymous_event(appts[0])
  logging.debug(ae.__dict__)
  logging.debug(ae.GetRecurrencePattern())

  return appts

def dump_test_calendar(start: Optional[datetime.datetime] = None, end: Optional[datetime.datetime] = None, name: str = "TestCalendar", fpath: Optional[str] = None) -> icalendar.Calendar:
  # import pickle
  # import json
  import icalendar
  import tempfile

  if not fpath:
    raise ValueError("No file path provided")

  cal_folder = get_outlook_calendar_folder(name=name)

  if cal_folder is None:
    raise ValueError("No Outlook calendar folder found")

  cal_exporter = cal_folder.GetCalendarExporter()
  if cal_exporter is None:
    raise ValueError("No Outlook calendar exporter found")

  cal_exporter.CalendarDetail = w32a_cal.CalendarDetail.olFreeBusyAndSubject
  if start:
    cal_exporter.StartDate = start.strftime(w32a_cal.OUTLOOK_DATETIME_FORMAT)
  if end:
    cal_exporter.EndDate = end.strftime(w32a_cal.OUTLOOK_DATETIME_FORMAT)
  # cal_exporter.IncludeWholeCalendar = True
  # cal_exporter.IncludePrivateDetails = True

  logging.debug("Dump calendar %s to %s", name, fpath)
  ical = None

  # file = tempfile.NamedTemporaryFile(delete=False)
  with tempfile.NamedTemporaryFile(delete_on_close = False) as file:

    filename = file.name
    file.close()

    logging.debug("tempfile: %s", filename)
    cal_exporter.SaveAsICal(filename)
    with open (filename, "r") as written_file:
      ical = icalendar.Calendar.from_ical(written_file.read())
    logging.debug(ical.to_ical().decode('utf-8'))

  if fpath:
    with open(fpath, 'wb') as f:
      f.write(ical.to_ical())

  return ical

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
    row.append(e.Start.Format(w32a_cal.OUTLOOK_DATETIME_FORMAT))
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
    start_str: str = dtstart.dt.strftime(w32a_cal.OUTLOOK_DATETIME_FORMAT) if dtstart is not None else ""
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
