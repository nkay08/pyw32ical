from typing import Optional
import datetime
from w32a_cal import BusyStatus, MeetingStatus, Importance, RecurrenceState, RecurrenceType, OUTLOOK_DATE_FORMAT

def datetime_to_w32str(dt: datetime.datetime) -> str:
    # TODO: check
    return dt.strftime(OUTLOOK_DATE_FORMAT)
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")

# Approximation of W32 (Outlook) event related objects that are available via pywin32

class W32Exception:
    def __init__(self,
                 original_date: datetime.datetime,
                 deleted: bool = False,
                 event: Optional['W32Event'] = None) -> None:

        self.OriginalDate: str  = datetime_to_w32str(original_date)
        self.Deleted: bool = deleted
        self.AppointmentItem: Optional['W32Event'] = event

class W32RecurrencePattern:

    def __init__(self,
                 recurrence_type: RecurrenceType,
                 interval: int,
                 occurrences: Optional[int] = None,
                 end: Optional[datetime.datetime] = None,
                 no_end: bool = False,
                 exceptions: list[W32Exception] = []) -> None:
        self.RecurrenceType: RecurrenceType = recurrence_type
        self.Interval: int = interval
        self.NoEndDate: bool = no_end

        self.Occurrences: Optional[int] = None
        self.PatternEndDate: Optional[str] = None
        self.EndTime: Optional[str] = None
        if not no_end:
            if occurrences is not None and occurrences > 0:
                self.Occurrences = occurrences
            elif end is not None:
                self.PatternEndDate = datetime_to_w32str(end.date())
                self.EndTime = datetime_to_w32str(end.time())
            else:
                raise ValueError("Either occurrences or end date must be specified")

        self.Exceptions: list[W32Exception] = exceptions

class W32Event:
    def __init__(self,
                 id: str,
                 subject: str,
                 start: datetime.datetime,
                 duration: Optional[int] = None,
                 end: Optional[datetime.datetime] = None,
                 creation_time: Optional[datetime.datetime] = None,
                 modification_time: Optional[datetime.datetime] = None,
                 all_day: bool = False,
                 body: str = "",
                 organizer: str = "",
                 busy_status: Optional[BusyStatus] = BusyStatus.FREE,
                 meeting_status: Optional[MeetingStatus] = MeetingStatus.RECEIVED,
                 importance: Optional[Importance] = Importance.NORMAL,
                 location: str = "",
                 categories: str = "",
                 req_attendees: list[str] = [],
                 opt_attendees: list[str] = [],
                 recurring: bool = False,
                 recurrence_state: RecurrenceState = RecurrenceState.NOT_RECURRING,
                 recurrence_pattern: Optional[W32RecurrencePattern] = None,
                 exceptions: list[tuple[datetime.datetime, bool, Optional[datetime.datetime]]] = [],
                 ) -> None:
        self.EntryID: str = id
        self.GlobalAppointmentId: str = id
        self.StartUTC: str = datetime_to_w32str(start)
        self.Subject: str = subject

        self.Duration: Optional[int] = None
        self.EndUTC: Optional[str] = None

        if duration:
            self.Duration = duration
        else:
            self.EndUTC: str = datetime_to_w32str(end)

        self.CreationTime: str = datetime_to_w32str(creation_time) if creation_time else self.StartUTC
        self.LastModificationTime: str  = datetime_to_w32str(modification_time) if modification_time else self.CreationTime

        self.AllDayEvent: bool = all_day
        self.Body: str = body
        self.Organizer: str  = organizer
        self.BusyStatus: BusyStatus = busy_status
        self.MeetingStatus: MeetingStatus = meeting_status
        self.Importance: Importance = importance
        self.Location: str  = location
        self.Categories: str  = categories
        self.RequiredAttendees: str  = ";".join(req_attendees) if len(req_attendees) > 0 else ""
        self.OptionalAttendees: str  = ";".join(opt_attendees) if len(opt_attendees) > 0 else ""

        self.IsRecurring: bool = recurring
        self.RecurrenceState: RecurrenceState = recurrence_state

        self._RecurrencePattern: Optional[W32RecurrencePattern] = None

        if recurring and recurrence_state == RecurrenceState.MASTER:
            self._RecurrencePattern: W32RecurrencePattern = recurrence_pattern
            exceptions = []
            if len(exceptions) > 0:
                for e in exceptions:
                    new_event = None
                    if e[2] is not None:
                        new_dur_min = duration if duration else (end - start).minutes
                        new_end = e[2] + datetime.timedelta(minutes=new_dur_min)
                        new_event = W32Event(id=id,subject=subject,start=e[2], duration=duration, end=new_end)
                    new_exc = W32Exception(e[0], e[1], new_event)
                    exceptions.append(new_exc)
                self._RecurrencePattern.Exceptions = exceptions


    def GetRecurrencePattern(self) -> W32RecurrencePattern:
        if self.IsRecurring and self.RecurrenceState == RecurrenceState.MASTER:
            return self._RecurrencePattern
        return None