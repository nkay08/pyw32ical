from typing import Optional
import datetime
import pytz
from tzlocal.windows_tz import tz_win
from w32a_cal import BusyStatus, MeetingStatus, Importance, RecurrenceState, RecurrenceType, OUTLOOK_DATETIME_FORMAT, _win32_day_of_week_mask_valid_for_type

def datetime_to_w32str(dt: datetime.datetime) -> str:
    # TODO: check
    return dt.strftime(OUTLOOK_DATETIME_FORMAT)
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
                 day_of_week_mask: Optional[int] = None,
                 month_of_year: Optional[int] = None,
                 day_of_month: Optional[int] = None,
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

        self.DayOfWeekMask: Optional[int] = None
        self.MonthOfYear: Optional[int] = None
        self.DayOfMonth: Optional[int] = None

        if _win32_day_of_week_mask_valid_for_type(recurrence_type):
            self.DayOfWeekMask = day_of_week_mask

        if (recurrence_type == RecurrenceType.MONTHLY
            or recurrence_type == RecurrenceType.MONTHLY_NTH
            or recurrence_type == RecurrenceType.YEARLY
            or recurrence_type == RecurrenceType.YEARLY_NTH
            ):
            self.DayOfMonth = day_of_month
        if (recurrence_type == RecurrenceType.YEARLY
            or recurrence_type == RecurrenceType.YEARLY_NTH
            ):
            self.MonthOfYear = month_of_year

        self.Exceptions: list[W32Exception] = exceptions

class W32TimeZone:
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.timezone

    def __init__(self, id: str, name: Optional[str] = None,
                 bias: Optional[int] = 0,
                 standard_date: Optional[datetime.datetime] = None, standard_bias: Optional[int] = 0,
                 daylight_date: Optional[datetime.datetime] = None, daylight_bias: Optional[int] = 0) -> None:

        self.ID = id
        self.Name = name if name is not None else id

        self.Bias = bias

        self.StandardDate = standard_date
        self.StandardBias = standard_bias

        self.DaylightBias = daylight_bias
        self.DaylightDate = daylight_date

        self.DaylightDesignation = None
        self.StandardDesignation = None
        self.Class = None
        self.Application = None
        self.Session = None
        self.Parent = None


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

        self.Subject: str = subject

        start_tz: datetime.tzinfo | None = start.tzinfo


        if start_tz is not None:
            start_tz_str = tz_win.get(start.tzname())
        else:
            start_tz_str = tz_win.get(pytz.utc.tzname())
            start.replace(tzinfo=pytz.utc)

        self.StartTimeZone = W32TimeZone(id=start_tz_str)

        self.Start = datetime_to_w32str(start)
        self.StartUTC: str = datetime_to_w32str(start.astimezone(pytz.utc))

        self.Duration: Optional[int] = None
        self.End: Optional[str] = None
        self.EndUTC: Optional[str] = None

        if duration:
            self.Duration = duration
        if end is not None:
            end_tz: datetime.tzinfo | None = end.tzinfo
            if end_tz is not None:
                end_tz_str = tz_win.get(end.tzname())
            else:
                if start_tz is not None:
                    end_tz_str = tz_win.get(start_tz)
                    end.replace(tzinfo=start_tz)
                else:
                    end_tz_str = tz_win.get(pytz.utc.tzname())
                    end.replace(tzinfo=pytz.utc)
            self.End: str = datetime_to_w32str(end)
            self.EndUTC: str = datetime_to_w32str(end.astimezone(pytz.utc))
            self.EndTimeZone = W32TimeZone(id=end_tz_str)

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


class AnonymousObject:

    @staticmethod
    def make_callable(event: object, method: object) -> callable:
        value: object = method()
        def _callable(self):
            return value
        return _callable

    def __init__(self, properties: dict, methods: dict, event: Optional[object] = None):
        self.__dict__.update(properties)
        # for k,v in methods.items():
        #     setattr(self, k, self.make_callable(event, v))#

    def GetRecurrencePattern(self) -> object:
        return getattr(self, 'RecurrencePattern', None)

def make_value_callable(value):

    def _callable():
        return value

    return _callable

def get_win32_event_property_dict(win32_event) -> dict:

    w32_start_tz = getattr(win32_event, 'StartTimeZone', None)
    if (w32_start_tz is not None):
        start_tz = {'ID': w32_start_tz.ID, 'Name': w32_start_tz.Name, 'Bias': w32_start_tz.Bias,
                    'StandardDate': w32_start_tz.StandardDate, 'StandardBias': w32_start_tz.StandardBias,
                    'DaylightDate': w32_start_tz.DaylightDate, 'DaylightBias': w32_start_tz.DaylightBias}
    else:
        start_tz = {'ID': tz_win.get(pytz.utc.tzname()), 'Name': "UTC", 'Bias': 0,
            'StandardDate': None, 'StandardBias': 0,
            'DaylightDate': None, 'DaylightBias': 0}

    w32_end_tz = getattr(win32_event, 'EndTimeZone', None)
    if (w32_end_tz is not None):
        end_tz = {'ID': w32_end_tz.ID, 'Name': w32_end_tz.Name, 'Bias': w32_end_tz.Bias,
                    'StandardDate': w32_end_tz.StandardDate, 'StandardBias': w32_end_tz.StandardBias,
                    'DaylightDate': w32_end_tz.DaylightDate, 'DaylightBias': w32_end_tz.DaylightBias}
    else:
        end_tz = {'ID': tz_win.get(pytz.utc.tzname()), 'Name': "UTC", 'Bias': 0,
                    'StandardDate': None, 'StandardBias': 0,
                    'DaylightDate': None, 'DaylightBias': 0}

    win32_event_properties = {
        'AllDayEvent': getattr(win32_event, 'AllDayEvent', None),
        # 'Application': getattr(win32_event, 'Application', None), #<COMObject <unknown>>
        # 'Attachments': getattr(win32_event, 'Attachments', None), # <COMObject <unknown>>
        'AutoResolvedWinner': getattr(win32_event, 'AutoResolvedWinner', None),
        'BillingInformation': getattr(win32_event, 'BillingInformation', None),
        'Body': getattr(win32_event, 'Body', None),
        'BusyStatus': getattr(win32_event, 'BusyStatus', None),
        'Categories': getattr(win32_event, 'Categories', None),
        'Class': getattr(win32_event, 'Class', None),
        'Companies': getattr(win32_event, 'Companies', None),
        # 'Conflicts': getattr(win32_event, 'Conflicts', None), # <COMObject <unknown>>
        'ConversationID': getattr(win32_event, 'ConversationID', None),
        'ConversationIndex': getattr(win32_event, 'ConversationIndex', None),
        'ConversationTopic': getattr(win32_event, 'ConversationTopic', None),
        'CreationTime': getattr(win32_event, 'CreationTime', None),
        'DownloadState': getattr(win32_event, 'DownloadState', None),
        'Duration': getattr(win32_event, 'Duration', None),
        'End': getattr(win32_event, 'End', None),
        'EndInEndTimeZone': getattr(win32_event, 'EndInEndTimeZone', None),
        'EndTimeZone': end_tz, # https://learn.microsoft.com/en-us/office/vba/api/outlook.timezone
        'EndUTC': getattr(win32_event, 'EndUTC', None),
        'EntryID': getattr(win32_event, 'EntryID', None),
        'ForceUpdateToAllAttendees': getattr(win32_event, 'ForceUpdateToAllAttendees', None),
        # 'FormDescription': getattr(win32_event, 'FormDescription', None), # <COMObject <unknown>>
        # 'GetInspector': getattr(win32_event, 'GetInspector', None), # <COMObject <unknown>>
        'GlobalAppointmentID': getattr(win32_event, 'GlobalAppointmentID', None),
        'Importance': getattr(win32_event, 'Importance', None),
        'InternetCodepage': getattr(win32_event, 'InternetCodepage', None),
        'IsConflict': getattr(win32_event, 'IsConflict', None),
        'IsRecurring': getattr(win32_event, 'IsRecurring', None),
        # 'ItemProperties': getattr(win32_event, 'ItemProperties', None), # <COMObject <unknown>>
        'LastModificationTime': getattr(win32_event, 'LastModificationTime', None),
        'Location': getattr(win32_event, 'Location', None),
        'MarkForDownload': getattr(win32_event, 'MarkForDownload', None),
        'MeetingStatus': getattr(win32_event, 'MeetingStatus', None),
        'MeetingWorkspaceURL': getattr(win32_event, 'MeetingWorkspaceURL', None),
        'MessageClass': getattr(win32_event, 'MessageClass', None),
        'Mileage': getattr(win32_event, 'Mileage', None),
        'NoAging': getattr(win32_event, 'NoAging', None),
        'OptionalAttendees': getattr(win32_event, 'OptionalAttendees', None),
        'Organizer': getattr(win32_event, 'Organizer', None),
        'OutlookInternalVersion': getattr(win32_event, 'OutlookInternalVersion', None),
        'OutlookVersion': getattr(win32_event, 'OutlookVersion', None),
        # 'Parent': getattr(win32_event, 'Parent', None), # <COMObject <unknown>>
        # 'PropertyAccessor': getattr(win32_event, 'PropertyAccessor', None), # <COMObject <unknown>>
        # 'Recipients': getattr(win32_event, 'Recipients', None), # <COMObject <unknown>>
        'RecurrenceState': getattr(win32_event, 'RecurrenceState', None),
        'ReminderOverrideDefault': getattr(win32_event, 'ReminderOverrideDefault', None),
        'ReminderPlaySound': getattr(win32_event, 'ReminderPlaySound', None),
        'ReminderSet': getattr(win32_event, 'ReminderSet', None),
        'ReminderSoundfile': getattr(win32_event, 'ReminderSoundfile', None),
        'ReplyTime': getattr(win32_event, 'ReplyTime', None),
        'RequiredAttendees': getattr(win32_event, 'RequiredAttendees', None),
        'Resources': getattr(win32_event, 'Resources', None),
        'ResponseRequested': getattr(win32_event, 'ResponseRequested', None),
        'ResponseStatus': getattr(win32_event, 'ResponseStatus', None),
        'RTFBody': getattr(win32_event, 'RTFBody', None),
        'Saved': getattr(win32_event, 'Saved', None),
        # 'SendUsingAccount': getattr(win32_event, 'SendUsingAccount', None), # <COMObject <unknown>>
        'Sensitivity': getattr(win32_event, 'Sensitivity', None),
        # 'Session': getattr(win32_event, 'Session', None), # <COMObject <unknown>>
        'Size': getattr(win32_event, 'Size', None),
        'Start': getattr(win32_event, 'Start', None),
        'StartInStartTimeZone': getattr(win32_event, 'StartInStartTimeZone', None),
        'StartTimeZone': start_tz, # https://learn.microsoft.com/en-us/office/vba/api/outlook.timezone
        'StartUTC': getattr(win32_event, 'StartUTC', None),
        'Subject': getattr(win32_event, 'Subject', None),
        'UnRead': getattr(win32_event, 'UnRead', None),
        # 'UserProperties': getattr(win32_event, 'UserProperties', None), # <COMObject <unknown>>
    }

    return win32_event_properties

def get_win32_property_dict_full(win32_event) -> dict:

    win32_event_properties = get_win32_event_property_dict(win32_event)

    # r_pattern = None
    # if win32_event.IsRecurring:
    r_pattern = win32_event.GetRecurrencePattern()

    r_pattern_obj = None

    if r_pattern is not None:
        exceptions = r_pattern.Exceptions if r_pattern else []
        exceptions_obj_list = []
        if exceptions:
            for ex in exceptions:
                if getattr(ex, 'AppointmentItem') is not None:
                    ex_appItem_dict = get_win32_event_property_dict(ex.AppointmentItem)
                    ex_appItem_obj: AnonymousObject = AnonymousObject(ex_appItem_dict, {}, event=ex.AppointmentItem)
                    ex_dict = {
                        # 'Application': ex.Application, #<COMObject <unknown>>
                        'AppointmentItem': ex_appItem_obj,
                        # 'Class': getattr(ex, 'Class', None), # <COMObject <unknown>>
                        'Deleted': getattr(ex, 'Deleted', None),
                        # 'Parent': getattr(ex, 'Parent', None), # <COMObject <unknown>>
                        # 'Session': getattr(ex, 'Session', None), # <COMObject <unknown>>
                        }
                    ex_obj: AnonymousObject = AnonymousObject(ex_dict, {})
                    exceptions_obj_list.append(ex_obj)


        r_pattern_dict = {
            # 'Application': getattr(r_pattern, 'Application', None), # <COMObject <unknown>>
            # 'Class': getattr(r_pattern, 'Class', None), # <COMObject <unknown>>
            'DayOfMonth': getattr(r_pattern, 'DayOfMonth', None),
            'DayOfWeekMask': getattr(r_pattern, 'DayOfWeekMask', None),
            'Duration': getattr(r_pattern, 'Duration', None),
            'EndTime': getattr(r_pattern, 'EndTime', None),
            'Exceptions': exceptions_obj_list,
            'Instance': getattr(r_pattern, 'Instance', None),
            'Interval': getattr(r_pattern, 'Interval', None),
            'MonthOfYear': getattr(r_pattern, 'MonthOfYear', None),
            'NoEndDate': getattr(r_pattern, 'NoEndDate', None),
            'Occurrences': getattr(r_pattern, 'Occurrences', None),
            # 'Parent': getattr(r_pattern, 'Parent', None), # <COMObject <unknown>>
            'PatternEndDate': getattr(r_pattern, 'PatternEndDate', None),
            'PatternStartDate': getattr(r_pattern, 'PatternStartDate', None),
            'RecurrenceType': getattr(r_pattern, 'RecurrenceType', None),
            'Regenerate': getattr(r_pattern, 'Regenerate', None),
            # 'Session': getattr(r_pattern, 'Session', None), # <COMObject <unknown>>
            'StartDate': getattr(r_pattern, 'StartTime', None),
        }

        r_pattern_obj: AnonymousObject = AnonymousObject(r_pattern_dict, {})

    win32_event_properties['RecurrencePattern'] = r_pattern_obj

    return win32_event_properties

def make_anonymous_event(win32_event) -> AnonymousObject:
    props = get_win32_property_dict_full(win32_event)
    ae = AnonymousObject(props, {}, event=win32_event)
    return ae