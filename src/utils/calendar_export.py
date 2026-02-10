"""
Calendar export functionality for power outage schedules.

Provides functionality to export power outage schedules to iCalendar format
with automatic reminders.
"""

from datetime import datetime, timedelta
from typing import Dict, Tuple, Any
from icalendar import Calendar, Event, Alarm
import pytz


class CalendarExporter:
    """Handles exporting power outage schedules to iCalendar format."""

    TIMEZONE = 'Europe/Kiev'
    CALENDAR_NAME_TEMPLATE = 'Графік відключень - {queue}'
    EVENT_SUMMARY_TEMPLATE = '[!] Відключення електроенергії - {queue}'
    EVENT_DESCRIPTION_TEMPLATE = (
        'Планове відключення електроенергії\n'
        'Черга: {queue}\n'
        'Час: {start_time} - {end_time}'
    )
    ALARM_DESCRIPTION_TEMPLATE = '[!] Відключення через {minutes} хвилин - {queue}'

    @staticmethod
    def export_to_ics(data: Dict, date_obj: datetime, filename: str) -> Tuple[bool, str]:
        """
        Export power outage schedule to iCalendar file.

        Args:
            data: Dictionary containing queue name and outage information
            date_obj: Date object for the schedule
            filename: Path to save the .ics file

        Returns:
            Tuple of (success: bool, message: str)
        """
        queue_name = data.get('queue', '')
        outages = data.get('outages', [])

        if not outages:
            return False, "Немає відключень для експорту"

        cal = CalendarExporter._create_calendar(queue_name)
        tz = pytz.timezone(CalendarExporter.TIMEZONE)

        for outage in outages:
            event = CalendarExporter._create_event(
                outage,
                date_obj,
                queue_name,
                tz
            )
            cal.add_component(event)

        try:
            with open(filename, 'wb') as f:
                f.write(cal.to_ical())
            return True, filename
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _create_calendar(queue_name: str) -> Calendar:
        """
        Create base calendar object with metadata.

        Args:
            queue_name: Name of the power outage queue

        Returns:
            Configured Calendar object
        """
        cal = Calendar()
        cal.add('prodid', '-//Schedule Analyzer//UA')
        cal.add('version', '2.0')
        cal.add('x-wr-calname', CalendarExporter.CALENDAR_NAME_TEMPLATE.format(queue=queue_name))
        cal.add('x-wr-timezone', CalendarExporter.TIMEZONE)
        return cal

    @staticmethod
    def _create_event(outage: Dict[str, str],
                     date_obj: datetime,
                     queue_name: str,
                     tz: Any) -> Event:
        """
        Create calendar event for a single outage period.

        Args:
            outage: Dictionary with 'start' and 'end' time strings
            date_obj: Date of the outage
            queue_name: Name of the power outage queue
            tz: Timezone object

        Returns:
            Configured Event object with alarms
        """
        start_time_str = outage['start']
        end_time_str = outage['end']

        start_dt = CalendarExporter._parse_datetime(start_time_str, date_obj, tz)
        end_dt = CalendarExporter._parse_datetime(end_time_str, date_obj, tz, is_end=True)

        event = Event()
        event.add('summary', CalendarExporter.EVENT_SUMMARY_TEMPLATE.format(queue=queue_name))
        event.add('dtstart', start_dt)
        event.add('dtend', end_dt)
        event.add('description', CalendarExporter.EVENT_DESCRIPTION_TEMPLATE.format(
            queue=queue_name,
            start_time=start_time_str,
            end_time=end_time_str
        ))
        event.add('location', 'Україна')
        event.add('status', 'CONFIRMED')

        event.add_component(CalendarExporter._create_alarm(15, queue_name))
        event.add_component(CalendarExporter._create_alarm(5, queue_name))

        return event

    @staticmethod
    def _parse_datetime(time_str: str,
                       date_obj: datetime,
                       tz: Any,
                       is_end: bool = False) -> datetime:
        """
        Parse time string and combine with date.

        Args:
            time_str: Time string in HH:MM format
            date_obj: Date object
            tz: Timezone object
            is_end: Whether this is an end time (handles 24:00 case)

        Returns:
            Localized datetime object
        """
        hour, minute = map(int, time_str.split(':'))

        if is_end and hour == 24 and minute == 0:
            return tz.localize(datetime.combine(
                date_obj.date() + timedelta(days=1),
                datetime.min.time()
            ))

        return tz.localize(datetime.combine(
            date_obj.date(),
            datetime.min.time().replace(hour=hour, minute=minute)
        ))

    @staticmethod
    def _create_alarm(minutes_before: int, queue_name: str) -> Alarm:
        """
        Create alarm/reminder for event.

        Args:
            minutes_before: Minutes before event to trigger alarm
            queue_name: Name of the queue for alarm description

        Returns:
            Configured Alarm object
        """
        alarm = Alarm()
        alarm.add('action', 'DISPLAY')
        alarm.add('description', CalendarExporter.ALARM_DESCRIPTION_TEMPLATE.format(
            minutes=minutes_before,
            queue=queue_name
        ))
        alarm.add('trigger', timedelta(minutes=-minutes_before))
        return alarm

