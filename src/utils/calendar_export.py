from datetime import datetime, timedelta
from icalendar import Calendar, Event, Alarm
import pytz


class CalendarExporter:
    
    @staticmethod
    def export_to_ics(data, date_obj, filename):
        queue_name = data.get('queue', '')
        outages = data.get('outages', [])

        if not outages:
            return False, "Немає відключень для експорту"

        cal = Calendar()
        cal.add('prodid', '-//Schedule Analyzer//UA')
        cal.add('version', '2.0')
        cal.add('x-wr-calname', f'Графік відключень - {queue_name}')
        cal.add('x-wr-timezone', 'Europe/Kiev')

        tz = pytz.timezone('Europe/Kiev')

        for outage in outages:
            start_time_str = outage['start']
            end_time_str = outage['end']

            start_hour, start_min = map(int, start_time_str.split(':'))
            end_hour, end_min = map(int, end_time_str.split(':'))

            start_dt = tz.localize(datetime.combine(
                date_obj.date(),
                datetime.min.time().replace(hour=start_hour, minute=start_min)
            ))

            if end_hour == 24 and end_min == 0:
                end_dt = tz.localize(datetime.combine(
                    date_obj.date() + timedelta(days=1),
                    datetime.min.time()
                ))
            else:
                end_dt = tz.localize(datetime.combine(
                    date_obj.date(),
                    datetime.min.time().replace(hour=end_hour, minute=end_min)
                ))

            event = Event()
            event.add('summary', f'[!] Відключення електроенергії - {queue_name}')
            event.add('dtstart', start_dt)
            event.add('dtend', end_dt)
            event.add('description',
                     f'Планове відключення електроенергії\n'
                     f'Черга: {queue_name}\n'
                     f'Час: {start_time_str} - {end_time_str}')
            event.add('location', 'Україна')
            event.add('status', 'CONFIRMED')

            alarm_15 = Alarm()
            alarm_15.add('action', 'DISPLAY')
            alarm_15.add('description', f'[!] Відключення через 15 хвилин - {queue_name}')
            alarm_15.add('trigger', timedelta(minutes=-15))
            event.add_component(alarm_15)

            alarm_5 = Alarm()
            alarm_5.add('action', 'DISPLAY')
            alarm_5.add('description', f'[!] Відключення через 5 хвилин - {queue_name}')
            alarm_5.add('trigger', timedelta(minutes=-5))
            event.add_component(alarm_5)

            cal.add_component(event)

        try:
            with open(filename, 'wb') as f:
                f.write(cal.to_ical())
            return True, filename
        except Exception as e:
            return False, str(e)
