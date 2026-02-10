"""
Time calculation utilities for power outage schedule analysis.

Provides functions for calculating total outage duration and sorting queues.
"""

from typing import Dict, List, Tuple


def calculate_total_time(outages: List[Dict[str, str]]) -> Tuple[int, int, int]:
    """
    Calculate total duration of power outages.

    Args:
        outages: List of outage dictionaries with 'start' and 'end' time strings

    Returns:
        Tuple of (hours, minutes, total_minutes)
    """
    total_minutes = 0

    for outage in outages:
        start_time = outage['start']
        end_time = outage['end']
        
        start_hour, start_min = map(int, start_time.split(':'))
        end_hour, end_min = map(int, end_time.split(':'))
        
        start_total_min = start_hour * 60 + start_min
        end_total_min = end_hour * 60 + end_min
        
        if end_total_min == 0:
            end_total_min = 24 * 60
        
        duration = end_total_min - start_total_min
        total_minutes += duration
    
    hours = total_minutes // 60
    minutes = total_minutes % 60

    return hours, minutes, total_minutes


def queue_sort_key(item: Dict) -> Tuple[int, int]:
    """
    Generate sort key for queue comparison results.

    Args:
        item: Dictionary containing 'queue' field with queue name

    Returns:
        Tuple of (main_number, sub_number) for sorting
    """
    queue_name = item['queue']
    parts = queue_name.split()

    if len(parts) >= 2:
        try:
            main_num = int(parts[1].split('-')[0])
            sub_num = int(parts[1].split('-')[1])
            return (main_num, sub_num)
        except (ValueError, IndexError):
            return (99, 99)

    return (99, 99)


