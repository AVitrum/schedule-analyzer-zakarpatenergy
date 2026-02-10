"""
Schedule analyzer module for power outage detection from images.

This module provides functionality to analyze power outage schedule images
by detecting colored regions and converting them to time intervals.
"""

from typing import Tuple, List
from PIL import Image


class ScheduleConfig:
    """Configuration constants for schedule image analysis."""

    TARGET_WIDTH: int = 1280
    TARGET_HEIGHT: int = 335
    START_X: int = 128
    END_X: int = 1270
    PIXELS_PER_HALF_HOUR: int = 24
    TOTAL_HALF_HOURS: int = 48
    COLOR_THRESHOLD: int = 50

    QUEUE_COLORS = {
        (254, 255, 3): "Черга 1",
        (146, 210, 74): "Черга 2",
        (253, 193, 0): "Черга 3",
        (0, 178, 237): "Черга 4",
        (236, 126, 49): "Черга 5",
        (179, 126, 218): "Черга 6"
    }

    QUEUE_COORDINATES = {
        "Черга 1-1": 90,
        "Черга 1-2": 109,
        "Черга 2-1": 127,
        "Черга 2-2": 146,
        "Черга 3-1": 170,
        "Черга 3-2": 190,
        "Черга 4-1": 214,
        "Черга 4-2": 233,
        "Черга 5-1": 255,
        "Черга 5-2": 275,
        "Черга 6-1": 298,
        "Черга 6-2": 315,
    }


def resize_image(img: Image.Image) -> Image.Image:
    """
    Resize image to target dimensions if necessary.

    Args:
        img: PIL Image object to resize

    Returns:
        Resized PIL Image object
    """
    if img.size != (ScheduleConfig.TARGET_WIDTH, ScheduleConfig.TARGET_HEIGHT):
        img = img.resize(
            (ScheduleConfig.TARGET_WIDTH, ScheduleConfig.TARGET_HEIGHT),
            Image.Resampling.LANCZOS
        )
    return img


def calculate_rgb_distance(color1: Tuple[int, int, int],
                           color2: Tuple[int, int, int]) -> float:
    """
    Calculate Euclidean distance between two RGB colors.

    Args:
        color1: First RGB color tuple
        color2: Second RGB color tuple

    Returns:
        Euclidean distance between colors
    """
    return sum((a - b) ** 2 for a, b in zip(color1, color2)) ** 0.5


def is_outage_color(pixel: Tuple[int, ...],
                    threshold: int = ScheduleConfig.COLOR_THRESHOLD) -> bool:
    """
    Check if pixel color indicates a power outage.

    Args:
        pixel: RGB pixel tuple from image
        threshold: Maximum distance to consider colors matching

    Returns:
        True if pixel matches any outage color
    """
    for color in ScheduleConfig.QUEUE_COLORS.keys():
        if calculate_rgb_distance(pixel[:3], color) < threshold:
            return True
    return False


def time_to_string(half_hour_index: int) -> str:
    """
    Convert half-hour index to time string.

    Args:
        half_hour_index: Index representing 30-minute interval (0-48)

    Returns:
        Formatted time string (HH:MM)
    """
    hours = half_hour_index // 2
    minutes = 30 if half_hour_index % 2 else 0

    if hours == 24:
        return "24:00"

    return f"{hours:02d}:{minutes:02d}"


def analyze_row(img: Image.Image, y_coord: int) -> List[Tuple[int, int]]:
    """
    Analyze a horizontal row in the schedule image for outage periods.

    Args:
        img: PIL Image object to analyze
        y_coord: Y coordinate of the row to analyze

    Returns:
        List of tuples containing (start_index, end_index) for each outage period
    """
    outages = []
    in_outage = False
    start_index = None

    for half_hour in range(ScheduleConfig.TOTAL_HALF_HOURS):
        x = (ScheduleConfig.START_X +
             half_hour * ScheduleConfig.PIXELS_PER_HALF_HOUR +
             ScheduleConfig.PIXELS_PER_HALF_HOUR // 2)
        pixel = img.getpixel((x, y_coord))
        has_outage = is_outage_color(pixel)

        if has_outage and not in_outage:
            in_outage = True
            start_index = half_hour
        elif not has_outage and in_outage:
            in_outage = False
            outages.append((start_index, half_hour))

    if in_outage:
        outages.append((start_index, ScheduleConfig.TOTAL_HALF_HOURS))

    return outages


def main() -> None:
    """
    Main function for standalone script execution.
    Analyzes schedule image and prints results for all queues.
    """
    img = Image.open("img.png")

    print("Прогнозовані години відключення електроенергії на 12.01.2026р.")
    print("=" * 60)

    for queue_name, y_coord in ScheduleConfig.QUEUE_COORDINATES.items():
        outages = analyze_row(img, y_coord)

        if outages:
            print(f"\n{queue_name}:")
            for start, end in outages:
                start_time = time_to_string(start)
                end_time = time_to_string(end)
                print(f"  {start_time} - {end_time}")
        else:
            print(f"\n{queue_name}: Без відключень")


if __name__ == "__main__":
    main()
