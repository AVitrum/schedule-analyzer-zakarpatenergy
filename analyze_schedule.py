from PIL import Image

START_X = 128
END_X = 1270
PIXELS_PER_HALF_HOUR = 24
TOTAL_HALF_HOURS = 48

COLORS = {
    (254, 255, 3): "Черга 1",
    (146, 210, 74): "Черга 2", 
    (253, 193, 0): "Черга 3",
    (0, 178, 237): "Черга 4",
    (236, 126, 49): "Черга 5",
    (179, 126, 218): "Черга 6"
}

def rgb_distance(c1, c2):
    return sum((a - b) ** 2 for a, b in zip(c1, c2)) ** 0.5

def is_outage_color(pixel, threshold=50):
    for color in COLORS.keys():
        if rgb_distance(pixel[:3], color) < threshold:
            return True
    return False

def time_to_string(half_hour_index):
    hours = half_hour_index // 2
    minutes = 30 if half_hour_index % 2 else 0
    if hours == 24:
        return "24:00"
    return f"{hours:02d}:{minutes:02d}"

def analyze_row(img, y_coord):
    outages = []
    in_outage = False
    start_index = None
    
    for half_hour in range(TOTAL_HALF_HOURS):
        x = START_X + half_hour * PIXELS_PER_HALF_HOUR + PIXELS_PER_HALF_HOUR // 2
        pixel = img.getpixel((x, y_coord))
        
        has_outage = is_outage_color(pixel)
        
        if has_outage and not in_outage:
            in_outage = True
            start_index = half_hour
        elif not has_outage and in_outage:
            in_outage = False
            outages.append((start_index, half_hour))
        
    if in_outage:
        outages.append((start_index, TOTAL_HALF_HOURS))
    
    return outages

def main():
    img = Image.open("img.png")
    
    queues = {
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
    
    print("Прогнозовані години відключення електроенергії на 12.01.2026р.")
    print("=" * 60)
    
    for queue_name, y_coord in queues.items():
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
