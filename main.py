import json
import re
from datetime import datetime, timedelta, timezone
from icalendar import Calendar, Event, vText
import uuid

def parse_time(time_str):
    """提取Z（UTC）的时间，格式如 '2000-01-01T09:00:00Z'"""
    # 获取 '09:00:00'
    t_str = time_str.split('T')[1].replace('Z', '')
    return datetime.strptime(t_str, "%H:%M:%S").time()

def parse_week_pattern(pattern_str):
    """解析周数模式为范围或独立周。"""
    parts = [p.strip() for p in pattern_str.split(',')]
    parsed = []
    for p in parts:
        if '-' in p:
            s, e = p.split('-')
            parsed.append({'type': 'range', 'start': int(s), 'end': int(e)})
        else:
            parsed.append({'type': 'single', 'week': int(p)})
    return parsed

def get_event_date(week1_monday, week_num, scheduled_day):
    """
    根据第几周和星期几获取具体的日期
    week1_monday: 第一周星期一的日期
    week_num: 课程发生在第几周 (1-indexed)
    scheduled_day: 0是周一, ..., 6是周日
    """
    days_offset = (week_num - 1) * 7 + int(scheduled_day)
    return week1_monday + timedelta(days=days_offset)

def map_weekday(scheduled_day):
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    return days[int(scheduled_day)]

def convert_to_ics(json_path, ics_path, week1_start_date_str):
    # 解析第一周的星期一 (格式 YYYY-MM-DD)
    week1_monday = datetime.strptime(week1_start_date_str, "%Y-%m-%d").date()
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    cal = Calendar()
    cal.add('prodid', '-//University Schedule to ICS Converter//EN')
    cal.add('version', '2.0')

    for item in data:
        name = item.get("name", "Unknown Course")
        location = item.get("location", "")
        staff = item.get("staff", "Unknown")
        activity_type = item.get("activityType", "Other")
        week_pattern_str = item.get("weekPattern", "")
        scheduled_day = int(item.get("scheduledDay", 0))
        
        start_time_obj = parse_time(item.get("startTime"))
        end_time_obj = parse_time(item.get("endTime"))
        
        # 详细描述，包含任课教师、周数、星期几
        description = f"Type: {activity_type}\nStaff: {staff}\nWeek Pattern: {week_pattern_str}\nDay: {map_weekday(scheduled_day)}"
        
        patterns = parse_week_pattern(week_pattern_str)
        
        for pattern in patterns:
            # 对于第一项，我们创建一个事件。如果是范围，我们将添加RRULE (Series循环)
            start_week = pattern['start'] if pattern['type'] == 'range' else pattern['week']
            
            event_start_date = get_event_date(week1_monday, start_week, scheduled_day)
            
            # 由于原时间含Z代表UTC，所以应用UTC时区
            dt_start = datetime.combine(event_start_date, start_time_obj, tzinfo=timezone.utc)
            dt_end = datetime.combine(event_start_date, end_time_obj, tzinfo=timezone.utc)
            
            event = Event()
            event.add('uid', str(uuid.uuid4()) + '@schedule.local')
            event.add('summary', name)
            event.add('dtstart', dt_start)
            event.add('dtend', dt_end)
            event.add('location', location)
            event.add('description', description)
            
            # 使用列表为同一个事件添加两个 Category（确保颜色标签在前）
            color_mapping = {
                "Lecture": "Purple Category",
                "Lab": "Green Category",
                "Seminar": "Blue Category",
                "Practical": "Orange Category",
                "Tutorial": "Pink Category",
                "Comp.Lab": "Dark Blue Category"
            }
            categories = []
            if activity_type in color_mapping:
                categories.append(color_mapping[activity_type])
            categories.append(activity_type)
                
            event.add('categories', categories)
            
            if pattern['type'] == 'range':
                # 计算循环次数
                count = pattern['end'] - pattern['start'] + 1
                event.add('rrule', {'freq': 'weekly', 'count': count})
                
            cal.add_component(event)

    # 写入ICS文件
    with open(ics_path, 'wb') as f:
        f.write(cal.to_ical())
        
    print(f"✅ 转换完成！已将 {json_path} 保存为 {ics_path} 。")
    print("您可以将其直接拖入或者导入 Microsoft Outlook。")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Convert time.json to Outlook compatible ICS file.")
    parser.add_argument('--input', type=str, default='time.json', help="Path to input json file.")
    parser.add_argument('--output', type=str, default='schedule.ics', help="Path to output ics file.")
    
    print("=== JSON to Outlook ICS Schedule Converter ===")
    user_date = input("📅 请输入[第一周星期一]的日期 (格式如 2024-02-26): ")
    
    try:
        # validate date format
        datetime.strptime(user_date, "%Y-%m-%d")
        convert_to_ics('time.json', 'schedule.ics', user_date)
    except ValueError:
        print("❌ 错误：日期格式不正确，请确保使用 YYYY-MM-DD 的格式。")
    except FileNotFoundError:
        print("❌ 错误：找不到 time.json，请确保文件位于当前目录下。")
    except Exception as e:
        print(f"❌ 发生意外错误：{str(e)}")
