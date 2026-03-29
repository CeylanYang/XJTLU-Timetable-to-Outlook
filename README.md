# XJTLU timetable to Outlook ICS Converter

[**English**](./README.md) | [**中文简体**](./README_zh.md)

This is a lightweight Python tool designed to convert your university course schedules (stored in JSON format) into `.ics` (iCalendar) files perfectly optimized for Microsoft Outlook.

## ✨ Features
* **Outlook Series Support**: Automatically parses continuous week patterns (e.g., "1-6") and creates native Outlook Recurring Series (RRULE).
* **Auto-Coloring via Categories**: Intelligently assigns Outlook's default color categories alongside your class types (e.g., maps `Lecture` to `Purple Category`), so your calendar looks colorful and organized instantly!
* **Detailed Descriptions**: Populates the event body with Staff/Professor names, Week Patterns, and Day of Week details.
* **UTC Timezone Aware**: Accurately handles `Z` (UTC) time conversions.
* **Custom Start Date**: Prompts for your university's specific "Week 1 Monday" date during execution and figures everything else out automatically.

---

## 🚀 How to Use

### 1. Prerequisites
Ensure you have Python installed. Then, install the required `icalendar` library:
```bash
pip install icalendar
```

### 2. Prepare your Data
1. Open your E-Bridge timetable webpage.
2. Press `F12` to open Developer Tools (taking Chrome as an example).
3. Navigate to the **Network** tab, filter by Fetch/XHR, and find the specific data request as shown in the image below.
   
   ![How to find JSON data](readme/41dc3f402a7c498c567a8480c5c3905b.jpg)

4. Copy the entire raw JSON response.
5. Replace all contents inside the `time.json` file in this repository with your copied data.

### 3. Run the Script
Open your terminal and execute:
```bash
python main.py
```

### 4. Provide the Start Date
The console will prompt you to enter the starting date of your semester ("Week 1 Monday"):
```text
=== JSON to Outlook ICS Schedule Converter ===
📅 请输入[第一周星期一]的日期 (格式如 2024-02-26): 2024-03-04
```
Type your date in `YYYY-MM-DD` format and hit **Enter**.

### 5. Import to Outlook
The script will instantly generate a `schedule.ics` file in the same directory. You can import it using either of the following methods:

**Method A: Outlook Desktop App**
Simply double-click the `schedule.ics` file to open it with Microsoft Outlook, or drag and drop it directly into your Outlook Calendar view!

**Method B: Outlook Web Version**
1. Browse to [Outlook Live Calendar](https://outlook.live.com/calendar/view/week).
2. Click **Add calendar** on the left panel.
3. Select **Upload from file**, click **Browse** to choose your generated `schedule.ics`.
4. Choose the calendar you want to import into, and hit **Import**.

![Outlook Screenshot 1](readme/162a6bf1a1afb6661c1b22547f8c64fe.jpg)
![Outlook Screenshot 2](readme/sample.png)

---

## 📅 Color Mapping logic
The script automatically assigns double tags (Color Tag + Feature Tag) in this exact order to populate colors in Outlook:
- **Lecture** $\rightarrow$ `Purple Category` & `Lecture`
- **Lab** $\rightarrow$ `Green Category` & `Lab`
- **Seminar** $\rightarrow$ `Blue Category` & `Seminar`
- **Comp.Lab** $\rightarrow$ `Dark Blue Category` & `Comp.Lab`
- **Tutorial** $\rightarrow$ `Pink Category` & `Tutorial`
- **Practical** $\rightarrow$ `Orange Category` & `Practical`

---