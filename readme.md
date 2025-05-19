# 📅 Google Calendar to Google Sheets Time Tracker

This Google Apps Script project syncs events from a Google Calendar to a Google Sheet for time tracking. It extracts essential details like event name, start and end time, and automatically calculates durations in both time and numeric formats.

---

## 🔧 Features

- 🗓️ **Sync Events:** Fetches all events from a specified calendar.
- 📥 **Dynamic Range:** Automatically determines the earliest start and latest end date of events.
- 🧹 **Auto Cleanup:** Removes outdated entries before syncing.
- ⏱️ **Auto Duration Calculation:**
  - Time duration (formatted: hh:mm)
  - Numeric duration (in decimal hours)
- 📊 **Minimal and Clean Data Output**

---

## 📄 Data Columns

| Column             | Description                                 |
|--------------------|---------------------------------------------|
| `Calendar Id`      | ID of the calendar the event belongs to     |
| `Event Name`       | Title of the calendar event                 |
| `Begin`            | Start date and time of the event            |
| `End`              | End date and time of the event              |
| `Time Duration`    | Calculated duration in `hh:mm` format       |
| `Numeric Duration` | Duration in decimal hours (e.g., 1.5 hours) |
| `Name`             | Name of the calendar                        |

---

## 🚀 How to Use

1. Open your **Google Spreadsheet**.
2. Navigate to `Extensions > Apps Script`.
3. Paste the complete script into the editor.
4. Save and reload the spreadsheet.
5. A new menu will appear: **`Custom Menu > Sync Calendar Events`**.
6. Go to the `All Hours` sheet and enter a valid **Google Calendar ID** in cell `B2`.
7. Click the **"Sync Calendar Events"** option in the menu to fetch and record the events.

---

## 🔑 Calendar ID Format

- For your own calendar, it's usually your Gmail address: `yourname@gmail.com`.
- For shared or group calendars, open **Google Calendar > Settings > Calendar Settings**, and copy the **Calendar ID**.

---

## ✅ Use Cases

- Tracking personal or employee working hours.
- Client billing reports based on event durations.
- Monitoring shared calendar activities.

---

## 📝 Notes

- The script retrieves all events between the years 2000 and 2100.
- Duration formulas are automatically set using Google Sheets formula logic.
- Only rows matching the entered calendar ID are updated.

---

## 📌 Sheet Requirements

Make sure your Google Sheet contains a sheet named `All Hours`, with the following column headers starting at row 1:

---

## 📬 Feedback

Have suggestions or want to add more features (e.g., date filters, UI prompts)? Feel free to open an issue or contribute!
