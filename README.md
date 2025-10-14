
# Cali: XLSX to ICS Converter 📅

Cali is a robust Python utility designed to effortlessly convert event data from a standard Excel spreadsheet (`.xlsx`) into the iCalendar format (`.ics`), which is compatible with most modern calendar applications (Google Calendar, Outlook, Apple Calendar, etc.).

It features template recognition, allowing it to parse sheets with basic or complex schedules, including recurring events.

## Features ✨

* **Automatic Template Recognition:** Distinguishes between **Basic** (simple dates/times) and **Medium** (separate date/time columns + recurrence) spreadsheet formats.
* **Recurrence Rule (RRULE) Generation:** Automatically creates iCalendar recurrence rules for events with a `FREQUENCY` column.
* **Column Aliasing:** Maps common column name variations (e.g., "Title," "Event Title," "Name") to a single canonical field.
* **Error Logging:** Generates a detailed `.log` file that tracks which sheets were processed and lists any rows skipped due to data errors (e.g., invalid dates or recurrence values).

## Installation 💻

Cali requires Python 3.7+ and several key dependencies for handling Excel and iCalendar files.

### Prerequisites

1. **Python 3.x**

### Install Dependencies

Use `pip` to install the required libraries:

**Bash**

```
pip install pandas openpyxl icalendar pytz
```

## Usage 🚀

The application is run via the Command Line Interface (`cli.py`). It requires two arguments: the input Excel file and the desired output ICS file name.

### Command

**Bash**

```
python cli.py <input_file.xlsx> <output_file.ics>
```

### Example

To convert a file named `my-schedule.xlsx` and save the output as `my-calendar.ics`:

**Bash**

```
python cli.py my-schedule.xlsx my-calendar.ics
```

This command will create two files:

1. `my-calendar.ics` (The final calendar file)
2. `my-calendar.log` (A summary of the process and any data warnings)

## Spreadsheet Templates 📑

Cali recognizes and processes columns based on the fields found in your spreadsheet. Headers are case-insensitive.

### 1. Medium Template (Recommended for Recurrence)

This template is required if you need events to repeat (`RRULE`). It separates dates and times and includes a dedicated frequency column.

| Canonical Name         | Required | Alias Examples                                    | Data Format                                                         |
| ---------------------- | -------- | ------------------------------------------------- | ------------------------------------------------------------------- |
| **TITLE**        | Yes      | `Title`,`Event Title`,`Name`                | Text                                                                |
| **START_DATE**   | Yes      | `Start Date`,`Date Start`                     | Standard Date (e.g., 1/10/2025)                                     |
| **END_DATE**     | Yes      | `End Date`,`Date End`                         | Standard Date                                                       |
| **START_TIME**   | Yes      | `Start Time`,`Time Start`                     | Standard Time (e.g., 9:00 AM, 21:00:00)                             |
| **END_TIME**     | Yes      | `End Time`,`Time End`                         | Standard Time                                                       |
| **FREQUENCY**    | No       | `Frequency`,`Recurrence`,`Repeats`,`Days` | **Comma-separated days (e.g.,`Monday, Wednesday, Friday`)** |
| **LOCATION**     | No       | `Location`,`Place`,`Address`                | Text                                                                |
| **OWNER**        | No       | `Owner`,`Organizer`                           | Email address (e.g.,`user@domain.com`)                            |
| **PARTICIPANTS** | No       | `Participants`,`Attendees`                    | Comma-separated emails (e.g.,`a@a.com, b@b.com`)                  |
| **LINK**         | No       | `Link`,`URL`,`Zoom`                         | URL                                                                 |

### 2. Basic Template

Used for simple, one-off events where `START_TIME` and `END_TIME` contain full datetime stamps.

| Canonical Name       | Required | Alias Examples                     |
| -------------------- | -------- | ---------------------------------- |
| **TITLE**      | Yes      | `Title`,`Event Title`,`Name` |
| **START_TIME** | Yes      | `StartTime`,`Start Time`       |
| **END_TIME**   | Yes      | `EndTime`,`End Time`           |
