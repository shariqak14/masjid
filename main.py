from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from datetime import datetime, timedelta, date
from tkinter import *
from tkinter.ttk import *
from calendar import Calendar, day_name

import requests

MONTH = "June"
YEAR = "2021"

lunar_months = [
    "Muharram",
    "Safar",
    "Rabi' al-awwal",
    "Rabi' al-thani",
    "Jumada al-awwal",
    "Jumada al-thani",
    "Rajab",
    "Sha'aban",
    "Ramadan",
    "Shawwal",
    "Dhu al-Qi'dah",
    "Dhu al-Hijjah",
]

def get_prayer_time(month):
    apiURL = "http://www.islamicfinder.us/index.php/api/prayer_times"

    apiParams = {
        "country": "US",
        "zipcode": "06029",
        "juristic": 1,
        "time_format": 2,
        "show_entire_month": 1,
        "date": month,
    }

    response = requests.get(apiURL, params=apiParams)

    prayerTimes = list(response.json()["results"].values())

    return prayerTimes


def get_lunar_date(year, month, day):
    apiURL = "http://www.islamicfinder.us/index.php/api/calendar"

    apiParams = {
        "year": year,
        "month": month,
        "day": day,
    }

    response = requests.get(apiURL, params=apiParams)

    _, month, day = response.json()["to"].split("-")

    return int(month), int(day)


def format_cell(
    cell, center=True, bold=False, italic=False, font_size=12, font="Arial"
):
    if center:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.paragraphs[0].runs[0].font.bold = bold
    cell.paragraphs[0].runs[0].italic = italic
    cell.paragraphs[0].runs[0].font.size = Pt(font_size)
    cell.paragraphs[0].runs[0].font.name = font


####################################################################################
# Edit Document
####################################################################################

doc = Document("June_2021.docx")
table = doc.tables[0]

month_heading = table.cell(0, 10)
month_heading.text = MONTH.upper() + " " + YEAR
format_cell(month_heading, bold=True, font_size=15, font="Times New Roman")

month_row_title = table.cell(1, 2)
month_row_title.text = MONTH + " " + YEAR + " CE"
format_cell(month_row_title, bold=True, font_size=11, font="Times New Roman")

####################################################################################
# Insert Date
####################################################################################

uniq_lunar_months = []

for j in range(3, 33):
    my_date = date(int(YEAR), int("06"), j - 2)
    name_of_day = str(day_name[my_date.weekday()])[:3]

    day_col = table.cell(j, 2)
    day_col.text = name_of_day
    format_cell(day_col, bold=True)

    day_no_col = table.cell(j, 3)
    day_no_col.text = str(j - 2)
    format_cell(day_no_col, bold=True)

    lunar_month, lunary_day = get_lunar_date(int(YEAR), int("06"), j - 2)

    lunary_day_col = table.cell(j, 4)

    if j == 3:
        lunary_day_col.text = lunar_months[lunar_month - 1] + " " + str(lunary_day)
        format_cell(lunary_day_col, bold=True, italic=True, font_size=10)
    elif lunary_day == 1:
        lunary_day_col.text = lunar_months[lunar_month - 1] + " " + str(lunary_day) + "*"
        format_cell(lunary_day_col, bold=True, italic=True, font_size=10)
    else:
        lunary_day_col.text = str(lunary_day)
        format_cell(lunary_day_col)

    if lunar_month not in uniq_lunar_months:
        uniq_lunar_months.append(lunar_month)

month_1, month_2 = uniq_lunar_months[0], uniq_lunar_months[1]
lunar_col_title = table.cell(2, 4)
lunar_col_title.text = lunar_months[month_1 - 1] + "\n" + lunar_months[month_1 - 2] + "\n" + "1442 AH"
format_cell(lunar_col_title, bold=True, italic=True, font_size=10)

####################################################################################
# Insert Prayer Times
####################################################################################

prayer_times = get_prayer_time(month=MONTH)

num_of_rows = len(prayer_times)

prayers = {"Fajr": 5, "Duha": 8, "Dhuhr": 9, "Asr": 11, "Maghrib": 13, "Isha": 15}

for prayer, col_no in prayers.items():
    times = list(map(lambda x: x[prayer], prayer_times))

    for j in range(3, 33):
        pr_time = table.cell(j, col_no)
        pr_time.text = times[j - 3]  # prayer[:2]
        format_cell(pr_time)

        if prayer == "Maghrib":
            iqamah_time = datetime.strptime(times[j - 3], "%H:%M") + timedelta(
                minutes=10
            )
            iqamah_time = iqamah_time.strftime("%H:%M")
            iqamah_time = iqamah_time if iqamah_time[0] != "0" else iqamah_time[1:]

            iqamah_cell = table.cell(j, col_no + 1)
            iqamah_cell.text = iqamah_time
            format_cell(iqamah_cell, bold=True)

doc.save("grid2.docx")

# window = Tk()

# window.title("Masjid Calendar")

# window.geometry('350x200')

# combo = Combobox(window)

# combo['values'] = [
#     "January",
#     "February",
#     "March",
#     "April",
#     "May",
#     "June",
#     "July",
#     "August",
#     "September",
#     "October",
#     "November",
#     "December",
# ]

# currentMonth = int(datetime.now().strftime('%m'))
# print(currentMonth)

# combo.current(currentMonth) # set the selected item

# combo.grid(column=0, row=0)

# window.mainloop()
