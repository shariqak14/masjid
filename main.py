from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from datetime import datetime, timedelta, date
from calendar import Calendar, day_name, monthrange

from tkinter import *
from tkinter.ttk import *

from api import get_prayer_time, get_lunar_date
from constants import LUNAR_MONTHS_ENG, LUNAR_MONTHS_ARABIC, GREG_MONTHS

MONTH = "February"
YEAR = "2022"

def number_of_days_in_month(year=2021, month=2):
    return monthrange(year, month)[1]

def format_cell(
    cell, center=True, bold=False, italic=False, font_size=12, font="Arial"
):
    run = cell.paragraphs[0].runs[0]

    run.font.bold = bold
    run.font.size = Pt(font_size)
    run.font.name = font

    run.italic = italic

    if center: 
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

####################################################################################
# Edit Document
####################################################################################

NUM_DAYS = number_of_days_in_month(int(YEAR), GREG_MONTHS[MONTH])
doc = Document("templates/" + str(NUM_DAYS) + "_days.docx")

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

for j in range(3, NUM_DAYS + 3):
    my_date = date(int(YEAR), GREG_MONTHS[MONTH], j - 2)
    name_of_day = str(day_name[my_date.weekday()])[:3]

    day_col = table.cell(j, 2)
    day_col.text = name_of_day
    format_cell(day_col, bold=True)

    day_no_col = table.cell(j, 3)
    day_no_col.text = str(j - 2)
    format_cell(day_no_col, bold=True)

    lunar_month, lunary_day = get_lunar_date(int(YEAR), GREG_MONTHS[MONTH], j - 2)

    lunary_day_col = table.cell(j, 4)

    if j == 3:
        lunary_day_col.text = LUNAR_MONTHS_ENG[lunar_month - 1] + " " + str(lunary_day)
        format_cell(lunary_day_col, bold=True, italic=True, font_size=10.5)
    elif lunary_day == 1:
        lunary_day_col.text = LUNAR_MONTHS_ENG[lunar_month - 1] + " " + str(lunary_day) + "*"
        format_cell(lunary_day_col, bold=True, italic=True, font_size=10.5)
    else:
        lunary_day_col.text = str(lunary_day)
        format_cell(lunary_day_col)

    if lunar_month not in uniq_lunar_months:
        uniq_lunar_months.append(lunar_month)

month_1, month_2 = uniq_lunar_months[0], uniq_lunar_months[1]
lunar_col_title = table.cell(2, 4)
lunar_col_title.text = LUNAR_MONTHS_ARABIC[month_1 - 1] + "\n" + LUNAR_MONTHS_ARABIC[month_1 - 2] + "\n" + "1442 AH"
lunar_col_title.paragraphs[0].runs[0].font.complex_script = True
lunar_col_title.paragraphs[0].runs[0].font.rtl = True
format_cell(lunar_col_title, font_size=10)

####################################################################################
# Insert Prayer Times
####################################################################################

prayer_times = get_prayer_time(month=MONTH)

num_of_rows = len(prayer_times)

prayers = {"Fajr": 5, "Duha": 8, "Dhuhr": 9, "Asr": 11, "Maghrib": 13, "Isha": 15}

for prayer, col_no in prayers.items():
    times = list(map(lambda x: x[prayer], prayer_times))

    for j in range(3, NUM_DAYS + 3):
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

doc.save("calendars/" + MONTH + "_" + YEAR + ".docx")

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
