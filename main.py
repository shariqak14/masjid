from docx import Document, enum
from docx.shared import Pt

import requests

MONTH = "September"
YEAR = "2021"


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


months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

prayer_times = get_prayer_time(month=MONTH)

num_of_rows = len(prayer_times)

##################################################
# Edit Document
##################################################

def format_cell(cell, center=True, bold=False, font_size=12, font="Arial"):
    if center: cell.paragraphs[0].alignment = enum.text.WD_ALIGN_PARAGRAPH.CENTER

    cell.paragraphs[0].runs[0].font.bold = bold

    cell.paragraphs[0].runs[0].font.size = Pt(font_size)

    cell.paragraphs[0].runs[0].font.name = font


doc = Document("June_2021.docx")
table = doc.tables[0]

month_heading = table.cell(0, 10)
month_heading.text = MONTH.upper() + " " + YEAR
format_cell(month_heading, bold=True, font_size=15, font="Times New Roman")

month_row_title = table.cell(1, 2)
month_row_title.text = MONTH + " " + YEAR + " CE"
format_cell(month_row_title, bold=True, font_size=11, font="Times New Roman")

##################################################
# Insert Prayer Times
##################################################

prayers = {'Fajr': 5, 'Duha': 8, 'Dhuhr': 9, 'Asr': 12, 'Maghrib': 14, 'Isha': 16}

for prayer, col_no in prayers.items():
    # if prayer not in ['Duha', 'Dhuhr']:
    for j in range(3, 33):
        times = list(map(lambda x: x[prayer], prayer_times))
        pr_time = table.cell(j, col_no)
        pr_time.text = prayer[:2] # times[j-3]
        format_cell(pr_time)

# table.add_row()
doc.save("grid2.docx")