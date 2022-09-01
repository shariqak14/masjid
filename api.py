from datetime import datetime
from constants import GREG_MONTHS
import requests

def get_prayer_time(month):
    apiURL = "http://www.islamicfinder.us/index.php/api/prayer_times"

    date = str(datetime.now().year) + "-" + str(GREG_MONTHS[month]) + "-01"

    apiParams = {
        "country": "US",
        "zipcode": "06029",
        "juristic": 1,
        "time_format": 2,
        "show_entire_month": 1,
        "date": date,
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
