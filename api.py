from constants import GREG_MONTHS
import requests

def get_prayer_time(month, year):
    apiURL = "http://www.islamicfinder.us/index.php/api/prayer_times"

    date = year + "-" + str(GREG_MONTHS[month]) + "-01"

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

def get_islamic_year(month, year):
    apiURL = "http://www.islamicfinder.us/index.php/api/calendar"

    apiParams = {
        "day": 1,
        "month": GREG_MONTHS[month],
        "year": year,
    }

    response = requests.get(apiURL, params=apiParams)

    islamicYear = response.json()["to"][:4]

    return islamicYear

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
