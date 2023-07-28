import collections
import datetime
import pandas
import os

from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

load_dotenv()

DATA_FILE = os.getenv("DATA_FILE")
SHEET_NAME = os.getenv("SHEET_NAME")
FOUNDATION_YEAR = int(os.getenv("FOUNDATION_YEAR"))
FOUNDATION_MONTH = int(os.getenv("FOUNDATION_MONTH"))
FOUNDATION_DAY = int(os.getenv("FOUNDATION_DAY"))
SECONDS_IN_YEAR = int(os.getenv("SECONDS_IN_YEAR"))
CURRENT_YEAR = datetime.date.today()


def read_xlsx(data_file, sheet_name):
    excel_data_df = pandas.read_excel(data_file,
                                      sheet_name=sheet_name,
                                      na_values='nan', keep_default_na=False)
    all_drinks = excel_data_df.to_dict(orient='records')
    drinks = collections.defaultdict(list)
    for drink in all_drinks:
        drinks[drink['Категория']].append(drink)
    return drinks


def format_year(year):
    if year >= 5 and year <= 20:
        return f"{year} лет"
    if year > 100 and (
        str(year)[-1] + str(year)[-2] in ('11', '12', '13', '14')
                      ):
        return f"{year} лет"
    remainder_of_division = year % 10
    if remainder_of_division == 1:
        return f"{year} год"
    if remainder_of_division > 1 and remainder_of_division < 5:
        return f"{year} года"
    return f"{year} лет"


def count_age(foundation_year, current_year):
    foundation_year = datetime.date(
        year=FOUNDATION_YEAR, month=FOUNDATION_MONTH, day=FOUNDATION_DAY
    )
    age = int(
        (current_year - foundation_year).total_seconds() / SECONDS_IN_YEAR)
    return format_year(age)


if __name__ == '__main__':
    env = Environment(loader=FileSystemLoader("."),
                      autoescape=select_autoescape(["html"]))
    template = env.get_template("template.html")

    drinks_list = read_xlsx(DATA_FILE, SHEET_NAME)

    rendered_page = template.render(
        age=count_age(FOUNDATION_YEAR, CURRENT_YEAR),
        drinks=drinks_list
        )

    with open("index.html", "w", encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(("0.0.0.0", 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
