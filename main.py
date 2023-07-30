import collections
import datetime
import pandas
import os

from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

SECONDS_IN_YEAR = 31557600


def read_xlsx(data_file, sheet_name):
    excel_data_df = pandas.read_excel(
        data_file, sheet_name=sheet_name, na_values="nan",
        keep_default_na=False
    )
    all_drinks = excel_data_df.to_dict(orient="records")
    drinks = collections.defaultdict(list)
    for drink in all_drinks:
        drinks[drink["Категория"]].append(drink)
    return drinks


def format_year(year):
    if year >= 5 and year <= 20:
        return f"{year} лет"
    if year > 100 and (
                str(year)[-1] + str(year)[-2] in ("11", "12", "13", "14")):
        return f"{year} лет"
    remainder_of_division = year % 10
    if remainder_of_division == 1:
        return f"{year} год"
    if remainder_of_division > 1 and remainder_of_division < 5:
        return f"{year} года"
    return f"{year} лет"


def count_age(foundation_date, current_year):
    age = (current_year - foundation_date).total_seconds() / SECONDS_IN_YEAR
    return format_year(int(age))


if __name__ == "__main__":

    load_dotenv()

    data_file = os.getenv("DATA_FILE", default='wine3.xlsx')
    sheet_name = os.getenv("SHEET_NAME", default='Лист1')
    date = os.getenv("FOUNDATION_DATE", default='1920-01-01')
    foundation_date = datetime.datetime.fromisoformat(date)
    today = datetime.date.today()
    today_datetime = datetime.datetime.combine(
        today, datetime.datetime.min.time())

    env = Environment(
        loader=FileSystemLoader("."), autoescape=select_autoescape(["html"])
    )
    template = env.get_template("template.html")

    drinks_list = read_xlsx(data_file, sheet_name)

    rendered_page = template.render(
        age=count_age(foundation_date, today_datetime), drinks=drinks_list
    )

    with open("index.html", "w", encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(("0.0.0.0", 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
