import collections
import datetime
import pandas

from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

FOUNDATION_YEAR = 1920
FOUNDATION_MONTH = 1
FOUNDATION_DAY = 1
SECONDS_IN_YEAR = 31557600

data_file = 'wine3.xlsx'
sheet_name = 'Лист1'


def read_xlsx(data_file, sheet_name):
    excel_data_df = pandas.read_excel(data_file,
                                      sheet_name=sheet_name,
                                      na_values='nan', keep_default_na=False)
    drinks = excel_data_df.to_dict(orient='records')
    drinks_list = collections.defaultdict(list)
    for drink in drinks:
        drinks_list[drink['Категория']].append(drink)
    return drinks_list


def year_format(year):
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


def age(foundation_year):
    current_year = datetime.date.today()
    foundation_year = datetime.date(
        year=FOUNDATION_YEAR, month=FOUNDATION_MONTH, day=FOUNDATION_DAY
    )
    age = int(
        (current_year - foundation_year).total_seconds() / SECONDS_IN_YEAR)
    return year_format(age)


env = Environment(loader=FileSystemLoader("."),
                  autoescape=select_autoescape(["html"]))

template = env.get_template("template.html")

drinks_list = read_xlsx(data_file, sheet_name)

rendered_page = template.render(
    age=age(FOUNDATION_YEAR),
    drinks=drinks_list
    )

with open("index.html", "w", encoding="utf8") as file:
    file.write(rendered_page)


server = HTTPServer(("0.0.0.0", 8000), SimpleHTTPRequestHandler)
server.serve_forever()
