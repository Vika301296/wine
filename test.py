import collections
import pandas
from pprint import pprint

# excel_data_df = pandas.read_excel('wine.xlsx',
#                                   sheet_name='Лист1')
# print(excel_data_df.to_dict(orient='records'))

data_file = 'wine2.xlsx'
sheet_name = 'Лист1'


def read_xlsx(data_file, sheet_name):
    excel_data_df = pandas.read_excel(data_file,
                                      sheet_name=sheet_name,
                                      na_values='nan', keep_default_na=False)
    return excel_data_df.to_dict(orient='records')


wines = read_xlsx(data_file, sheet_name)
# pprint((wines))

# new_dict = {
#     'Белые вина': [],
#     'Красные вина': [],
#     'Напитки': []
# }

new_dict = collections.defaultdict(list)
for wine in wines:
    new_dict[wine['Категория']].append(wine)


# for wine in wines:
#     wine_category = wine['Категория']
#     if wine_category not in new_dict:
#         new_dict[wine_category] = []
#         new_dict[wine_category].append(wine)
#     new_dict[wine_category].append(wine)

# red_wines_list = []
# white_wines_list = []
# drinks_list = []

# for wine in wines:
#     if wine['Категория'] == 'Белые вина':
#         white_wines_list.append(wine)
#     elif wine['Категория'] == 'Красные вина':
#         red_wines_list.append(wine)
#     else:
#         drinks_list.append(wine)

# new_dict['Белые вина'] = white_wines_list
# new_dict['Красные вина'] = red_wines_list
# new_dict['Напитки'] = drinks_list

pprint(new_dict)
