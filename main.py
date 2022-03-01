import argparse
import collections
import datetime
import pandas
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


WINERY_OPENING_YEAR = 1920


if __name__ == '__main__':

    # parser = argparse.ArgumentParser()
    # parser.add_argument('way', help='Путь к файлу')
    # args = parser.parse_args()

    product_from_excel = pandas.read_excel('products.xlsx',
                                            sheet_name='Лист1',
                                            na_values=['N/A', 'NA'],
                                            keep_default_na=False)

    wines = product_from_excel.to_dict(orient='records')

    grouped_products = collections.defaultdict(list)

    for wine in wines:
        grouped_products[wine['Категория']].append(wine)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        winery_age=datetime.datetime.now().year - WINERY_OPENING_YEAR,
        products=grouped_products)

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('127.0.0.1', 9000), SimpleHTTPRequestHandler)
    server.serve_forever()
