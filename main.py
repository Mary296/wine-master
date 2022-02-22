import collections
import datetime
import pandas
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


if __name__ == '__main__':
    product_from_excel = pandas.read_excel('wine3.xlsx',
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
        year=datetime.datetime.now().year - 1920,
        products=grouped_products)

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('127.0.0.1', 9000), SimpleHTTPRequestHandler)
    server.serve_forever()
