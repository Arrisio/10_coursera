from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup
import bs4
import random
import string
import os
import argparse


def get_list_of_random_n_cources_urls(cources_urls_xml, n_cources=20):
    feed_soup = BeautifulSoup(cources_urls_xml, 'lxml')
    return random.sample(
        [xml_tag.get_text() for xml_tag in feed_soup.findAll('loc')],
        k=n_cources
    )


def extract_data_from_cource_web_page(web_page):

    metadata = {
        'Name': {'class_': 'title display-3-text'},
        'Language': {'class_': 'rc-Language'},
        'Nearest start date': {'class_': 'rc-StartDateString'},
        'Raiting': {'class_': 'ratings-text'},
        'Number of weeks': {'class_': 'week-heading',
                            'type_': 'count_elements'
                            }
    }

    feed_soup = BeautifulSoup(web_page, 'html.parser')

    exctracted_data = {}
    for param_name, param_preperties in metadata.items():
        if not param_preperties.get('type_'):
            param_value = feed_soup.find(class_=param_preperties.get('class_'))
            if isinstance(param_value, bs4.element.Tag):
                exctracted_data[param_name] = param_value.get_text()
            else:
                exctracted_data[param_name] = None

        elif param_preperties.get('type_') == 'count_elements':
            param_value = feed_soup.find_all(
                class_=param_preperties.get('class_')
            )
            if isinstance(param_value, list):
                exctracted_data[param_name] = len(param_value)
            else:
                exctracted_data[param_name] = None

    return exctracted_data


def get_web_page(url):
    response = requests.get(url)
    if response.ok:
        response.encoding = 'utf-8'
        return response.text


def set_auto_width_excel_cols(worksheet, indent=4):
    for i, column in enumerate(worksheet.iter_cols()):
        worksheet.column_dimensions[
            string.ascii_uppercase[i]].width = indent + max(
            [len(str(cell.value)) for cell in column]
        )


def make_excel_workbook_from_table(table):
    work_book = Workbook()
    work_sheet = work_book.active

    for row in table:
        work_sheet.append(row)

    set_auto_width_excel_cols(work_sheet)

    return work_book


def parse_arguments():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        '-f',
        action='store',
        dest='filepath',
        type=validate_path_to_save_file,
        help='filepath to save result file',
        default='coursera.xlsx'
    )

    parser.add_argument(
        '-n',
        action='store',
        dest='cources_number',
        type=int,
        help='numbet of cources to get information',
        default=20
    )

    return parser.parse_args()


def validate_path_to_save_file(filepath):
    result_dir = os.path.split(filepath)[0]
    if result_dir and not os.path.isdir(result_dir):
        raise argparse.ArgumentTypeError('Invalid output path')
    return filepath


if __name__ == '__main__':
    params = parse_arguments()

    random_cources_ulr_list = get_list_of_random_n_cources_urls(
        get_web_page('https://www.coursera.org/sitemap~www~courses.xml'),
        n_cources=params.cources_number
    )

    table_header = [
        'Name', 'Language', 'Nearest start date', 'Raiting', 'Number of weeks'
    ]

    cources_table = [table_header]
    for url in random_cources_ulr_list:
        cource_data = extract_data_from_cource_web_page(
            web_page=get_web_page(url)
        )
        cources_table.append(
            [cource_data.get(param_name) for param_name in table_header]
            )

    excel_workbook = make_excel_workbook_from_table(cources_table)
    excel_workbook.save(params.filepath)
