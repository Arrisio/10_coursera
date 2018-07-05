from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup
import bs4
import random
import string
import os
import argparse


def get_list_of_cources_urls(
        url='https://www.coursera.org/sitemap~www~courses.xml'
):
    response = requests.get(url)
    feed_soup = BeautifulSoup(response.text, 'lxml')
    return [xml_tag.get_text() for xml_tag in feed_soup.findAll('loc')]


def parse_web_page(web_page, metadata):
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


def get_cource_web_page(url):
    response = requests.get(url)
    if response.ok:
        response.encoding = 'utf-8'
        return response.text


def make_cources_data_dict(cources_urls):
    cources_metadata = {
        'Name': {'class_': 'title display-3-text', 'order': 0},
        'Language': {'class_': 'rc-Language', 'order': 1},
        'Nearest start date': {'class_': 'rc-StartDateString', 'order': 2},
        'Raiting': {'class_': 'ratings-text', 'order': 3},
        'Number of weeks': {'class_': 'week-heading',
                            'type_': 'count_elements', 'order': 4
                            }
    }

    cources_data = {
        'metadata': {'attributes_names': sorted(
                cources_metadata.keys(),
                key=lambda key: cources_metadata.get(key).get('order', 100))
        },
        'data': []
    }
    for url in cources_urls:
        cources_data['data'].append(
            parse_web_page(
                web_page=get_cource_web_page(url),
                metadata=cources_metadata
            )
        )

    return cources_data


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


def make_table_from_data_dict(dict_):
    """ Input dictionary must contain key "data" where contains
        a list of dictionaries with data. Also it could contain key "metadata"
        and funtion tries to exctract column names from it"""

    columns_names = (
            dict_.get('metadata', {}).get('attributes_names') or
            list(dict_.get('data', [{}])[0].keys())
    )
    table = [dict_.get('metadata', {}).get('attributes_names')]

    for record in dict_.get('data'):
        row = []
        for attribute_name in columns_names:
            row.append(record.get(attribute_name))
        table.append(row)

    return table


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

    random_cources_ulr_list = random.sample(
        get_list_of_cources_urls(),
        k=params.cources_number
    )

    excel_workbook = make_excel_workbook_from_table(
        table=make_table_from_data_dict(
            make_cources_data_dict(random_cources_ulr_list)
        )
    )

    excel_workbook.save(params.filepath)
