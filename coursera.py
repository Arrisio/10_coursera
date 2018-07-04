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


def parse_web_page(web_page):
    attr_mapping =  {
        'Name': {'class_': 'title display-3-text'},
        'Language': {'class_': 'rc-Language'},
        'Nearest start date': {'class_': 'rc-StartDateString'},
        'Raiting': {'class_': 'ratings-text'},
        'Number of weeks': {'class_': 'week-heading',
                            'type_': 'count_elements'}
    }

    feed_soup = BeautifulSoup(web_page, 'html.parser')

    parced_data = []
    for _, attr_params in attr_mapping.items():
        if not attr_params.get('type_'):
            param_value = feed_soup.find(class_=attr_params.get('class_'))
            if isinstance(param_value, bs4.element.Tag):
                parced_data.append(param_value.get_text())
            else:
                parced_data.append(None)

        elif attr_params.get('type_') == 'count_elements':
            param_value = feed_soup.find_all(class_=attr_params.get('class_'))
            if isinstance(param_value, list):
                parced_data.append(len(param_value))
            else:
                parced_data.append(None)

    return parced_data


def combine_array_with_cources_data(cources_urls):
    result_array = [[course_name for course_name in cources_attr_mapping]]
    for url in cources_urls:

        response = requests.get(url)
        if response.ok:
            response.encoding = 'utf-8'
            result_array.append(
                parse_web_page(response.text, cources_attr_mapping)
            )

    return result_array


def set_auto_width_excel_cols(worksheet, indent=4):
    for i, column in enumerate(worksheet.iter_cols()):
        worksheet.column_dimensions[
            string.ascii_uppercase[i]].width = indent + max(
            [len(str(cell.value)) for cell in column]
        )


def make_excel_workbook_from_array(array):
    work_book = Workbook()
    work_sheet = work_book.active

    for row in array:
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

    rnd_cources_ulr_list = random.sample(
        get_list_of_cources_urls(),
        k=params.cources_number
    )

    excel_workbook = make_excel_workbook_from_array(
        array=combine_array_with_cources_data(rnd_cources_ulr_list)
    )

    excel_workbook.save(params.filepath)
