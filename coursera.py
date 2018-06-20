from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup
import bs4
from random import shuffle
import string
import os
import argparse


def get_list_of_n_cources_urls(
    courses_list_url='https://www.coursera.org/sitemap~www~courses.xml',
    headers={
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; '
                      'rv:45.0) Gecko/20100101 Firefox/45.0'
    },
    number_cources=20
):
    response = requests.get(courses_list_url, headers=headers)
    feed_soup = BeautifulSoup(response.text, 'lxml')
    courses_list = [xml_tag.get_text() for xml_tag in feed_soup.findAll('loc')]
    shuffle(courses_list)
    return courses_list[:number_cources]


def parse_web_page(web_page, attr_mapping):
    feed_soup = BeautifulSoup(web_page, 'html.parser')

    parced_data = []
    for attr_name, attr_params in attr_mapping.items():
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


def combine_array_with_cources_data(
        cources_urls,
        headers={
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; '
                          'rv:45.0) Gecko/20100101 Firefox/45.0'
        },
        cources_attr_mapping={
            'Name': {'class_': 'title display-3-text'},
            'Language': {'class_': 'rc-Language'},
            'Nearest start date': {'class_': 'rc-StartDateString'},
            'Raiting': {'class_': 'ratings-text'},
            'Number of weeks': {'class_': 'week-heading',
                                'type_': 'count_elements'}
        }
):
    result_array = [[course_name for course_name in cources_attr_mapping]]
    for url in cources_urls:

        response = requests.get(url, headers=headers, allow_redirects=True)
        if response.ok:
            response.encoding = 'utf-8'
            result_array.append(
                parse_web_page(response.text, cources_attr_mapping)
            )

    return result_array


def auto_set_width_excel_cols(worksheet, indent=4):
    for i, column in enumerate(worksheet.iter_cols()):
        worksheet.column_dimensions[
            string.ascii_uppercase[i]].width = indent + max(
            [len(str(cell.value)) for cell in column]
        )


def array_to_excel_workbook(array):
    work_book = Workbook()
    work_sheet = work_book.active

    for row in array:
        work_sheet.append(row)

    auto_set_width_excel_cols(work_sheet)

    return work_book


def parse_arguments():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        '-f',
        action='store',
        dest='filepath',
        type=path_to_save_file,
        help='filepath to save result file',
        default='coursera.xmlx'
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


def path_to_save_file(filepath):
    result_dir = os.path.split(filepath)[0]
    if result_dir and not os.path.isdir(result_dir):
        raise argparse.ArgumentTypeError('Invalid output path')
    return filepath


if __name__ == '__main__':
    params = parse_arguments()

    excel_workbook = array_to_excel_workbook(
        combine_array_with_cources_data(
            get_list_of_n_cources_urls(number_cources=params.cources_number)
        )
    )

    excel_workbook.save(params.filepath)
