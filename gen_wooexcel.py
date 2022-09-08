#!/usr/bin/env python
import csv
import os

DEFAULT_INPUT_FILENAME = "data/data.csv"
DEFAULT_OUTPUT_FILENAME_PREFIX = "wooexcel_"

input_filename = input("请将要处理的文件拖动到窗口范围内。")
if len(input_filename) < 1:
    input_filename = DEFAULT_INPUT_FILENAME

working_directory = os.path.dirname(input_filename)

output_filename = os.path.join(
    working_directory,
    DEFAULT_OUTPUT_FILENAME_PREFIX + os.path.split(input_filename)[1]
)

fieldnames = [
    'Type',
    'SKU',
    'Name',
    'Description',
    'Tax status',
    'Tax class',
    'In stock?',
    'Stock',
    'Allow customer reviews?',
    'Saleprice',
    'Regularprice',
    'Categories',
    'Tags',
    'Images',
    'Parent',
    'Position',
    'Attribute 1 name',
    'Attribute 1 value(s)',
    'Attribute 1 visible',
    'Attribute 1 global',
    'Attribute 2 name',
    'Attribute 2 value(s)',
    'Attribute 2 visible',
    'Attribute 2 global',
    'Attribute 3 name',
    'Attribute 3 value(s)',
    'Attribute 3 visible',
    'Attribute 3 global',
]


def fill_row(is_parent, data, idx, a1v=None):
    return {
        'Type': 'variable' if is_parent else 'variation',
        'SKU': data['SKU'] if is_parent else data['SKU'] + '-' + str(idx),
        'Name': data['Name'],
        'Description': data['Description'],
        'Tax status': 'texable',
        'Tax class': '' if is_parent else 'parent',
        'In stock?': '1' if is_parent else '',
        'Stock': '' if is_parent else data['Stock'],
        'Allow customer reviews?': '1' if is_parent else '0',
        'Saleprice': '',
        'Regularprice': '' if is_parent else '60',
        'Categories': data['Categories'] if is_parent else '',
        'Tags': '',
        'Images': data['Images'] if is_parent else '',
        'Parent': '' if is_parent else data['SKU'],
        'Position': str(idx),
        'Attribute 1 name': data['Attribute 1 name'],
        'Attribute 1 value(s)': data['Attribute 1 value(s)'] if is_parent else a1v,
        'Attribute 1 visible': '1' if is_parent else '',
        'Attribute 1 global': '1',
        'Attribute 2 name': '',
        'Attribute 2 value(s)': '',
        'Attribute 2 visible': '',
        'Attribute 2 global': '',
        'Attribute 3 name': '',
        'Attribute 3 value(s)': '',
        'Attribute 3 visible': '',
        'Attribute 3 global': '',
    }


with open(input_filename, 'r', newline='') as input_file:
    reader = csv.DictReader(input_file)
    output_file = open(output_filename, 'w', newline='')
    writer = csv.DictWriter(output_file, fieldnames=fieldnames)
    writer.writeheader()
    for row in reader:
        print("Processing: {sku}".format(sku=row['SKU']))
        is_parent = True
        parent_sku = row['SKU']
        idx = 0
        r = fill_row(is_parent, row, idx)  # fill parent line
        writer.writerow(r)  # write parent line

        # fill child line
        is_parent = False
        for a1v in row['Attribute 1 value(s)'].split(','):
            idx += 1
            r = fill_row(is_parent, row, idx, a1v=a1v)
            writer.writerow(r)  # write parent line

    output_file.close()
