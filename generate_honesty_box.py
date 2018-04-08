# -*- coding: utf-8 -*-
"""generate_honesty_box. Write to XLSX."""
import argparse
import numpy as np
import pandas as pd
import xlsxwriter
import yaml

parser = argparse.ArgumentParser()
parser.add_argument(
    "output_configuration",
    help="""Specify Columns to show.
        Must have aliases for 'name', 'budget' and 'pay_in'
        and minimum 1 alias beginning with 'buy'.
        Others will not be recognized."""
)
parser.add_argument(
    "csv_path",
    help="Specify Name and Debt CSV File location. Must be provided."
)
parser.add_argument(
    "--out", "--output_path",
    help="Specify where to put the xlsx file.",
    default="list.xlsx"
)
parser.add_argument(
    "--cap", "-c",
    help="Where to put the cap to stop people making debt.",
    default=-20,
    type=int
)
args = parser.parse_args()
print args

debt_conf_path = args.output_configuration
csv_path = args.csv_path
out = args.out

print debt_conf_path
print csv_path
print out
print args.cap


with open(debt_conf_path, 'r') as stream:
    c = yaml.load(stream)
    # TODO: Try Except logic
schulden = pd.read_csv(csv_path, encoding='utf-8')
schulden.sort_values('name', inplace=True)
schulden.reset_index(drop=True, inplace=True)
buy_keys = [c[k] for k in c if k.startswith('buy')]

budget_col = len(buy_keys) * 10 + 1
pay_in_col = len(buy_keys) * 10 + 2

# INITIATE XLSXWRITER OJECT
wb = xlsxwriter.Workbook(args.out)
ws = wb.add_worksheet()

# OPTIONS:
ws.fit_to_pages(1, 0)  # fit table to full width
ws.set_landscape()
ws.set_paper(9)  # A4
ws.center_horizontally()


# see http://xlsxwriter.readthedocs.io/format.html#set_border
# and http://xlsxwriter.readthedocs.io/format.html
normal_format = wb.add_format(
    {'align': 'center', 'valign': 'vcenter', 'border': 1})
eur_format = wb.add_format(
    {'num_format': u'#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]', 'border': 2})
header_format = wb.add_format(
    {'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 2})

# default row height
ws.set_default_row(10)
# change column width to 1 for buying fields
ws.set_column(1, len(buy_keys) * 10, 1, normal_format)
# first column width
column_width = np.percentile(
    [len(str(value)) for value in schulden['name']], 80)
ws.set_column(0, 0, column_width, header_format)
# header row height and format
ws.set_row(0, 20, header_format)
# TODO
ws.set_column(len(buy_keys) * 10 + 1, len(buy_keys) * 10 + 2, 10, eur_format)

# WRITE DATA
ws.write(0, 0, c['name'])
for i, key in enumerate(buy_keys):
    ws.merge_range(0, 10 * i + 1, 0, 10 * (i + 1), key)
ws.write(0, budget_col, c['budget'])
ws.write(0, pay_in_col, c['pay_in'])
for i, df_entry in schulden.iterrows():
    row_start = 1 + i * 2
    row_end = 2 + i * 2
    ws.merge_range(row_start, 0, row_end, 0, df_entry['name'])
    if df_entry['val'] < args.cap:
        ws.merge_range(
            row_start, 1, row_end, len(buy_keys) * 10,
            c['cap_format_str'].format(args.cap))
    ws.merge_range(row_start, budget_col, row_end, budget_col, df_entry['val'])
    ws.merge_range(row_start, pay_in_col, row_end, pay_in_col, '')
wb.close()
