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
        Others will not be recognized.""",
)
parser.add_argument(
    "csv_path", help="Specify Name and Debt CSV File location. Must be provided."
)
parser.add_argument(
    "--csv_new", help="Specify new path for calculated csv file.", default=""
)
parser.add_argument(
    "--out",
    "--output_path",
    help="Specify where to put the xlsx file.",
    default="list.xlsx",
)
parser.add_argument(
    "--cap",
    "-c",
    help="Where to put the cap to stop people making debt.",
    default=-20,
    type=int,
)
args = parser.parse_args()

if args.csv_new == "":
    csv_new = (lambda s: ".".join(s[:-1]) + "_copy." + s[-1])(args.csv_path.split("."))
else:
    csv_new = args.csv_new

# ### read csv ###

schulden = pd.read_csv(args.csv_path, encoding="utf-8", sep=";", decimal=",")
if not (
    all([x in schulden.columns for x in ["name", "val", "einzahlung"]])
    or (("name" in schulden.columns) & (len(schulden.columns) > 1))
):
    raise AssertionError(
        "Either provide 'name' and 'col' in input or 'name' and value cols"
    )
schulden.sort_values("name", inplace=True)
schulden.reset_index(drop=True, inplace=True)
if "val" not in schulden.columns:
    schulden["val"] = 0
schulden.fillna(0, inplace=True)

schulden["val"] = schulden["val"].astype(float)
for col in np.setdiff1d(schulden.columns, ["name", "val", "einzahlung"]):
    schulden["val"] -= schulden[col] * float(col.replace(",", "."))
    schulden[col].astype(str)
    schulden[col] = ""
schulden["val"] += schulden["einzahlung"]
schulden["einzahlung"].astype(str)
schulden["einzahlung"] = ""
schulden["val"] = schulden["val"].round(2)
schulden.to_csv(csv_new, sep=",", decimal=".", encoding="utf-8", index=False)


# ### write to xlsx  ###
with open(args.output_configuration, "r") as stream:
    c = yaml.load(stream)
    # TODO: Try Except logic

buy_keys = [c[k] for k in sorted([k for k in c if k.startswith("buy")])]
# unfortunately this hack is needed for sorted cols, so you can specify order
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
ws.repeat_rows(0)  # repeat first row
# default row height
ws.set_default_row(10)
# change column width to 1 for buying fields
ws.set_column(1, len(buy_keys) * 10, 1)

# see http://xlsxwriter.readthedocs.io/format.html#set_border
# and http://xlsxwriter.readthedocs.io/format.html
default_format = wb.add_format({"border": 1})
cap_string_format = wb.add_format({"border": 1, "valign": "vcenter", "align": "center"})
eur_format = wb.add_format(
    {
        "num_format": u"[>{}]#,##0.00 [$€-407];[<{}][RED]-#,##0.00 [$€-407]".format(
            args.cap / 2, args.cap / 2
        ),
        "bottom": 1,
        "left": 2,
        "right": 2,
        "valign": "vcenter",
        "align": "right",
    }
)
# formatter for header
header_format = wb.add_format(
    {"bold": 1, "align": "center", "valign": "vcenter", "bottom": 2, "right": 2}
)
# formatter for name
right_border = wb.add_format({"bottom": 1, "right": 2, "valign": "vcenter"})

slim_style = 4  # which style to use for slim lines


inner_top_format = wb.add_format({"bottom": slim_style, "right": slim_style})
inner_bottom_format = wb.add_format({"bottom": 1, "right": slim_style})
inner_bold_top_format = wb.add_format({"bottom": slim_style, "right": 2})
inner_bold_bottom_format = wb.add_format({"bottom": 1, "right": 2})

# first column width
column_width = np.percentile([len(valu) for valu in schulden["name"]], 80)
ws.set_column(0, 0, column_width, right_border)
# header row height and format
ws.set_row(0, height=20)


for i, df_entry in schulden.iterrows():
    row_start = 1 + i * 2
    row_end = 2 + i * 2
    ws.merge_range(row_start, 0, row_end, 0, df_entry["name"])
    if df_entry["val"] < args.cap:
        # debt is under money cap, user should pay
        ws.merge_range(
            row_start,
            1,
            row_end,
            len(buy_keys) * 10,
            c["cap_format_str"].format(args.cap),
            cap_string_format,
        )
    else:
        for i in range(1, len(buy_keys) + 1):
            ws.write(row_start, i * 10, "", inner_bold_top_format)
            ws.write(row_end, i * 10, "", inner_bold_bottom_format)
        for i in range(1, (len(buy_keys)) * 10):
            if i % 10:
                ws.write(row_start, i, "", inner_top_format)
                ws.write(row_end, i, "", inner_bottom_format)
    ws.merge_range(row_start, budget_col, row_end, budget_col, df_entry["val"])
    ws.merge_range(row_start, pay_in_col, row_end, pay_in_col, "")
    ws.set_row(row_end)

# WRITE DATA TO COLUMN NAME FIELDS
ws.write(0, 0, c["name"], header_format)
for i, key in enumerate(buy_keys):
    ws.merge_range(0, 10 * i + 1, 0, 10 * (i + 1), key, header_format)
ws.write(0, budget_col, c["budget"], header_format)
ws.write(0, pay_in_col, c["pay_in"], header_format)
# SET FORMAT OF LAST TWO COLS
ws.set_column(budget_col, pay_in_col, 10, eur_format)

wb.close()
