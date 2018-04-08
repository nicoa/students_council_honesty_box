# students_council_honesty_box
Automatically show the actual state of an honesty box system in Excel Spreadsheet.

From a list of names and values generate an excel sheet with options to buy items. Gives ability to block persons if they fall under a certain threshold / cap and to force them to pay in some credit.


## Example Data
Example Data lies in `example_data`:
- debt_conf (german example)
- example_people.csv (from [1] and [2])


## Required Packages
- xlsxwriter
- pandas
- argparse
- yaml

## Usage
Example call:
`python generate_honesty_box.py --out example_data/list.xlsx example_data/debt_conf.yaml example_data/example_people.csv`

Get help with `python generate_honesty_box.py -h`:

```
usage: generate_honesty_box.py [-h] [--out OUT] [--cap CAP]
                               output_configuration csv_path

positional arguments:
  output_configuration  Specify Columns to show. Must have aliases for 'name',
                        'budget' and 'pay_in' and minimum 1 alias beginning
                        with 'buy'. Others will not be recognized.
  csv_path              Specify Name and Debt CSV File location. Must be
                        provided.

optional arguments:
  -h, --help            show this help message and exit
  --out OUT, --output_path OUT
                        Specify where to put the xlsx file.
  --cap CAP, -c CAP     Where to put the cap to stop people making debt.
```



[1]: http://listofrandomnames.com/index.cfm?textarea
[2]: https://www.w3.org/2001/sw/rdb2rdf/wiki/Lists_of_generic_names_for_use_in_examples