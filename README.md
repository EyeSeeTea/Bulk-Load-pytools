# Bulk Load Pytools

Python tools to generate Bulk Load Templates

## Installation

To install the dependencies required to run this script, use the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

If you prefer to use a virtual environment, you can create one and activate it before installing the dependencies:

```bash
python3 -m venv env
source env/bin/activate
pip install -r requirements.txt
```

## make_quantitative_bulk_load_file.py

This script processes CSV files from the "Data Extraction Tool" into "Bulk Load" XLSX files.
The script needs a Bulk Load template. It can either be supplied with the `--xlsx_template` argument or by placing a template named "Quantitative_Data_UHCPW_Template.xlsx" in the same folder as the script.
The output file will be named as the input CSV file, but with XLSX extension.
There's two options to get the adjusted value, either from the CSV file real_value column or calculating it from the value column. The data for the `--currency` is retrieved from: [google sheet](https://docs.google.com/spreadsheets/d/1lEHQ9i-LO7gl0RWaJgYcfOJHjPefVXJbhgJ0Gn3iUPQ#gid=56805701).

### Usage

```
python3 make_quantitative_bulk_load_file.py [-h | --help] [-x | --xlsx_template <XLSX_TEMPLATE>] [-r | --real_value | -c | --currency] [-d | --debug] indicators_csv
```

#### Positional Arguments

`indicators_csv`: The source CSV file.

#### Options

`-h`, `--help`: Show the help message and exit.

`-x XLSX_TEMPLATE`, `--xlsx_template XLSX_TEMPLATE`: The Bulk Load Quantitative XLSX template file path. If empty, the script will try to open "Quantitative_Data_UHCPW_Template.xlsx".

`-r`, `--real_value`: Use `real_value` instead of `value` from the CSV source file, cant be used with `-c`/`--currency`.

`-c`, `--currency`: Apply currency adjustment to the applicable values, cant be used with `-r`/`--real_value`.

`-d`, `--debug`: Print debug logs into a `log.json` file.

### Examples

Simple use:

```bash
python3 make_quantitative_bulk_load_file.py SPA_fp_indicators.csv
```

Specifying a template and using the 'real_value' column (if not 'NA') from the source file:

```bash
python3 make_quantitative_bulk_load_file.py SPA_fp_indicators.csv -r --xlsx_template=~/docs/Quantitative_Template.xlsx
```

Specifying a template and using the currency adjustment when applicable:

```bash
python3 make_quantitative_bulk_load_file.py SPA_fp_indicators.csv -c --xlsx_template=~/docs/Quantitative_Template.xlsx
```

Using debug flag:

```bash
python3 make_quantitative_bulk_load_file.py SPA_fp_indicators.csv -d
```

## make_qualitative_bulk_load_file.py

This script processes DOCX files into "Bulk Load" XLSX files.
The script needs a Bulk Load template. It can either be supplied with the `--xlsx_template` argument or by placing a template named "Quantitative*Data_UHCPW_Template.xlsx" in the same folder as the script.
The output file will be a XLSX file named *\<COUNTRY>\_\<YEAR>\_Qualitative*Data.xlsx*.

### Usage

```
python3 make_qualitative_bulk_load_file.py [-h] [-x XLSX_TEMPLATE] [-d] docx_filename
```

#### Positional Arguments

`docx_filename`: The path to the DOCX source file.

#### Options

`-h`, `--help`: Show the help message and exit.

`-x XLSX_TEMPLATE`, `--xlsx_template XLSX_TEMPLATE`: The Bulk Load Quantitative XLSX template file path. If empty, the script will try to open "Quantitative_Data_UHCPW_Template.xlsx".

`-d`, `--debug`: Print debug logs into a `log.json` file.

### Examples

Simple use:

```bash
python3 make_qualitative_bulk_load_file.py summary_tables.docx
```

Specifying a template and using debug flag:

```bash
python3 make_qualitative_bulk_load_file.py summary_tables.docx --xlsx_template=~/docs/Qualitative_Template.xlsx -d
```
