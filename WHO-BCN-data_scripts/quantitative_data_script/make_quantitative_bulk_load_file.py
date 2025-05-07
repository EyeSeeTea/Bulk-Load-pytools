from copy import deepcopy
import os
import csv
import sys
import json
import argparse
import difflib
from argparse import ArgumentParser
import traceback
from collections import namedtuple
import openpyxl
import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl.cell import Cell, MergedCell
from data import NAMES_DICT, OLD_NAMES_DICT, COMBO_LIST, COUNTRY_DICT, INDICATOR_IGNORING_SERVICE, INDICATOR_IGNORING_QUINTILE


MetadataIds = namedtuple("MetadataIds", "indicators, countries, combos")

COC_DEFAULT_ID = ""
COC_TOTAL_ID = ""


def get_new_name(indicator_name: str, service: str):
    """Maps old data elements names to new ones

    Args:
        indicator_name (str): Old data element name
        service (str): data element categoryOptionCombos name

    Raises:
        ValueError: If a DE can't be mapped 

    Returns:
        new_indicator_name (str): New data element name
    """

    if indicator_name in OLD_NAMES_DICT:
        debug(f'Found old name: {indicator_name}')
        new_name = OLD_NAMES_DICT[indicator_name]
        if isinstance(new_name, dict):
            if service not in new_name:
                raise ValueError('get_new_name: cant map old indicator')

        return new_name[service] if isinstance(new_name, dict) else new_name
    else:
        return indicator_name


def make_combo_string(quintile: str, service: str):
    """Maps quintile and service to a categoryOptionCombos name

    Args:
        quintile (str): categoryOptions name
        service (str): categoryOptionCombos name

    Returns:
        combo (str): categoryOptionCombos name
    """

    if service == 'NA':
        result = quintile
        if quintile == "Total":
            result = "default"
    elif quintile == 'NA':
        result = service
    else:
        combo = quintile + ', ' + service
        combo_alt = service + ', ' + quintile

        if combo in COMBO_LIST:
            result = combo
        elif combo_alt in COMBO_LIST:
            result = combo_alt

    return result


def check_for_empty_csv_fields(**elements):
    """Checks if there are empty values in the CSV file and prints warning
    """

    for name, value in elements.items():
        if name != 'row' and not value:
            print(f'WARNING: Empty {name} variable in CSV file in row:\n{elements["row"]}')


def currency_converter(amount: str, country: str, year: str, figure: str):
    """Applies currency conversion to the CSV file values

    Args:
        amount (str): Original value
        country (str): Country code
        year (str): Year string
        figure (str): Figure code

    Returns:
        adjusted_amount (str): New value with currency conversion applied
    """

    currency_table = pd.read_csv(
        "https://docs.google.com/spreadsheets/d/1lEHQ9i-LO7gl0RWaJgYcfOJHjPefVXJbhgJ0Gn3iUPQ/export?format=csv&gid=56805701",
        decimal=","
    )

    currency_figures = ["F5", "F9", "F10a", "F10b", "F10c", "F10d", "F10e", "F10f", "F26"]

    debug(f'currency_converter amount: {amount} | country_code: {country} | year: {year} | figure: {figure}')

    if figure not in currency_figures:
        debug('currency_converter figure not in list')
        return amount

    # NOTE: Temporal fixes until the currency_table gets fixed
    if country == "GRC":
        country = "GRE"
    elif country == "DNK":
        country = "DEN"
    elif country == "IRL":
        country = "IRE"
    elif country == "ROU":
        country = "ROM"
    elif country == "NLD":
        country = "NET"

    try:
        coefficient = currency_table[
            (currency_table['code'] == country)
        ][year].values[0]
    except KeyError:
        # If no coefficient available for year, get the closest year to present
        last_year = next(reversed(currency_table.keys()))
        coefficient = currency_table[
            (currency_table['code'] == country)
        ][last_year].values[0]
    except Exception:
        traceback.print_exc()
        sys.exit(1)

    debug('currency_converter coefficient:', coefficient)

    adjusting_for_inflation = round(float(amount) / coefficient, 2)

    return str(adjusting_for_inflation)


def get_csv_indicator_value(value: str, real_value: str, real_flag: bool = False):
    """Checks if real_value exists and its not "NA" if --real_value flag is set

    Args:
        value (str): CSV Value field
        real_value (str): CSV real_value field

    Returns:
        real_value (str): Appropriate value based on --real_value flag
    """

    if real_flag:
        return real_value if real_value != "NA" else value

    return value


def create_dict_if_dont_exist(dictionary: dict, key: str):
    """check if key is in nested dictionary and creates a new empty dict if its not

    Args:
        dict (dict): Nested dictionary
        key (str): dictionary to check
    """

    if key not in dictionary:
        dictionary[key] = {}


def extract_values_from_csv(filename: str, real_flag: bool = False, currrency_flag: bool = False):
    """Given a CSV file name, creates a dictionary of the CSV file data

    Args:
        filename (str): CSV file name
        real_flag (bool): Flag for using the real values
        currency_flag (bool): Flag to apply currency converter

    Returns:
        values (dict): Dictionary with the CSV file data
    """

    values = {}

    try:
        with open(filename, 'r', encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                indicator_name = row['indicator_name']
                country = row['country']
                year = row['year']
                quintile = row['quintile']
                service = row['service']
                if currrency_flag and indicator_name != NAMES_DICT["POVERTY_LINE_OLD_NAME"]:
                    figure = row['figure_id']
                    value = currency_converter(row['value'], country, year, figure)
                else:
                    value = get_csv_indicator_value(row['value'], row['real_value'], real_flag)

                check_for_empty_csv_fields(
                    row=row,
                    indicator_name=indicator_name,
                    country=country,
                    year=year,
                    quintile=quintile,
                    service=service,
                    value=value
                )

                indicator_name = get_new_name(indicator_name, service)

                country_name = COUNTRY_DICT[country]

                create_dict_if_dont_exist(values, country_name)
                create_dict_if_dont_exist(values[country_name], year)
                create_dict_if_dont_exist(values[country_name][year], indicator_name)

                service = 'NA' if indicator_name in INDICATOR_IGNORING_SERVICE else service
                quintile = 'NA' if indicator_name in INDICATOR_IGNORING_QUINTILE else quintile

                cat_opt_combo = make_combo_string(quintile, service)
                if value != 'NA' and value is not None:
                    values[country_name][year][indicator_name][cat_opt_combo] = value
                else:
                    debug('Empty CSV value in row: ', row)

        return values
    except Exception:
        traceback.print_exc()
        sys.exit(1)


def get_metadata_ids(workbook: Workbook):
    """Crates a named tuple containing dictionaries with the ids of indicators, 
    countries and combos used, the ids are extracted from the bulk load template

    Args:
        workbook (Workbook): XLSX file with the bulk load template

    Returns:
        ids (MetadataIds): named tuple containing dictionaries with the ids of indicators, countries and combos used
    """

    global COC_DEFAULT_ID, COC_TOTAL_ID

    indicators_id_dict = {}
    countries_id_dict = {}
    combos_id_dict = {}
    sheet = workbook['Metadata']

    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
        identifier = row[0]
        type_col = row[1]
        name = str(row[2]).strip()

        if type_col == 'categoryOptionCombos':
            combos_id_dict[identifier] = name

            if name == "default":
                COC_DEFAULT_ID = identifier
            if name == "Total":
                COC_TOTAL_ID = identifier

        if type_col == 'dataElements':
            indicators_id_dict[name] = identifier

        if type_col == 'organisationUnit':
            countries_id_dict[name] = identifier

    return MetadataIds(indicators_id_dict, countries_id_dict, combos_id_dict)


def get_indicator_id(ids: MetadataIds, name: str):
    """Gets the ID of the provided DE and raises an error if the ID is not found

    Args:
        ids (MetadataIds): named tuple containing dictionaries with the ids of indicators, countries and combos used
        name (str): Data element name

    Returns:
        id (str): ID of the provided DE
    """

    try:
        return ids.indicators[name]
    except KeyError:
        print(f'ERROR: Data element "{name}" can\'t be matched with an ID, check metadata')
        print(f'Closest candidates: {difflib.get_close_matches(name, ids.indicators.keys())}')
        return None


def update_latest_dict(latest_pre_2019_des_dict: dict, indicator_name: str, year: str, combo_id: str, value: str):
    """Check if the DE need to populate a latest version of it and returns the updated latest_pre_2019_des_dict

    Args:
        latest_pre_2019_des_dict (dict): Dictionary storing the latest year, combo and value
        indicator_name (str): Data element name
        year (str): Data element value year
        combo_id (str): Data element value combo
        value (str): Data element value
    """

    latest_year = None
    if indicator_name in latest_pre_2019_des_dict and int(year) < 2020:
        latest_year = latest_pre_2019_des_dict[indicator_name].get(combo_id, ['0'])[0]

        if int(latest_year) < int(year):
            latest_pre_2019_des_dict[indicator_name][combo_id] = [year, value]

        debug(
            f'update_latest_dict check: indicator_name: "{indicator_name}" | year: {year}" | latest_year: {latest_year}'
        )


def store_latest_data(data: dict, ids: MetadataIds, latest_pre_2019_des_dict: dict, country_id: str):
    """Stores the data values for the latest DEs for the given country

    Args:
        data (dict): Dictionary with the matched data
        ids (MetadataIds): named tuple containing dictionaries with the ids of indicators, countries and combos used
        latest_pre_2019_des_dict (dict): Dictionary storing the latest year, combo and value
        country_id (str): Data country ID
    """

    for indicator_name, latest_dict in latest_pre_2019_des_dict.items():
        for latest_combo, combo_data in latest_dict.items():
            latest_year = combo_data[0]
            latest_value = combo_data[1]
            if latest_year != "0":
                latest_indicator_name = indicator_name + " - 2019 or LAY"
                last_indicator_id = get_indicator_id(ids, latest_indicator_name)
                if not last_indicator_id:
                    continue

                create_dict_if_dont_exist(data[country_id][latest_year], last_indicator_id)

                data[country_id][latest_year][last_indicator_id][latest_combo] = latest_value
                debug(
                    f'store_latest_data check: "{latest_indicator_name}" | {latest_year}" | {data[country_id][latest_year][last_indicator_id][latest_combo]}'
                )


def make_matched_values(csv_values_dict: dict, ids: MetadataIds):
    """Maps the formNames and data of csv_values_dict with the metadata ids, applies direct monthly transformation

    Args:
        csv_values_dict (dict): Dictionary with the CSV data
        ids (MetadataIds): named tuple containing dictionaries with the ids of indicators, countries and combos used

    Returns:
        (dict): Dictionary with the CSV values indexed by metadata ids
    """

    data = {}

    for country, country_data in csv_values_dict.items():
        latest_pre_2019_des_dict = {
            NAMES_DICT["CATA_HEALTHCARE_TOTAL_NAME"]: {},
            NAMES_DICT["OOP_CHE_NAME"]: {},
            NAMES_DICT["GGHED_GGE_NAME"]: {},
            NAMES_DICT["CATA_QUINTILE_NAME"]: {},
            NAMES_DICT["CATA_TOTAL_NAME"]: {},
            NAMES_DICT["FURTHERIMPOV_CATA_NAME"]: {},
            NAMES_DICT["IMPOV_CATA_NAME"]: {},
            NAMES_DICT["UN_EUSILC_DENTAL_QUINTILE_NAME"]: {}
        }

        country_id = ids.countries[country]
        data[country_id] = {}

        for year, indicators in country_data.items():
            create_dict_if_dont_exist(data[country_id], year)

            for indicator_name, indicator_combos in indicators.items():
                indicator_id = get_indicator_id(ids, indicator_name)
                if not indicator_id:
                    continue

                create_dict_if_dont_exist(data[country_id][year], indicator_id)

                store_transformation_de(indicator_name, indicator_id)

                for combo_name, value in indicator_combos.items():
                    combo_ids = []
                    for id_code, name in ids.combos.items():
                        if name == combo_name:
                            combo_ids.append(id_code)
                    combo_id = '|'.join(combo_ids)

                    if check_mean_monthly_indicator(indicator_name):
                        debug("check_mean_monthly_indicator: ", indicator_name, value, float(value)/12)
                        value = str(float(value)/12)

                    update_latest_dict(latest_pre_2019_des_dict, indicator_name, year, combo_id, value)

                    data[country_id][year][indicator_id][combo_id] = value

        store_latest_data(data, ids, latest_pre_2019_des_dict, country_id)

    debug("pre_2019_de_names: ", dump_json_var(latest_pre_2019_des_dict))
    return data


# TRANSFORMATIONS
SHARE_HH_WITH_OOP_TOTAL = None
SHARE_HH_NO_OOP_TOTAL = None
SHARE_HH_WITH_OOP_QUINTILE = None
SHARE_HH_NO_OOP_QUINTILE = None
GGHED_CHE = None
VHI_CHE = None
OOP_CHE = None
OTHER_CHE = None


def check_mean_monthly_indicator(indicator_name: str):
    """Check if data element needs the monthly transformation

    Args:
        indicator_name (str): Form Name of the data element

    Returns:
        (bool): Boolean value of the check
    """

    mean_monthly_names = [NAMES_DICT["SEL_MONTHLY_NAME"], NAMES_DICT["CTP_MONTHLY_NAME"]]

    return indicator_name in mean_monthly_names


def store_transformation_de(indicator_name: str, indicator_id: str):
    """Stores the data elements ids needed for transformations 

    Args:
        indicator_name (str): data element form name
        indicator_id (str): data element id
    """

    global SHARE_HH_WITH_OOP_TOTAL, SHARE_HH_NO_OOP_TOTAL, SHARE_HH_WITH_OOP_QUINTILE, SHARE_HH_NO_OOP_QUINTILE
    global GGHED_CHE, VHI_CHE, OOP_CHE, OTHER_CHE

    if indicator_name == NAMES_DICT["SHARE_HH_WITH_OOP_TOTAL_NAME"]:
        SHARE_HH_WITH_OOP_TOTAL = indicator_id
    elif indicator_name == NAMES_DICT["SHARE_HH_NO_OOP_TOTAL_NAME"]:
        SHARE_HH_NO_OOP_TOTAL = indicator_id
    elif indicator_name == NAMES_DICT["SHARE_HH_WITH_OOP_QUINTILE_NAME"]:
        SHARE_HH_WITH_OOP_QUINTILE = indicator_id
    elif indicator_name == NAMES_DICT["SHARE_HH_NO_OOP_QUINTILE_NAME"]:
        SHARE_HH_NO_OOP_QUINTILE = indicator_id
    elif indicator_name == NAMES_DICT["GGHED_CHE_NAME"]:
        GGHED_CHE = indicator_id
    elif indicator_name == NAMES_DICT["VHI_CHE_NAME"]:
        VHI_CHE = indicator_id
    elif indicator_name == NAMES_DICT["OOP_CHE_NAME"]:
        OOP_CHE = indicator_id
    elif indicator_name == NAMES_DICT["OTHER_CHE_NAME"]:
        OTHER_CHE = indicator_id


def get_indicator_value(metadata_dict: dict, country_id: str, year: str, indicator_id: str, combo_id: str, default: str | None = None):
    """Get the value of an indicator from the metadata dictionary.
    
    Args:
        metadata_dict (dict): Dictionary with the values indexed by metadata id
        country_id (str): Value country ID
        year (str): Value year
        indicator_id (str): Value data element ID
        combo_id (str): Value combo ID
        default (str | None, optional): Value to return if not found, print find error if not provided. Defaults to None.

    Returns:
        value (str): Found value or default
    """

    try:
        return metadata_dict[country_id][year][indicator_id][combo_id]
    except KeyError:
        if default:
            return default

        print(f'ERROR: Can\'t find value for country: {country_id} year: {year} de: {indicator_id} combo: {combo_id}')
        return None


def get_spending_share_indicator(metadata_dict: dict, ids: dict, de: str, name: str):
    """Tries to get the data element value and prints a warning if no value can be found

    Args:
        metadata_dict (dict): Dictionary with the values indexed by metadata id
        ids (dict): Dictionary with the requested data element country, year and combo
        de (str): Data element id
        name (str): Data element name

    Returns:
        (str | None): Value of the requested data element or None in case of error
    """

    try:
        return float(metadata_dict[ids["country_id"]][ids["year"]][de][ids["combo_id"]])
    except KeyError:
        print(f'WARNING: Data element "{name}" for OU {ids["country_id"]} - {ids["year"]} is missing')
        return None


def make_transformations(metadata_dict: dict):
    """Performs transformations for the 'Share of households with out-of-pocket payments for health care' and 
    'Other spending as a share of current spending on health'

    Args:
        metadata_dict (dict): Dictionary with the values indexed by metadata id
    """

    new_metadata_dict = deepcopy(metadata_dict)

    households_ids = {
        SHARE_HH_WITH_OOP_TOTAL: SHARE_HH_NO_OOP_TOTAL,
        SHARE_HH_WITH_OOP_QUINTILE: SHARE_HH_NO_OOP_QUINTILE
    }

    for country_id, country_data in new_metadata_dict.items():
        for year, indicators in country_data.items():
            for indicator_id, indicator_combos in indicators.items():
                if indicator_id in households_ids.keys():
                    debug(f"make_transformations: {country_id} - {year} - {indicator_id}")

                    without_id = households_ids[indicator_id]
                    if not bool(indicator_combos):
                        indicator_combos = new_metadata_dict[country_id][year][without_id]

                    for combo_id, _ in indicator_combos.items():
                        without_value = get_indicator_value(new_metadata_dict, country_id, year, without_id, combo_id)

                        if without_value:
                            create_dict_if_dont_exist(new_metadata_dict[country_id][year][indicator_id], combo_id)
                            new_metadata_dict[country_id][year][indicator_id][combo_id] = str(
                                100 - float(without_value))
                            debug(
                                f"make_transformations calc: {country_id} - {year} - {indicator_id} - {new_metadata_dict[country_id][year][indicator_id][combo_id]}"
                            )

            if any(ind in indicators.keys() for ind in [GGHED_CHE, VHI_CHE, OOP_CHE, OTHER_CHE]):
                ids = {"country_id": country_id, "year": year, "combo_id": COC_DEFAULT_ID}

                gghed_che_value = get_spending_share_indicator(
                    new_metadata_dict, ids, GGHED_CHE, NAMES_DICT["GGHED_CHE_NAME"])
                vhi_che_value = get_spending_share_indicator(
                    new_metadata_dict, ids, VHI_CHE, NAMES_DICT["VHI_CHE_NAME"])
                oop_che_value = get_spending_share_indicator(
                    new_metadata_dict, ids, OOP_CHE, NAMES_DICT["OOP_CHE_NAME"])

                create_dict_if_dont_exist(new_metadata_dict[country_id][year], OTHER_CHE)
                create_dict_if_dont_exist(new_metadata_dict[country_id][year][OTHER_CHE], COC_DEFAULT_ID)

                if gghed_che_value and vhi_che_value and oop_che_value:
                    debug(f"make_transformations: {country_id} - {year} - {OTHER_CHE}")
                    new_metadata_dict[country_id][year][OTHER_CHE][COC_DEFAULT_ID] = str(
                        100-(gghed_che_value + vhi_che_value + oop_che_value)
                    )
                    debug(
                        f"make_transformations calc: {country_id} - {year} - {OTHER_CHE} - {new_metadata_dict[country_id][year][OTHER_CHE][COC_DEFAULT_ID]}"
                    )
                else:
                    print(
                        f'WARNING: Data element "{NAMES_DICT["OTHER_CHE_NAME"]}" for OU {country_id} - {year} is missing values for transformation'
                    )
                    new_metadata_dict[country_id][year][OTHER_CHE][COC_DEFAULT_ID] = ""

        return new_metadata_dict


def write_org_unit(last_cell: Cell, metadata_dict: dict):
    """Writes the countries in the CSV data to the bulk load file

    Args:
        last_cell (Cell): Previous cell of the column
        metadata_dict (dict): Dictionary with the values indexed by metadata id
    """

    for country_id, country_data in metadata_dict.items():
        for _ in country_data:
            new_cell = last_cell.offset(row=1, column=0)
            new_cell.value = f'=_{country_id}'

            last_cell = new_cell


def write_years(last_cell: Cell, metadata_dict: dict):
    """Writes the years in the CSV data to the bulk load file

    Args:
        last_cell (Cell): Previous cell of the column
        metadata_dict (dict): Dictionary with the values indexed by metadata id
    """

    for _, country_data in metadata_dict.items():
        for year in country_data:
            new_cell = last_cell.offset(row=1, column=0)
            new_cell.value = year

            last_cell = new_cell


def write_indicator(col_indicator: str, col_combo: str, last_cell: Cell, metadata_dict: dict):
    """Writes the data elements in the CSV data to the bulk load file

    Args:
        col_indicator (str): Id of the data element
        col_combo (str): Id of the data elements combo
        last_cell (Cell): Previous cell of the column
        metadata_dict (dict): Dictionary with the values indexed by metadata id

    Returns:
        (int): Number of data elements added to the bulk load file
    """

    count = 0
    offset = 0
    country_offset = 0

    for _, country_data in metadata_dict.items():
        years = list(country_data.keys())
        for year, indicators in country_data.items():
            offset = years.index(year) + country_offset
            for indicator_id, indicator_combos in indicators.items():
                if indicator_id == col_indicator:
                    for combo_id, value in indicator_combos.items():
                        ids = combo_id.split('|') if '|' in combo_id else combo_id
                        if col_combo in ids or (col_combo == COC_DEFAULT_ID and combo_id == COC_TOTAL_ID):
                            new_cell = last_cell.offset(row=1 + offset, column=0)
                            new_cell.value = value

                            count += 1

        # Offset for the next country
        country_offset += years.index(year) + 1

    return count


def write_values(workbook: Workbook, metadata_dict: dict, out_filename: str):
    """Writes the CSV data to a new bulk load file using workbook as a template

    Args:
        workbook (Workbook): XLSX file with the bulk load template
        metadata_dict (dict): Dictionary with the values indexed by metadata id

    Returns:
        (int): Number of data elements added to the bulk load file
    """

    sheet = workbook['Data Entry']
    workbook.active = workbook['Data Entry']
    count = 0

    for index, col in enumerate(sheet.iter_cols(min_row=4)):
        if index == 0:
            last_cell = col[-1]
            write_org_unit(last_cell, metadata_dict)
        if index == 1:
            last_cell = col[-1]
            write_years(last_cell, metadata_dict)
        if index == 2:
            pass
        if index > 2:
            if not isinstance(col[0], MergedCell):
                col_indicator = str(col[0].value).rsplit('=_', maxsplit=1)[-1]
            col_combo = str(col[1].value).rsplit('=_', maxsplit=1)[-1]
            last_cell = col[-1]

            count += write_indicator(col_indicator, col_combo,
                                     last_cell, metadata_dict)

    workbook.save(out_filename)

    debug(f'excel count: {count}')
    return count


def debug(*msg):
    """Writes the debug message to LOG_FILE
    """

    if DEBUG:
        with open(LOG_FILE, "a", encoding="utf-8") as log_file:
            print(*msg, file=log_file)


def dump_json_var(var: any):
    """Transforms var object to a JSON string

    Args:
        var (any): Object to be transformed

    Returns:
        (str): JSON string of the object
    """

    return json.dumps(var, indent=2)


def get_metadata_dict_len(metadata_dict: dict):
    """Gets the number of matched values from the CSV

    Args:
        metadata_dict (dict): Dictionary with the values indexed by metadata id

    Returns:
        (int): Number of matched values
    """

    lenght = 0

    for years in metadata_dict.values():
        for indicators in years.values():
            for combos in indicators.values():
                lenght += len(combos)
    return lenght


def filepath_exists(filepath: str):
    """Checks if path exists and its a file

    Args:
        filepath (str): Path to the file

    Returns:
        (bool): Value of the check
    """

    return os.path.isfile(filepath)


def get_template_path(parser: ArgumentParser, xlsx_template: str):
    """Returns a path to the bulk load template, if the --xlsx_template is not used returns the defaut template path.
    If the path is not valid prints error and exits.  

    Args:
        parser (ArgumentParser): Script argument parser. 
        xlsx_template (str): Value of --xlsx_template argument

    Returns:
        (str): Path to the template
    """
    if not filepath_exists(xlsx_template):
        parser.error(f'The template: {xlsx_template} doesn\'t exist')

    return xlsx_template


OUT_FILENAME = ''

DEBUG = False
LOG_FILE = 'log.json'


def main():
    parser = argparse.ArgumentParser(description='Process CSV from "Data Extraction Tool" into "Bulk Load" XLSX files. \
                                     The script needs a template, it either can be supplied with the --xlsx_template \
                                     argument or by placing a template named "Quantitative_Data_UHCPW_Template.xlsx" \
                                     in the same folder as the script.\
                                     Outputs to a XLSX file with same name as the source CSV one.')
    parser.add_argument('indicators_csv', type=str, help='Source CSV file')
    parser.add_argument('-x', '--xlsx_template', type=str, default='Quantitative_Data_UHCPW_Template.xlsx',
                        help='Bulk Load Quantitative XLSX template file path, if empty tries to open "Quantitative_Data_UHCPW_Template.xlsx"')
    adjusted_values_args = parser.add_mutually_exclusive_group()
    adjusted_values_args.add_argument('-r', '--real_value', action='store_true',
                                      help='Use real_value (if not NA) instead of value from the CSV source file, cant be used with -c/--currency')
    adjusted_values_args.add_argument('-c', '--currency', action='store_true',
                                      help='Apply currency adjustment to the applicable values, cant be used with -r/--real_value')
    parser.add_argument('-d', '--debug', action='store_true',
                        help='Display debug logs, its recommended to redirect the output into a file, e.g: ... > log.txt')
    args = parser.parse_args()

    if not filepath_exists(args.indicators_csv):
        parser.error(f'The source file: {args.indicators_csv} doesn\'t exist')

    global DEBUG
    out_filename = f'{args.indicators_csv.split(".csv")[0]}.xlsx'
    DEBUG = args.debug

    if DEBUG:
        f = open(LOG_FILE, 'w', encoding="utf-8")
        f.close()

    args.xlsx_template = get_template_path(parser, args.xlsx_template)

    debug('Source file:', args.indicators_csv)
    debug('Template:', args.xlsx_template)
    debug('Output file:', out_filename)

    try:
        wb = openpyxl.load_workbook(args.xlsx_template)
    except Exception as e:
        debug("openpyxl exception: ", e)
        traceback.print_exc()
        sys.exit(1)

    csv_values_dict = extract_values_from_csv(args.indicators_csv, args.real_value, args.currency)
    debug('csv_values_dict:\n ', dump_json_var(csv_values_dict))

    ids = get_metadata_ids(wb)

    debug(f'indicators ids:\n len: {len(ids.indicators)}\n values:\n', dump_json_var(ids.indicators))
    debug(f'countries ids:\n len: {len(ids.countries)}\n values:\n', dump_json_var(ids.countries))
    debug(f'combos ids:\n len: {len(ids.combos)}\n values:\n', dump_json_var(ids.combos))

    metadata_dict = make_matched_values(csv_values_dict, ids)
    metadata_dict = make_transformations(metadata_dict)

    csv_count = get_metadata_dict_len(metadata_dict)

    debug(f'metadata_dict count: {csv_count}\n')
    debug('metadata_dict:\n', dump_json_var(metadata_dict))

    excel_count = write_values(wb, metadata_dict, out_filename)
    debug(f'write_values count: {excel_count}\n')

    print(f'Processed {csv_count} entries from CSV file, written {excel_count} values to EXCEL')


if __name__ == '__main__':
    main()
