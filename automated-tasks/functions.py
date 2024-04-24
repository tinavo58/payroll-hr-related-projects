#!usr/bin/env python3

import pandas as pd
import re
import numpy as np


file = 'DDEmployeeReport.csv'
bu_file = 'buCode.xlsx'

bu = pd.read_excel(bu_file, usecols=['BU', 'BusinessUnit'])
bu_dict = dict(zip(bu['BU'], bu['BusinessUnit']))


def create_names(df, col=['Given Name', 'Family Name']):
    """
    This func returns full names from df
    """
    names = []
    for first, last in df[col].values:
        names.append(f"{first.strip()} {last.strip().title()}")
    return names


def remove_commas_address1(col):
    """
    Remove the extra commas at the end
    address1 for letter mailmerge
    """
    new_list_without_commas, new_list = [], []
    for each in col:
        each = " ".join(each.split())
        if each.endswith(', '): # there was an extra space after commas
            new_list_without_commas.append(each[:-2])
        else:
            new_list_without_commas.append(each)
    for _ in new_list_without_commas:
        _ = _.replace(' ,', '').replace(',', '')
        if _.upper():
            new_list.append(_.title())
        else:
            new_list.append(_)
    return new_list


def combine_address(df, col=['Suburb', 'State ', 'Postcode']):
    """
    Return address2
    """
    address2 = []
    for suburb, state, postcode in df[col].values:
        address2.append(f"{suburb} {state} {str(postcode)}")
    return address2


def get_float(text):
    a = re.sub('[^(0-9.]+', '', text)
    a = re.sub('[^0-9.]+', '-', a)
    return float(a)


def create_amount_float(df, col_name='Amount'):
    """
    strip "$" and convert string to float
    """
    amount = []
    for _ in df[col_name]:
        if _ == " ":
            amount.append(np.nan)
        elif isinstance(_, str):
            amount.append(get_float(_))
    return amount


def preferred_name(selected, col_name='Given Name'):
    chosen = input('Preferred name (input or n): ').lower()
    if chosen == 'n':
        firstname = selected[col_name]
    else:
        firstname = chosen.title()
    return firstname


def print_options(list_input):
    """
    Returns list of options
    """
    if list_input:
        if len(list_input) == 1:
            print(f"\nEmployee found: {list_input[0]}")
        else:
            print('\nEmployee(s) found:')
            for index, option in enumerate(list_input, start=1):
                print(f"({index}) {option}")
    else:
        print('No employees found with your name input!')


def select_options(list_input, df, col_name='Fullname'):
    """
    Return record based
    """
    selected_dict = {}

    if len(list_input) > 1:
        option = int(input('\nPlease selection your option: '))-1
        name_selected = list_input[option]
        selected = df.loc[df[col_name] == name_selected]
        for each in selected.columns:
            selected_dict[each] = selected[each].values[0]
        print(f"\nYou have selected: {selected_dict['Fullname']}")

    elif len(list_input) == 1:
        selected = df.loc[df[col_name] == list_input[0]]
        for each in selected.columns:
            selected_dict[each] = selected[each].values[0]

    else:
        pass

    return selected_dict


def names_with_middle(df, col=['Given Name', 'Middle Name', 'Family Name']):
    names = []
    for f, m, l in df[col].values:
        if isinstance(m, str):
            names.append(f"{l}, {f} {m}")
        else:
            names.append(f"{l}, {f}")
    return names


def generate_df(file, bu_dict=bu_dict):
    """
    return df EmployeeReport
    return all cols
    """
    df = pd.read_csv(file, encoding='cp1252', parse_dates=['Start Date','Birthdate'], dayfirst=True)

    # convert datetime cols
    df['Start Date'] = pd.to_datetime(df['Start Date'], dayfirst=True)
    df['Birthdate'] = pd.to_datetime(df['Birthdate'], dayfirst=True)
    df['Termination Date'] = pd.to_datetime(df['Termination Date'], dayfirst=True)

    # create names & addresses
    df['Fullname'] = create_names(df)
    df['address1'] = remove_commas_address1(df['Street'])
    df['address2'] = combine_address(df)

    # employment details
    df['FTE'] = np.around(np.array(df['Hours']/164.67), 2)
    df['Hourly Rate'] = create_amount_float(df, col_name='Hourly Rate')
    df["Base"] = np.around(np.array(df['Hours']) * np.array(df['Hourly Rate']) * 12, 2)
    f = lambda x: round(x * 1.1) if x < 235_680 else round(x +  23_568) if x >= 235_680 else 0
    df['TEC'] = [f(a) for a in df['Base']]

    # reformat
    df['Base'] = [f"{a:,.0f}" for a in df['Base']]
    df['TEC'] = [f"{a:,.0f}" for a in df['TEC']]

    # business unit & cost centre
    df['CC'] = df['Cost Centre'].apply(lambda x: x.split(' ', 1)[-1])
    df['BU'] = df['Cost Centre'].apply(lambda x: int(x.split('-')[1])).map(bu_dict)

    return df

df = generate_df(file, bu_dict)
