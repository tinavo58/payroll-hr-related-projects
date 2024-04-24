#!usr/bin/env python3
from mailmerge import MailMerge
import pandas as pd
import os, shutil
import win32com.client

file_location = 'Data Source'
new_location = 'Options MAR21 - Letters'
template_a_1 = os.path.join(file_location, '(a -1 ) Letter from xx - New Joiner.docx')
template_a_2 = os.path.join(file_location, '(a -2 ) Letter from xx - New Joiner 22-03.docx')
template_b = os.path.join(file_location, '(b) Letter from xx - Top-up.docx')
template_c = os.path.join(file_location, '(c) Letter from xx - Discretionary Retention.docx')
template_d = os.path.join(file_location, '(d) Letter from xx - Nil Top-up + Discretionary.docx')
template_e = os.path.join(file_location, '(e) Letter from xx - Founders.docx')
template_option_1 = os.path.join(file_location, '1a - Employee Share Option Plan Letter (incl Contractors) - Mar 2021 (Final).docx')
template_option_2 = os.path.join(file_location, '1b - Employee Share Option Plan Letter (incl Contractors) - Mar 2021 (Final).docx')
datasource = os.path.join(file_location, 'Data Source_Final.xlsx')
letter_3 = os.path.join(file_location, 'ESOP Plan Rules - Mar 2021.pdf')
letter_4 = os.path.join(file_location, 'Offer Information Statement for Employees - Mar 2021.pdf')

sheets = pd.ExcelFile(datasource)

df_letter = pd.read_excel(datasource, sheet_name=sheets.sheet_names[1],
                  usecols=['Letter from xx', 'firstName', 'lastName', 'last',
                           'preferredName', 'Total Options', 'address1', 'address2', 'Tranche1',
                           'Tranche2', 'Tranche3', 'emailAddress', '16 January 2018 $0.007',
                           '22 January 2018 $0.007', '30 January 2018 $0.007',
                           '31 March 2018 $0.039', '30 April 2018 $0.039',
                           '17 December 2018 $0.5359', '29 March 2019 $0.5359',
                           '29 March 2019 $0.7857', '29 July 2019 $0.7857',
                           '13 December 2019 $0.7857', '8 September 2020 $0.45']).fillna(0)

options_list = ['8 September 2020 $0.45',
                 '13 December 2019 $0.7857',
                 '29 July 2019 $0.7857',
                 '29 March 2019 $0.7857',
                 '29 March 2019 $0.5359',
                 '17 December 2018 $0.5359',
                 '30 April 2018 $0.039',
                 '31 March 2018 $0.039',
                 '30 January 2018 $0.007',
                 '22 January 2018 $0.007',
                 '16 January 2018 $0.007']

for idx in range(len(df_letter)):
    # assign variable
    letter_type = df_letter['Letter from xx'][idx]
    prefer = df_letter['preferredName'][idx]
    first = df_letter['firstName'][idx]
    last = df_letter['last'][idx]
    lastName = df_letter['lastName'][idx]
    email = df_letter['emailAddress'][idx]
    address1 = df_letter['address1'][idx]
    address2 = df_letter['address2'][idx]
    tr1 = df_letter['Tranche1'][idx]
    tr2 = df_letter['Tranche2'][idx]
    tr3 = df_letter['Tranche3'][idx]
    options = df_letter['Total Options'][idx]


    # get options history
    previous_options =[{
        'grantDate': '26 March 2021',
        'price': '$0.325',
        'previousOptions' : '{:,}'.format(options)
        }]

    for dateOptions in options_list:
        if df_letter[dateOptions][idx] != 0:
            option_history = {
                'grantDate': dateOptions.rsplit(' ', 1)[0],
                'price': dateOptions.rsplit(' ', 1)[-1],
                'previousOptions' : '{:,}'.format(int(df_letter[dateOptions][idx]))
                }
            previous_options.append(option_history)


    # letter 1
    # director
    if letter_type == 'director':
        doc_2 = MailMerge(template_option_1)
        doc_2.merge(
            first=first,
            email=email,
            address1=address1,
            last=lastName,
            tranche3='{:,}'.format(tr3),
            options='{:,}'.format(options),
            tranche2='{:,}'.format(tr2),
            address2=address2,
            tranche1='{:,}'.format(tr1)
            )
        ee_file_folder = os.path.join(new_location, 'Director', f"Options MAR21 - {prefer} {last}")
        if not os.path.exists(ee_file_folder):
            os.makedirs(ee_file_folder)
            doc_2.write(ee_file_folder + f"Employee Share Option Plan Letter Mar21 - {prefer[0]}.{last}.docx")
            shutil.copy(letter_3, ee_file_folder)
            shutil.copy(letter_4, ee_file_folder)


    # new joiner
    elif letter_type == 'new joiner':
        doc = MailMerge(template_a_1)
        doc.merge(
            prefer=prefer,
            last=last,
            totalOptions='{:,}'.format(options)
            )
        doc_2 = MailMerge(template_option_1)
        doc_2.merge(
            first=first,
            email=email,
            address1=address1,
            last=lastName,
            tranche3='{:,}'.format(tr3),
            options='{:,}'.format(options),
            tranche2='{:,}'.format(tr2),
            address2=address2,
            tranche1='{:,}'.format(tr1)
            )
        ee_file_folder = os.path.join(new_location, 'New Joiner', f"Options MAR21 - {prefer} {last}")
        if not os.path.exists(ee_file_folder):
            os.makedirs(ee_file_folder)
            doc.write(ee_file_folder + f"Letter from xx - {prefer} {last}.docx")
            doc_2.write(ee_file_folder + f"Employee Share Option Plan Letter Mar21 - {prefer[0]}.{last}.docx")
            shutil.copy(letter_3, ee_file_folder)
            shutil.copy(letter_4, ee_file_folder)


    # new joiner 22-03
    elif letter_type == 'new joiner-diff':
        doc = MailMerge(template_a_2)
        doc.merge(
            prefer=prefer,
            last=last,
            totalOptions='{:,}'.format(options)
            )
        doc_2 = MailMerge(template_option_2)
        doc_2.merge(
            first=first,
            email=email,
            address1=address1,
            last=lastName,
            tranche3='{:,}'.format(tr3),
            options='{:,}'.format(options),
            tranche2='{:,}'.format(tr2),
            address2=address2,
            tranche1='{:,}'.format(tr1)
            )
        ee_file_folder = os.path.join(new_location, 'New Joiner on 22-03', f"Options MAR21 - {prefer} {last}")
        if not os.path.exists(ee_file_folder):
            os.makedirs(ee_file_folder)
            doc.write(ee_file_folder + f"Letter from xx - {prefer} {last}.docx")
            doc_2.write(ee_file_folder + f"Employee Share Option Plan Letter Mar21 - {prefer[0]}.{last}.docx")
            shutil.copy(letter_3, ee_file_folder)
            shutil.copy(letter_4, ee_file_folder)

    # top-up
    elif letter_type == 'top-up':
        doc = MailMerge(template_b)
        doc.merge(
            prefer=prefer,
            last=last,
            totalOptions='{:,}'.format(options)
            )
        doc.merge_rows('grantDate', previous_options)
        doc_2 = MailMerge(template_option_1)
        doc_2.merge(
            first=first,
            email=email,
            address1=address1,
            last=lastName,
            tranche3='{:,}'.format(tr3),
            options='{:,}'.format(options),
            tranche2='{:,}'.format(tr2),
            address2=address2,
            tranche1='{:,}'.format(tr1)
            )
        ee_file_folder = os.path.join(new_location, 'Top-up', f"Options MAR21 - {prefer} {last}")
        if not os.path.exists(ee_file_folder):
            os.makedirs(ee_file_folder)
            doc.write(ee_file_folder + f"Letter from xx - {prefer} {last}.docx")
            doc_2.write(ee_file_folder + f"Employee Share Option Plan Letter Mar21 - {prefer[0]}.{last}.docx")
            shutil.copy(letter_3, ee_file_folder)
            shutil.copy(letter_4, ee_file_folder)

    # discretionary
    elif letter_type == 'discretionary':
        doc = MailMerge(template_c)
        doc.merge(
            prefer=prefer,
            last=last,
            totalOptions='{:,}'.format(options)
            )
        doc.merge_rows('grantDate', previous_options)
        doc_2 = MailMerge(template_option_1)
        doc_2.merge(
            first=first,
            email=email,
            address1=address1,
            last=lastName,
            tranche3='{:,}'.format(tr3),
            options='{:,}'.format(options),
            tranche2='{:,}'.format(tr2),
            address2=address2,
            tranche1='{:,}'.format(tr1)
            )
        ee_file_folder = os.path.join(new_location, 'Discretionary', f"Options MAR21 - {prefer} {last}")
        if not os.path.exists(ee_file_folder):
            os.makedirs(ee_file_folder)
            doc.write(ee_file_folder + f"Letter from xx - {prefer} {last}.docx")
            doc_2.write(ee_file_folder + f"Employee Share Option Plan Letter Mar21 - {prefer[0]}.{last}.docx")
            shutil.copy(letter_3, ee_file_folder)
            shutil.copy(letter_4, ee_file_folder)

    # nil-topup discretionary
    elif letter_type == 'nil top-up & discretionary':
        doc = MailMerge(template_d)
        doc.merge(
            prefer=prefer,
            last=last,
            totalOptions='{:,}'.format(options)
            )
        doc.merge_rows('grantDate', previous_options)
        doc_2 = MailMerge(template_option_1)
        doc_2.merge(
            first=first,
            email=email,
            address1=address1,
            last=lastName,
            tranche3='{:,}'.format(tr3),
            options='{:,}'.format(options),
            tranche2='{:,}'.format(tr2),
            address2=address2,
            tranche1='{:,}'.format(tr1)
            )
        ee_file_folder = os.path.join(new_location, 'Discretionary (no top-up)', f"Options MAR21 - {prefer} {last}")
        if not os.path.exists(ee_file_folder):
            os.makedirs(ee_file_folder)
            doc.write(ee_file_folder + f"Letter from xx - {prefer} {last}.docx")
            doc_2.write(ee_file_folder + f"Employee Share Option Plan Letter Mar21 - {prefer[0]}.{last}.docx")
            shutil.copy(letter_3, ee_file_folder)
            shutil.copy(letter_4, ee_file_folder)

    # founder
    elif letter_type == 'founder':
        doc = MailMerge(template_e)
        doc.merge(
            prefer=prefer,
            last=last,
            totalOptions='{:,}'.format(options)
            )
        doc_2 = MailMerge(template_option_1)
        doc_2.merge(
            first=first,
            email=email,
            address1=address1,
            last=lastName,
            tranche3='{:,}'.format(tr3),
            options='{:,}'.format(options),
            tranche2='{:,}'.format(tr2),
            address2=address2,
            tranche1='{:,}'.format(tr1)
            )
        ee_file_folder = os.path.join(new_location, 'Founder', f"Options MAR21 - {prefer} {last}")
        if not os.path.exists(ee_file_folder):
            os.makedirs(ee_file_folder)
            doc.write(ee_file_folder + f"Letter from xx - {prefer} {last}.docx")
            doc_2.write(ee_file_folder + f"Employee Share Option Plan Letter Mar21 - {prefer[0]}.{last}.docx")
            shutil.copy(letter_3, ee_file_folder)
            shutil.copy(letter_4, ee_file_folder)

word = win32com.client.Dispatch('Word.Application')

for dirpath, dirnames, filesnames in os.walk(new_location):
    for f in filesnames:
        if f.rsplit('.', 1)[-1] == 'docx':
            pdf = f.replace('.docx', '.pdf')
            in_file = (dirpath + '/' + f)
            new_file = (dirpath + '/' + pdf)
            doc = word.Documents.Open(in_file)
            doc.SaveAs(new_file, FileFormat=17)
            doc.Close()
        else:
            continue
word.Quit()
