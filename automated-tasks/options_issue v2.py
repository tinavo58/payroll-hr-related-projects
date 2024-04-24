#!usr/bin/env python3
from mailmerge import MailMerge
import pandas as pd
import os, shutil
from docx2pdf import convert
from PyPDF2 import PdfFileMerger
import functions


def divide_tranche(col):
    t1, t2, t3 = [], [], []
    for _ in col:
        a = _ // 3
        if _ % 3 == 1:
            t1.append(a)
            t2.append(a)
            t3.append(a+1)
        elif _ % 3 == 2:
            t1.append(a+1)
            t2.append(a+1)
            t3.append(a)
        else:
            t1.append(a)
            t2.append(a)
            t3.append(a)
    return t1, t2, t3


file = '210712 Draft ESOP.xlsx'

sheets = pd.ExcelFile(file).sheet_names
usecols = ['eeCode', 'Employee Name', '26 June 2021 \n[Discretionary]\n$0.3252',
       '26 June 2021 \n[New Hires]\n$0.3252', '2 August 2021 [New Hires]\n$0.325',
       '2 August 2021\n[Discretionary]\n$0.325', '26 June 2021 & 2 August 2021 \n[Total]\n$0.325']

df_options = pd.read_excel(file, sheet_name=sheets[0], header=1, usecols=usecols, skipfooter=1)
df_options['New Hire'] = df_options['26 June 2021 \n[New Hires]\n$0.3252'] + df_options['2 August 2021 [New Hires]\n$0.325']
df_options['Discretionary'] = df_options['26 June 2021 \n[Discretionary]\n$0.3252'] + df_options['2 August 2021\n[Discretionary]\n$0.325']
df_options.drop(['26 June 2021 \n[New Hires]\n$0.3252', '2 August 2021 [New Hires]\n$0.325', '26 June 2021 \n[Discretionary]\n$0.3252', '2 August 2021\n[Discretionary]\n$0.325'], 1, inplace=True)
df_options.rename(columns={'26 June 2021 & 2 August 2021 \n[Total]\n$0.325': 'Total'}, inplace=True)

cond_a = df_options['Total'] != 0
cond_b = ~df_options['Employee Name'].str.contains('Hurditch|Westwood', case=False)
df_selected = df_options.loc[cond_a & cond_b]

# read employee report into df
df_ee = pd.read_csv('DDEmployeeReport.csv', encoding='cp1252')
df_ee['Fullname'] = functions.create_names(df_ee[['Given Name', 'Family Name']])
df_ee['address1'] = functions.remove_commas_address1(df_ee['Street'])
df_ee['address2'] = functions.combine_address(df_ee[['Suburb', 'State ', 'Postcode']])
df_ee_names = df_ee.loc[:, ('Employee Code', 'Fullname')]
df_selected = df_selected.merge(df_ee_names, how='left', left_on='Employee Name', right_on='Fullname')
df_selected = df_selected.loc[:, ('Employee Code', 'Employee Name', 'New Hire', 'Discretionary', 'Total')]
df_selected['Type'] = ['new_hire' if a > 0 else 'discretionary' for a in df_selected['New Hire']]

# re-order columns
df_ee_names_address = df_ee.loc[:, ('Employee Code', 'Fullname', 'address1', 'address2', 'Email')]
df_selected = df_selected.merge(df_ee_names_address, how='left', on='Employee Code')
df_selected = df_selected.loc[:, ('Employee Code', 'Fullname', 'address1', 'address2', 'Email', 'Type', 'Total' )]

df_selected['tranche1'] = divide_tranche(df_selected['Total'])[0]
df_selected['tranche2'] = divide_tranche(df_selected['Total'])[1]
df_selected['tranche3'] = divide_tranche(df_selected['Total'])[-1]

# export to excel files
df_selected.to_excel('options_jul2021.xlsx', index=False)
df_ee_names.to_excel('names_codes.xlsx', index=False)


### create letters
template_disc = 'Templates_discretionary.docx'
l2 = 'Employee Share Option Plan Letter.docx'
l3 = 'ESOP Plan Rules.pdf'
l4 = 'Offer Information Statement.pdf'

location = 'Employee packs'

main_file = 'Options.xlsx'
df_final = pd.read_excel(main_file, sheet_name='datasource', skiprows=1)

options_list = ['29 Mar 2019 $0.5359', '29 Mar 2019 $0.7857',
       '29 Jul 2019 $0.7857', '13 Dec 2019 $0.7857', '8 Sep 2020 $0.4500',
       '26 Mar 2021 $0.325', '2 Aug 2021 $0.325']

options_list = list(reversed(options_list))

for _ in range(len(df_final)):
#     assign variables
    l_type = df_final['Type'][_]
    fullname = df_final['Fullname'][_]
    first = df_final['firstname'][_]
    prefer = df_final['prefer'][_]
    address1 = df_final['address1'][_]
    address2 = df_final['address2'][_]
    email = df_final['Email'][_]
    tranche1 = '{:,}'.format(df_final['tranche1'][_])
    tranche2 = '{:,}'.format(df_final['tranche2'][_])
    tranche3 = '{:,}'.format(df_final['tranche3'][_])
    options = '{:,}'.format(df_final['Total'][_])

    previous_options = [{
        'grantDate': '17 Dec 2021',
        'price': '$0.325',
        'options': options
    }]
    for dateOptions in options_list:
        if df_final[dateOptions][_] != 0:
            option_history = {
                'grantDate': dateOptions.rsplit(' ', 1)[0],
                'price': dateOptions.rsplit(' ', 1)[-1],
                'options' : '{:,}'.format(int(df_final[dateOptions][_]))
            }
            previous_options.append(option_history)

    doc_l2 = MailMerge(l2)
    doc_l2.merge(
        totalOptions=options,
        tranche2=tranche2,
        firstname=first,
        fullname=fullname,
        address2=address2,
        tranche1=tranche1,
        tranche3=tranche3,
        address1=address1,
        emailAddress=email
        )

    if l_type == 'discretionary':
        ee_folder = os.path.join(location, f"Options DEC21 - {fullname}")
        doc_l1 = MailMerge(template_disc)
        doc_l1.merge(
            totalOptions=options,
            preferName=prefer,
            fullName=fullname,
            )
        doc_l1.merge_rows('grantDate', previous_options)

        if not os.path.exists(ee_folder):
            os.makedirs(ee_folder)
            doc_l1.write(ee_folder + f"Letter from xx - {fullname}.docx")
            doc_l2.write(ee_folder + f"Employee Share Option Plan Letter - {fullname}.docx")
            convert(ee_folder)
            pdfs = [ee_folder + f"Letter from xx - {fullname}.pdf",
                    ee_folder + f"Employee Share Option Plan Letter - {fullname}.pdf"]
            merger = PdfFileMerger()
            for pdf in pdfs:
                merger.append(pdf)
            merger.write(ee_folder + f"ESOP Dec21 - {fullname}.pdf")
            merger.close
            shutil.copy(l3, ee_folder)
            shutil.copy(l4, ee_folder)
