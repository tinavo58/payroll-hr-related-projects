#!usr/bin/env python3
from mailmerge import MailMerge
import datetime as dt

# get the contract template
word_doc = 'Template - FullTime Contract.docx'

# get merged fields
contract = MailMerge(word_doc)

def cleaned_text(x):
    x = x.strip().title()
    return x

# name
fullName = cleaned_text(input('Full Name: '))
# FullName = fullName
firstName = cleaned_text(input('First Name: '))

# address
address1 = cleaned_text(input('Address (number & street): '))
suburb = cleaned_text(input('Suburb: '))
state = cleaned_text(input('State: '))
postcode = cleaned_text(input('Postcode: '))
address2 = suburb + ' ' + state.upper() + ' ' + postcode
fullAddress = address1 + ' ' + address2

# position
position = cleaned_text(input('Position: '))
managerPosition = cleaned_text(input('Manager Position Title: '))

#date
commencementDate = dt.datetime.strptime(input('Start Date: '),"%d/%m/%Y").strftime("%d %B %Y")
Date = dt.datetime.strptime(input('Date contract to be sent: '),"%d/%m/%Y").strftime("%d %B %Y")

# salary
baseSalary = input('Base Salary: ')
TEC = float(baseSalary) * 1.095 if float(baseSalary) < 228_360 else float(baseSalary) + 21694.2

baseSalary = "${:,.2f}".format(float(baseSalary))
TEC = "${:,.2f}".format(TEC)

# create contract
contract.merge(
    Address1=address1,
    Address2=address2,
    fullAddress=fullAddress,
    commencementDate=commencementDate,
    Date=Date,
    baseSalary=baseSalary,
    TEC=TEC,
    fullName=fullName,
#     FullName=FullName,
    firstName=firstName,
    position=position,
    managerPosition=managerPosition
)
contract.write(dt.date.today().strftime('%y%m%d') + ' Full Time Contract - ' + fullName +'.docx')
