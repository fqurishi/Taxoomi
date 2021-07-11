##############################################################################
#
# sales tax generator for simple used car dealership
#
# Copyright 2021 Faisl Qurishi, faisl@faislqurishi.dev
#
import xlsxwriter

# introduce program
print("Welcome to Taxoomi!\n")

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('new_sales_tax.xlsx')
worksheet = workbook.add_worksheet()

# Widen the columns to make the text clearer.
worksheet.set_column('A:L', 15)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})
money = workbook.add_format({'num_format': '$#,##0.00'})

# Write header text.
# phone number
worksheet.write('A1', '(510) 582-8000')
# month and year
# grab month and year
print("What month and year is it?\n")
month  = input("Month?\n")
year = input("Year?\n")
worksheet.write('D1', month.upper() + ' ' + year.upper())
# company name  
worksheet.write('E1', 'BAY AREA MOTOR')
# company address
worksheet.write('E2', '21572 MISSION BLVD HAYWARD CA 94541')

# Create columns
worksheet.write('A3', 'VIN')
worksheet.write('B3', 'COUNTY')
worksheet.write('C3', 'CITY')
worksheet.write('D3', 'PURCHASE')
worksheet.write('E3', 'SOLD')
worksheet.write('F3', 'WARRANTY')
worksheet.write('G3', 'GAP')
worksheet.write('H3', 'TAX')
worksheet.write('I3', 'LIC/REG')
worksheet.write('J3', 'SMOG')
worksheet.write('K3', 'DOC')


# Get user input for size of sales
sales_size = input("How many cars have we sold this month? (not wholesale)\n")
while True:
    if (sales_size.isnumeric()):
        break
    else:
        print("Please enter a positive number from 0.\n")
        sales_size = input("How many cars have we sold this month? (not wholesale)\n")


wholesale_size = input("How many cars have we sold wholesale this month?\n")
while True:
    if (wholesale_size.isnumeric()):
        break
    else:
        print("Please enter a positive number from 0.\n")
        wholesale_size = input("How many cars have we sold this month? (not wholesale)\n")

# Fill in values.
for x in range(3, int(sales_size)+3):
    print("Car " + str(x-2))
    vin = input("Enter VIN.\n")
    county = input("Enter County.\n")
    city = input("Enter City.\n")
    while True:
        try:
            purchase = float(input("Enter purchase price.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            sold = float(input("Enter sold price.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            warranty = float(input("Enter warranty price.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            gap = float(input("Enter GAP price.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            tax = float(input("Enter sales tax amount.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            lic_reg = float(input("Enter lic&reg fees.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            smog = float(input("Enter smog fees.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            doc = float(input("Enter doc fees.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    # start writing in the row
    worksheet.write(x, 0, vin)
    worksheet.write(x, 1, county)
    worksheet.write(x, 2, city)
    worksheet.write(x, 3, purchase, money)
    worksheet.write(x, 4, sold, money)
    worksheet.write(x, 5, warranty, money)
    worksheet.write(x, 6, gap, money)
    worksheet.write(x, 7, tax, money)
    worksheet.write(x, 8, lic_reg, money)
    worksheet.write(x, 9, smog, money)
    worksheet.write(x, 10, doc, money)

# Make wholesale header.
worksheet.write(int(sales_size)+4, 0, 'WHOLESALE')

# Fill in wholesale values.
for x in range(int(sales_size)+5, int(wholesale_size)+int(sales_size)+5):
    print("Wholesale Car " + str(x-(int(sales_size)+4)))
    vin = input("Enter VIN.\n")
    while True:
        try:
            purchase = float(input("Enter purchase price.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    while True:
        try:
            sold = float(input("Enter sold price.\n"))
        except ValueError:
            print("Please enter a positive number from 0.\n")
        else:
            break
    # start writing in the row
    worksheet.write(x, 0, vin)
    worksheet.write(x, 3, purchase, money)
    worksheet.write(x, 4, sold, money)

# Make Totals header
worksheet.write(int(wholesale_size)+int(sales_size)+6, 2, 'TOTALS:', bold)
# Fill in totals values
worksheet.write(int(wholesale_size)+int(sales_size)+7, 3, '=SUM(D1:D' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)
worksheet.write(int(wholesale_size)+int(sales_size)+7, 4, '=SUM(E1:E' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)
worksheet.write(int(wholesale_size)+int(sales_size)+7, 5, '=SUM(F1:F' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)
worksheet.write(int(wholesale_size)+int(sales_size)+7, 6, '=SUM(G1:G' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)
worksheet.write(int(wholesale_size)+int(sales_size)+7, 7, '=SUM(H1:H' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)
worksheet.write(int(wholesale_size)+int(sales_size)+7, 8, '=SUM(I1:I' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)
worksheet.write(int(wholesale_size)+int(sales_size)+7, 9, '=SUM(J1:J' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)
worksheet.write(int(wholesale_size)+int(sales_size)+7, 10, '=SUM(K1:K' + str(int(wholesale_size)+int(sales_size)+5) + ')', money)



workbook.close()