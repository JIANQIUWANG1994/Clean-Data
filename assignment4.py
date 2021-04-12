import openpyxl
import pandas as pd
excel_file = openpyxl.load_workbook('Assignment4.xlsx')
excel_sheet = excel_file['Example 1']

excel_sheet.delete_rows(idx=6)
excel_sheet.delete_rows(idx=9)
excel_file.save('edited.xlsx')

excel_workbook = 'edited.xlsx'
sheet = pd.read_excel(excel_workbook, sheet_name='Example 1')
new_sheet = sheet.fillna("")

product_list = []
excel_products = new_sheet['Product']

for product in excel_products:
    product_names = product.strip().title()
    product_list.append(product_names)

print(product_list)

sheet.insert(0,"product",product_list)
del sheet['Product']
print(sheet)

delivery_person_list = []
excel_delivery_person = new_sheet['Delivery Person']
for delivery_person in excel_delivery_person:
    person_name = delivery_person.strip().title()
    delivery_person_list.append(person_name)
print(delivery_person_list)
sheet.insert(7,"Delivery person",delivery_person_list)
del sheet['Delivery Person']
print(sheet)

sheet.to_excel("edited.xlsx")

