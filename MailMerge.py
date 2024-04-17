from mailmerge import MailMerge
import openpyxl
import os
import pandas as pd

print("*Excel Mail Merge* \nExcel files must follow one of the assessment templates to be compatible \nNotices must be in .docx format, if other, open the file and save as a .docx")
excel_directory = input("Enter excel location: ")
template_directory = input("Enter Docx template location: ")
export_directory = input("Enter export folder: ")

## Excel Parsing
bid = openpyxl.load_workbook(excel_directory, data_only=True)
ws = bid.active
count = 0
table_dict= {}
row_list = []
col_list = []
header_list = ()
for sheet in bid.worksheets:
    for row in sheet.iter_rows():
        for entry in row:
            try:
                if '001.' in entry.value:
                    count = count + 1
                    row_list.append(entry.row)
                    name = str(entry.offset(row=1, column=0).value)
                    address = "{} {}".format(str(entry.offset(row=2, column=0).value), str(entry.offset(row=3, column=0).value))
                    if entry.offset(row=0, column=1).value == None:
                        lot = ""
                    else:
                        lot = str(entry.offset(row=0, column=1).value)
                    block = str(entry.offset(row=0, column=2).value)
                    value = str(entry.offset(row=3, column=15).value)
                    project = str(entry.offset(row=0, column=17).value)
                    project_stripped = project.replace(" ","")
                    block_stripped = " ".join(block.split())
                    lot_stripped = " ".join(lot.split())
                    table_dict[entry.row] = {}
                    table_dict[entry.row]['Pin'] = entry.value
                    table_dict[entry.row]['Name'] = name
                    table_dict[entry.row]['Address'] = address
                    ##need to clean up lot and blocks, sometimes no lot number or worded strangely, clear whitespaces
                    table_dict[entry.row]['Description'] = 'Lot {} Block {}'.format(lot_stripped, block_stripped)
                    table_dict[entry.row]['Project'] = project_stripped
            except (AttributeError, TypeError):
                continue
start_row = row_list[0]
table_headers = start_row - 4
headers = list(ws.iter_rows(max_col=ws.max_column, min_row =table_headers, max_row=table_headers, values_only=True))
print("Count:{}".format(count))

## Mail Merge
template = template_directory
document = MailMerge(template)
merge_list = []
project_list = []
print_count = 0
for key in table_dict:
    if table_dict[key]['Project'] not in project_list:
        project_list.append(table_dict[key]['Project'])
for i in project_list:
    template = template_directory
    document = MailMerge(template)
    merge_list = []
    for key in table_dict:
        if table_dict[key]['Project'] == i:
            print_count = print_count + 1
            merge_list.append({'Pin': table_dict[key]['Pin'], 'Name': table_dict[key]['Name'], 'Description' : table_dict[key]['Description']})
    document.merge_templates(merge_list, separator='page_break')
    export_path = export_directory
    docx_file_name = '{}.docx'.format(i)
    print("Exporting {}...".format(docx_file_name))
    document.write(os.path.join(export_path, docx_file_name))
    merge_list.clear()
print("{} total pages created".format(print_count))

## Write dictionary to table for mailing labels
df = pd.DataFrame.from_dict(table_dict, orient='index')
for i in project_list:
    query_string = 'Project == "{}"'.format(i)
    filter = df.query(query_string) 
    print(filter)
    excel_file_name = "{}.xlsx".format(i)
    filter.to_excel(os.path.join(export_path, excel_file_name))
