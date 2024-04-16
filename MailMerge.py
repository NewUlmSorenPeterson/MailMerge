from mailmerge import MailMerge
import openpyxl
import os

bid = openpyxl.load_workbook(r'C:\Users\soren.peterson\Desktop\Tempshapes\2024_04_16\TestDocument.xlsx', data_only=True)
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
                    name = entry.offset(row=1, column=0).value
                    address = "{} {}".format(entry.offset(row=2, column=0).value, entry.offset(row=3, column=0).value)
                    lot = str(entry.offset(row=0, column=1).value)
                    block = entry.offset(row=0, column=2).value
                    value = entry.offset(row=3, column=15).value
                    project = entry.offset(row=0, column=17).value
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

template = r"C:\Users\soren.peterson\Desktop\Tempshapes\2024_04_16\sample.docx"
document = MailMerge(template)
merge_list = []
project_list = []
for key in table_dict:
    if table_dict[key]['Project'] not in project_list:
        project_list.append(table_dict[key]['Project'])
for i in project_list:
    template = r"C:\Users\soren.peterson\Desktop\Tempshapes\2024_04_16\sample.docx"
    document = MailMerge(template)
    merge_list = []
    for key in table_dict:
        if table_dict[key]['Project'] == i:
            merge_list.append({'Pin': table_dict[key]['Pin'], 'Name': table_dict[key]['Name'], 'Description' : table_dict[key]['Description']})
    document.merge_templates(merge_list, separator='page_break')
    export_path = r'C:\Users\soren.peterson\Desktop\Tempshapes\2024_04_16'
    file_name = '{}.docx'.format(i)
    print("Exporting {}...".format(file_name))
    document.write(os.path.join(export_path, file_name))
    merge_list.clear()
