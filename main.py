from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font
import gspread

gc = gspread.service_account()
new_spreadsheet = gc.create(title='all_new')
new_spreadsheet.add_worksheet(title='xyz', rows=100, cols=100)
worksheets = new_spreadsheet.worksheets()
dc = worksheets[-1]
print(worksheets)
for ws in range(len(worksheets)-1):
    new_spreadsheet.del_worksheet(worksheets[ws])

def copy_sheet_with_format(input_file_1 , input_file_2,  output_file):
    # Load the input workbook
    input_wb_1 = load_workbook(input_file_1)
    ws_list_wb1 = input_wb_1.sheetnames
    print(ws_list_wb1, "here")
    input_wb_2 = load_workbook(input_file_2)
    ws_list_wb2 = input_wb_2.sheetnames
    
    # Create a new workbook and worksheet
    output_wb = Workbook()
    sht_rqr = len(ws_list_wb1)+len(ws_list_wb2)
    o = 0
    for i in range(1,sht_rqr+1):
        output_wb.create_sheet(title=f"Sheet_page{i}")
    output_wb.remove(output_wb['Sheet'])
    
    # copying input 1 wb entirly to output
    for i in range(len(ws_list_wb1)):
        ws = input_wb_1[ws_list_wb1[i]]
        for row in ws.iter_rows(values_only=True):
            output_wb.worksheets[o].append(row)
        o = o+1
        
    for i in range(len(ws_list_wb2)):
        ws = input_wb_2[ws_list_wb2[i]]
        for row in ws.iter_rows(values_only=True):
            output_wb.worksheets[o].append(row)
        o = o+1
    
    o = 0
    # copy merged cells
    for i  in range(len(ws_list_wb1)):
        ws = input_wb_1[ws_list_wb1[i]]
        for merged_cell_range in ws.merged_cells.ranges:
            print(merged_cell_range.coord)
            output_wb.worksheets[o].merge_cells(str(merged_cell_range.coord))
        o = o+1
    for i  in range(len(ws_list_wb2)):
        ws = input_wb_2[ws_list_wb2[i]]
        for merged_cell_range in ws.merged_cells.ranges:
            print(merged_cell_range.coord)
            output_wb.worksheets[o].merge_cells(str(merged_cell_range.coord))
        o = o+1
        
    for ws in output_wb.worksheets:
        # Extract data from the worksheet
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append(row)

        # Extract merged cell information
        merged_cells = ws.merged_cells.ranges

        # Add a new worksheet to the Google Sheets spreadsheet
        new_ws = new_spreadsheet.add_worksheet(title=ws.title, rows=len(data), cols=len(data[0]))

        # Update the new worksheet with the extracted data
        new_ws.update(data, 'A1')

        # Merge cells in the new worksheet based on extracted merged cell information
        for merged_range in merged_cells:
            start_row, start_col, end_row, end_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col
            new_ws.merge_cells(start_row, start_col, end_row, end_col)
    # # Save the output workbook
    output_wb.save(output_file)
    print("Sheet copied successfully with format and merged cells.")

# Example usage:
input_file_1 = "./resources/ws1.xlsx"
input_file_2 = "./resources/ws2.xlsx"
output_file = "./resources/output.xlsx"

# wb1 = load_workbook(input_file_1)
# ws_list_wb1 = wb1.get_sheet_names()


copy_sheet_with_format(input_file_1, input_file_2, output_file)
new_spreadsheet.del_worksheet(dc)
new_spreadsheet.share('ashmit.gupta@impulsecompute.com', perm_type='user', role='writer')

