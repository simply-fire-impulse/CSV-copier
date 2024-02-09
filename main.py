import gspread

gc = gspread.service_account()

sh_src = gc.open('CSV-Formatting')
sh_dst = gc.open('CSV-Formatting-1')

ws_src = sh_src.get_worksheet(0)
ws_dst = sh_dst.get_worksheet(0)

# src sheet meta data
sh_src_md = sh_src.fetch_sheet_metadata()

data = ws_src.get_all_values()

# Clear existing data from the destination worksheet (optional)
ws_dst.clear()

# Update the destination worksheet with the retrieved values
ws_dst.update(data,'A1')

# extracting the data for merged cells from the sheet metadata
merged_cells = sh_src_md['sheets'][0]['merges']

merged_cell_ranges = []
for merged_cell in merged_cells:
    # Convert grid range to A1 notation
    start_cell = gspread.utils.rowcol_to_a1(merged_cell['startRowIndex'] + 1, merged_cell['startColumnIndex'] + 1)
    end_cell = gspread.utils.rowcol_to_a1(merged_cell['endRowIndex'], merged_cell['endColumnIndex'])
    # Get value of the top-left cell of the merged range
    str = f"{start_cell}:{end_cell}"
    merged_cell_ranges.append(str)

for i in range(len(merged_cell_ranges)):
    ws_dst.merge_cells(merged_cell_ranges[i])