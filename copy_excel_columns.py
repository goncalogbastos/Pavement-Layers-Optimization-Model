import openpyxl

def copy_excel_columns(source_file, target_file, source_columns, target_columns):
  # Load the workbook objects for the source and target files
  source_wb = openpyxl.load_workbook(source_file)
  target_wb = openpyxl.load_workbook(target_file)

  # Get the sheet names in the source and target workbooks
  source_sheet_names = source_wb.sheetnames
  target_sheet_names = target_wb.sheetnames

  # Iterate through the sheets in the source workbook
  for i, sheet_name in enumerate(source_sheet_names):
    # Get the sheet objects for the current sheet in the source and target workbooks
    source_sheet = source_wb[sheet_name]
    target_sheet = target_wb[target_sheet_names[i]]

    # Iterate through the rows in the source sheet
    for row in source_sheet.iter_rows():
      # Iterate through the cells in the source row
      for j, cell in enumerate(row):
        # If the current cell is in one of the source columns, copy its value to the corresponding cell in the target columns
        if j in source_columns:
          # Delete the value from the target cell
          target_cell = target_sheet.cell(row=cell.row, column=target_columns[source_columns.index(j)])
          target_cell.value = None

          # Copy the value from the source cell to the target cell
          target_cell.value = cell.value
          cell.value = None  # delete the value from the source cell

  # Save the changes to the target workbook
  target_wb.save(target_file)




copy_excel_columns('./output/layerSolutions_left.xlsx', './src/layerSolutions_left_Civil.xlsx', [1,2,3,4], [3,4,5,6])
copy_excel_columns('./output/layerSolutions_left.xlsx', './src/layerSolutions_left_Civil.xlsx', [5,6,7,8], [8,9,10,11])
copy_excel_columns('./output/layerSolutions_left.xlsx', './src/layerSolutions_left_Civil.xlsx', [9,10,11,12], [13,14,15,16])
copy_excel_columns('./output/layerSolutions_left.xlsx', './src/layerSolutions_left_Civil.xlsx', [13,14,15,16], [18,19,20,21])

copy_excel_columns('./output/layerSolutions_right.xlsx', './src/layerSolutions_right_Civil.xlsx', [1,2,3,4], [3,4,5,6])
copy_excel_columns('./output/layerSolutions_right.xlsx', './src/layerSolutions_right_Civil.xlsx', [5,6,7,8], [8,9,10,11])
copy_excel_columns('./output/layerSolutions_right.xlsx', './src/layerSolutions_right_Civil.xlsx', [9,10,11,12], [13,14,15,16])
copy_excel_columns('./output/layerSolutions_right.xlsx', './src/layerSolutions_right_Civil.xlsx', [13,14,15,16], [18,19,20,21])