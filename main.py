import openpyxl

# Load your existing Excel file
workbook = openpyxl.load_workbook('file.xlsx')
sheet = workbook.active

# Open a text file to write the data
with open('data.txt', 'w') as txt_file:
    # Iterate through each row in the sheet
    for row in sheet.iter_rows(values_only=True):
        # Iterate through each cell in the row
        for cell in row:
            # Write cell data to the text file
            txt_file.write(str(cell) + '\n')
        # Add two blank lines between cells
        txt_file.write('\n\n')

print("Data has been written to data.txt.")
