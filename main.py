import os
import pandas as pd
from openpyxl import load_workbook


# for any input file
    # find matching input row
    # find rent column
    # loop thru rent column and copy rents

print('==============================================================')
print('Welcome! Make sure to place your files in the "data" directory')
print('-----------------------------------')
print('Please select an option:')
print('1. Update rents')
print('2. Add new property')
#option = input('Your option: ')
option = '1'

if option == '1':
    print('You chose to update rent values.')
    print('-----------------------------------')

    # Will run once for each file in data directory
    for file in os.listdir('data'):
        print('PROCESSING FILE: ', file)

        # Open "Lbk Comp" Excel file
        wb = load_workbook('data/' + file)
        ws = wb.active

        # Locate "Avg Asking Rent/Unit" column
        rent_col = None
        for col in ws.iter_cols(1, ws.max_column):
            header = col[0].value
            if header == 'Avg Asking Rent/Unit':
                rent_col = col
                break

        # If rent column exists
        if rent_col:
            # Create list of all values in rent column
            # Uses list comprehension, ignores empty cells
            all_rent = [cell.value for cell in rent_col if cell.value is not None]

            df = pd.DataFrame(all_rent[1:], columns=['Avg Asking Rent/Unit'])

        print(df.head())



