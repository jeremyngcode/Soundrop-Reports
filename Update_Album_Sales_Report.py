from decimal import Decimal
import pandas as pd
from openpyxl import load_workbook

from settings import *
from catalog_release_titles import *
from Get_Soundrop_WHT import get_wht
# -------------------------------------------------------------------------------------------------

# Load csv file into a dataframe
soundrop_df = pd.read_csv(
	soundrop_csv_file, converters={'Amount Due in USD': Decimal}
)
summed_revenues = soundrop_df.groupby('Release Title')['Amount Due in USD'].sum()
summed_revenues = summed_revenues.apply(lambda x: round(x, 2))

# Display summed revenues and check if the total number of releases recorded is correct
full_catalog = albums_list | singles_list | albums_list_collabs | singles_list_collabs

print(summed_revenues)
print(f'TOTAL NUMBER OF RELEASES: {len(summed_revenues)}/{len(full_catalog)}\n')
print(f'TOTAL REVENUE: ${summed_revenues.sum()}')
print('-' * 100)

withholding_tax = get_wht(soundrop_csv_file, tax_rate)
after_tax_income = summed_revenues.sum() - withholding_tax
print(f'AFTER-TAX INCOME: ${after_tax_income}')
print()

# Check for release titles that do not exist in full_catalog.values()
not_found_in_catalog = [
	release_title for release_title in summed_revenues.index
	if release_title not in full_catalog.values()
]

if not_found_in_catalog:
	print('RELEASE TITLES NOT FOUND IN CATALOG:')
	for i, release_title in enumerate(not_found_in_catalog, 1):
		print(f'{i}.', release_title)
	print()



# Open Excel workbook
xl_wb = load_workbook(xl_file)
xl_sheet = xl_wb.worksheets[0]

# Delete previous data
for row in xl_sheet.iter_rows(min_row=2, min_col=1, max_col=11):
	for cell in row:
		cell.value = None

# Write summed revenues
print('Writing summed revenues..')

columns = (('A', 'B'), ('D', 'E'), ('G', 'H'), ('J', 'K'))
categories = (albums_list,  singles_list, albums_list_collabs, singles_list_collabs)

for (column, category) in zip(columns, categories):
	for (row, release_title) in enumerate(category.values(), 2):
		xl_sheet[f'{column[0]}{row}'] = release_title
		xl_sheet[f'{column[1]}{row}'] = summed_revenues.get(release_title, '[INVALID KEY]')
print()

# Update sheet title
xl_sheet.title = statement_period

xl_wb.save(xl_file)
