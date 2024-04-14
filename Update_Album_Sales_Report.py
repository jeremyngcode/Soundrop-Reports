from decimal import Decimal
import pandas as pd
from openpyxl import load_workbook
from xl_styles import *

from settings import *
from catalog_release_titles import *
from Get_Soundrop_WHT import get_wht
# -------------------------------------------------------------------------------------------------

# Load csv file into a dataframe
soundrop_df = pd.read_csv(
	soundrop_csv_file, converters={'Amount Due in USD': Decimal}
)

# For Sheet1 ------------------------------------------------------------------
# Retrieve summed revenues
summed_revenues = soundrop_df.groupby('Release Title')['Amount Due in USD'].sum()
summed_revenues = summed_revenues.apply(lambda x: round(x, 2))

# Display summed revenues and check if the total number of releases recorded is correct
full_catalog = albums_list | singles_list | albums_list_collabs | singles_list_collabs

print(summed_revenues)
print(f'TOTAL NUMBER OF RELEASES: {len(summed_revenues)}/{len(full_catalog)}\n')
print(f'TOTAL REVENUE: ${summed_revenues.sum()}')
print('-' * 100)

# Display tax info
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

# For Sheet2 ------------------------------------------------------------------
# Group for streaming data
soundrop_df_grouped = soundrop_df.groupby(
	['Artist', 'Channel', 'Service']
	)[['Quantity', 'Amount Due in USD']].sum()

soundrop_df_grouped = soundrop_df_grouped.loc['Jeremy Ng']
print(f'SOUNDROP GROUPED DATAFRAME:\n{soundrop_df_grouped}')
print()

soundrop_df_grouped_Sub = soundrop_df_grouped.loc['Subscription Streaming']
soundrop_df_grouped_AdS = soundrop_df_grouped.loc['Ad-Supported Streaming']

pd.set_option('mode.chained_assignment', None)

amazon_services = ('Amazon Ads', 'Amazon Music Unlimited', 'Amazon Prime')
soundrop_df_grouped_Sub.loc['Amazon'] = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service], amazon_services)
)
youtube_services = ('YouTube', 'YouTube Red')
soundrop_df_grouped_Sub.loc['YouTube Sub'] = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service], youtube_services)
)
for services in (amazon_services, youtube_services):
	for service in services:
		soundrop_df_grouped_Sub.drop(service, inplace=True)

for group in (soundrop_df_grouped_Sub, soundrop_df_grouped_AdS):
	group.loc['Spotify', 'Amount Due in USD'] += group.loc['Spotify Discovery Mode', 'Amount Due in USD']
	group.drop('Spotify Discovery Mode', inplace=True)

pd.set_option('mode.chained_assignment', 'warn')

# Group for downloads data
soundrop_df_grouped2 = soundrop_df.groupby(
	['Artist', 'Channel', 'Format', 'Release Title']
	)[['Quantity', 'Amount Due in USD']].sum()

soundrop_df_grouped_Dl = soundrop_df_grouped2.loc['Jeremy Ng', 'Download']
print(f'SOUNDROP DOWNLOADS DATA:\n{soundrop_df_grouped_Dl}')
print()

print(f'ALBUM TOTAL DOWNLOADS:\n{soundrop_df_grouped_Dl.loc["Album "].sum()}')
print(f'TRACK TOTAL DOWNLOADS:\n{soundrop_df_grouped_Dl.loc["Track "].sum()}')
print(f'TOTAL DOWNLOAD REVENUE: {soundrop_df_grouped_Dl["Amount Due in USD"].sum()}')
print()
# -------------------------------------------------------------------------------------------------



# Load template file
xl_wb = load_workbook(template_xl_file)
sheet1 = xl_wb.worksheets[0]
sheet2 = xl_wb.worksheets[1]

# Write Sheet1 ----------------------------------------------------------------
starting_row = 3
columns = (('A', 'B'), ('D', 'E'))

solo = (albums_list, singles_list)
collabs = (albums_list_collabs, singles_list_collabs)

# Write summed revenues
print('Writing summed revenues..')

for i, (col_pair, category) in enumerate(zip(columns, (solo, collabs))):
	title_col, rev_col = col_pair[0], col_pair[1]

	for j, release_list in enumerate(category):
		if not release_list:
			row = starting_row

		for row, release_title in enumerate(release_list.values(), starting_row):
			sheet1[f'{title_col}{row}'] = release_title
			sheet1[f'{rev_col}{row}'] = summed_revenues.get(release_title, '[INVALID KEY]')
		else:
			for cell in (f'{title_col}{row+1}', f'{rev_col}{row+1}'):
				sheet1[cell].fill = grey_fill

			title_cell = sheet1[f'{title_col}{row+2}']

			if j == 0:
				for cell in (f'{title_col}{row+2}', f'{rev_col}{row+2}'):
					sheet1[cell].fill = orange_fill

				title_cell.font = biu
				title_cell.value = 'Singles'
				if i == 1:
					title_cell.value += ' (Collabs)'
			else:
				title_cell.font = boldunderline
				title_cell.value = 'TOTAL Revenue (USD):'

				sheet1[f'{rev_col}{row+2}'] = f'=SUM({rev_col}3:{rev_col}{row})'

		starting_row = row + 3

	starting_row = 3

# Update sheet title
sheet1.title = statement_period
print()

# Write Sheet2 ----------------------------------------------------------------
col = 'B'

# Write Subscription Streaming data
print('Writing Subscription Streaming data..')

services_Sub = (
	'Apple Music',
	'Spotify',
	'Amazon',
	'YouTube Sub',
	'Deezer'
)

rows = (9, 14, 19, 24, 29)
for row, service in zip(rows, services_Sub):
	sheet2[f'{col}{row}'] = soundrop_df_grouped_Sub.loc[service, 'Quantity']

	row += 1
	sheet2[f'{col}{row}'] = round(
		soundrop_df_grouped_Sub.loc[service, 'Amount Due in USD'], 2
	)

# Write totals for Subscription Streaming data
sheet2[f'{col}4'] = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service, 'Quantity'], services_Sub)
)
rev_total_Sub = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service, 'Amount Due in USD'], services_Sub)
)
sheet2[f'{col}5'] = round(rev_total_Sub, 2)

# Write Ad-Supported Streaming data
print('Writing Ad-Supported Streaming data..')

services_AdS = (
	'Spotify',
)

rows = (40,)
for row, service in zip(rows, services_AdS):
	sheet2[f'{col}{row}'] = soundrop_df_grouped_AdS.loc[service, 'Quantity']

	row += 1
	sheet2[f'{col}{row}'] = round(
		soundrop_df_grouped_AdS.loc[service, 'Amount Due in USD'], 2
	)

# Write formula values for streaming data
print('Writing formula values for streaming data..')

rows = (11, 16, 21, 26, 31, 6, 42)
for row in rows:
	sheet2[f'{col}{row}'] = f'=({col}{row-1}/{col}{row-2})*1000'

# Write Downloads data
print('Writing Downloads data..')

sheet2[f'{col}50'] = soundrop_df_grouped_Dl.loc['Track ']['Quantity'].sum()
sheet2[f'{col}51'] = round(soundrop_df_grouped_Dl.loc['Track ']['Amount Due in USD'].sum(), 2)

sheet2[f'{col}53'] = soundrop_df_grouped_Dl.loc['Album ']['Quantity'].sum()
sheet2[f'{col}54'] = round(soundrop_df_grouped_Dl.loc['Album ']['Amount Due in USD'].sum(), 2)

sheet2[f'{col}56'] = round(soundrop_df_grouped_Dl['Amount Due in USD'].sum(), 2)

# Update sheet title
sheet2.title = statement_period + ' (Streaming + Downloads)'
print()

# Write reporting month as header ---------------------------------------------
header_cells = (sheet1['B1'], sheet1['E1'], sheet2['B1'])

for cell in header_cells:
	cell.value = f'/ {month}'

	if month == '01':
		cell.value = f'{year} / {month}'
		cell.fill = yellow_fill
# -------------------------------------------------------------------------------------------------

xl_wb.save(output_file)
