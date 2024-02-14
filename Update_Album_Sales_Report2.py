from decimal import Decimal
import pandas as pd
from openpyxl import load_workbook

from settings import *
# -------------------------------------------------------------------------------------------------

# Load csv file into a dataframe
soundrop_df = pd.read_csv(
	soundrop_csv_file, converters={'Amount Due in USD': Decimal}
)

# Group for streaming data
soundrop_df_grouped = soundrop_df.groupby(
	['Artist', 'Channel', 'Service']
	)[['Quantity', 'Amount Due in USD']].sum()

soundrop_df_grouped = soundrop_df_grouped.loc['Jeremy Ng']
print(f'SOUNDROP GROUPED DATAFRAME:\n{soundrop_df_grouped}')
print()

soundrop_df_grouped_Sub = soundrop_df_grouped.loc['Subscription Streaming']
soundrop_df_grouped_AdS = soundrop_df_grouped.loc['Ad-Supported Streaming']

amazon_services = ('Amazon Ads', 'Amazon Music Unlimited', 'Amazon Prime')
soundrop_df_grouped_Sub.loc['Amazon'] = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service], amazon_services)
)
youtube_services = ('YouTube', 'YouTube Red')
soundrop_df_grouped_Sub.loc['YouTube Sub'] = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service], youtube_services)
)

# Group for downloads data
soundrop_df_grouped2 = soundrop_df.groupby(
	['Artist', 'Channel', 'Format', 'Release Title']
	)[['Quantity', 'Amount Due in USD']].sum()

soundrop_df_grouped2 = soundrop_df_grouped2.loc['Jeremy Ng', 'Download']
print(f'SOUNDROP DOWNLOADS DATA:\n{soundrop_df_grouped2}')
print()

print(f'ALBUM TOTAL DOWNLOADS:\n{soundrop_df_grouped2.loc["Album "].sum()}')
print(f'TRACK TOTAL DOWNLOADS:\n{soundrop_df_grouped2.loc["Track "].sum()}')
print(f'TOTAL DOWNLOAD REVENUE: {soundrop_df_grouped2["Amount Due in USD"].sum()}')
print()



# Open Excel workbook
xl_wb = load_workbook(xl_file)
xl_sheet = xl_wb.worksheets[1]
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
for (row, service) in zip(rows, services_Sub):
	xl_sheet[f'{col}{row}'] = soundrop_df_grouped_Sub.loc[service, 'Quantity']

	row += 1
	xl_sheet[f'{col}{row}'] = round(
		soundrop_df_grouped_Sub.loc[service, 'Amount Due in USD'], 2
	)

# Write totals for Subscription Streaming data
xl_sheet[f'{col}4'] = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service, 'Quantity'], services_Sub)
)
rev_total_Sub = sum(
	map(lambda service: soundrop_df_grouped_Sub.loc[service, 'Amount Due in USD'], services_Sub)
)
xl_sheet[f'{col}5'] = round(rev_total_Sub, 2)

# Write Ad-Supported Streaming data
print('Writing Ad-Supported Streaming data..')

services_AdS = (
	'Spotify',
)

rows = (40,)
for (row, service) in zip(rows, services_AdS):
	xl_sheet[f'{col}{row}'] = soundrop_df_grouped_AdS.loc[service, 'Quantity']

	row += 1
	xl_sheet[f'{col}{row}'] = round(
		soundrop_df_grouped_AdS.loc[service, 'Amount Due in USD'], 2
	)

# Write formula values for streaming data
print('Writing formula values for streaming data..')

rows = (11, 16, 21, 26, 31, 6, 42)
for row in rows:
	xl_sheet[f'{col}{row}'] = f'=({col}{row-1}/{col}{row-2})*1000'

# Write Downloads data
print('Writing Downloads data..')

xl_sheet[f'{col}50'] = soundrop_df_grouped2.loc['Track ']['Quantity'].sum()
xl_sheet[f'{col}51'] = round(soundrop_df_grouped2.loc['Track ']['Amount Due in USD'].sum(), 2)

xl_sheet[f'{col}53'] = soundrop_df_grouped2.loc['Album ']['Quantity'].sum()
xl_sheet[f'{col}54'] = round(soundrop_df_grouped2.loc['Album ']['Amount Due in USD'].sum(), 2)

xl_sheet[f'{col}56'] = round(soundrop_df_grouped2['Amount Due in USD'].sum(), 2)
print()

# Update sheet title
xl_sheet.title = statement_period + ' (Streaming + Downloads)'

xl_wb.save(xl_file)
