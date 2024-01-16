from decimal import Decimal
import pandas as pd

from settings import *
# -------------------------------------------------------------------------------------------------

def get_wht(csv_file, tax_rate):
	tax_rate = Decimal(str(tax_rate))
	print(f'Withholding tax rate: {tax_rate:%}\n')

	# Load csv file into a dataframe
	soundrop_df = pd.read_csv(
		csv_file, converters={'Amount Due in USD': Decimal}
	)
	soundrop_df_grouped = soundrop_df.groupby(['Country', 'Channel'])['Amount Due in USD'].sum()
	soundrop_df_grouped = soundrop_df_grouped.loc['United States']

	taxable_channels = (
		'Ad-Supported Radio',
		'Ad-Supported Streaming',
		'Locker Service',
		'Subscription Streaming',
		'User Generated Content'
	)

	# Sum up total taxable income from each channel
	print('Summing up total taxable income..')
	taxable_income = 0

	for channel in taxable_channels:
		try:
			taxable_income += soundrop_df_grouped.loc[channel]
		except KeyError:
			print(f'No data for "{channel}".')

	print(f'Total taxable income: ${taxable_income:,.2f}')
	print()

	# Apply tax rate on taxable income
	withholding_tax = taxable_income * tax_rate
	print(f'WITHHOLDING TAX FOR {statement_period}: ${withholding_tax:,.2f}')

	total_income = soundrop_df['Amount Due in USD'].sum()
	print(f'% of total income: {withholding_tax / total_income:.2%}')
	print('-' * 100)

	return round(withholding_tax, 2)

# -------------------------------------------------------------------------------------------------
if __name__ == '__main__':
	get_wht(soundrop_csv_file, tax_rate)
