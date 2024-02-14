import pathlib
# -------------------------------------------------------------------------------------------------

# EDIT ACCORDINGLY
statement_period = 'YYYY-MM'

soundrop_csv_file = pathlib.Path(
	f"C:/path/to/my/latest_soundrop_statement/{statement_period}.csv"
)

# Script writes to this file
xl_file = pathlib.Path(
	"C:/path/to/my/excel_file.xlsx"
)

tax_rate = 0.3
