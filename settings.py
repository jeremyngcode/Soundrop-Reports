from os import path
import pathlib
# -------------------------------------------------------------------------------------------------

# EDIT ACCORDINGLY
year = 'YYYY'
month = 'MM'
statement_period = f'{year}-{month}'

soundrop_csv_file = pathlib.Path(
	f"C:/path/to/my/latest_soundrop_statement/{statement_period}.csv"
)
template_xl_file = pathlib.Path(
	path.join(path.dirname(__file__), "xl-template.xlsx")
)

# Script outputs to this file path
output_file = pathlib.Path(
	"C:/path/to/my/script_output.xlsx"
)

tax_rate = 0.3
