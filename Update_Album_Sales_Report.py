from decimal import Decimal
import pandas as pd
from openpyxl import load_workbook

from settings import *
from Get_Soundrop_WHT import get_wht
# -------------------------------------------------------------------------------------------------

# EDIT ACCORDINGLY - Catalog release titles
albums_list = {
	'album1' : 'Final Fantasy: Piano Collections',
	'album2' : 'A Disney Collection',
	'album3' : 'Final Fantasy: Piano Flashback',
	'album4' : 'Chopin: Mazurkas',
	'album5' : 'Final Fantasy: Piano Collections, Vol. 2',
	'album6' : 'Anime Piano: Shining in the Sky',
	'album7' : 'Final Fantasy: Piano Collections, Vol. 3',
	'album8' : 'A Disney Collection, Vol. 2',
	'album9' : 'Kingdom Hearts: Piano Collections',
	'album10': 'Final Fantasy: Piano Collections, Vol. 4',
	'album11': 'NieR Gestalt & Replicant Piano Collections',
	'album12': 'Ghibli Piano Collection'
}
singles_list = {
	'single1' : 'The Random Sketches (Nos. 5, 9 & 10)',
	'single2' : 'Mia & Sebastian\'s Theme (From "La La Land")',
	'single3' : 'Learn to Be Lonely (From "The Phantom of the Opera")',
	'single4' : '13 Preludes, Op. 32: XII. Allegro in G-Sharp Minor',
	'single5' : 'Too Good at Goodbyes',
	'single6' : '4 Piano Pieces, Op. 119: I. Intermezzo in B Minor',
	'single7' : '24 Preludes, Op. 28: IV. Largo in E Minor',
	'single8' : 'Thank You, Lord',
	'single9' : 'Lost Sandbar',
	'single10': 'Nocturne No. 20 in C-Sharp Minor, Op. Posth.',
	'single11': 'Final Fantasy III: Piano Opera',
	'single12': 'Final Fantasy: 2 Town Themes',
	'single13': 'Remember Me (From "Coco") [Arranged by Hirohashi Makiko]',
	'single14': 'ICARO -Piano Arrangement- (From "Shadow Hearts: Covenant")',
	'single15': 'Rydia & Edward',
	'single16': 'Chopin: 24 Preludes, Op. 28: Nos. 6 & 7',
	'single17': 'Penelo\'s Theme (From "Final Fantasy XII") [Piano Collections]',
	'single18': 'Voice Of No Return',
	'single19': 'The Tower (From "NieR:Automata") [Piano Collections]',
	'single20': 'Schala\'s Theme (From "Chrono Trigger")',
	'single21': 'The Fading Stories -Qingce Night- (From "Genshin Impact")'
}
albums_list_collabs = {
}
singles_list_collabs = {
	'single_collab1': 'Song of the Ancients (From "NieR:Automata")',
	'single_collab2': 'Merry Go Round Of Life (From "Howl\'s Moving Castle") [Erhu Cover]',
	'single_collab3': 'Always With Me (From "Spirited Away") [Erhu Cover]',
	'single_collab4': 'Vague Hope - Cold Rain (From "NieR:Automata")',
	'single_collab5': 'Carrying You (From "Laputa: Castle in the Sky") [Erhu Cover]',
	'single_collab6': 'Winding River (From "Genshin Impact") [Erhu Cover] ',
	'single_collab7': 'Mononoke Hime (From "Princess Mononoke") [Erhu Cover]'
}

full_catalog = albums_list | singles_list | albums_list_collabs | singles_list_collabs

# Load csv file into a dataframe
soundrop_df = pd.read_csv(
	soundrop_csv_file, converters={'Amount Due in USD': Decimal}
)
summed_revenues = soundrop_df.groupby('Release Title')['Amount Due in USD'].sum()
summed_revenues = summed_revenues.apply(lambda x: round(x, 2))

# Display summed revenues and check if the total number of releases recorded is correct
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

# xl_wb.save(xl_file)
