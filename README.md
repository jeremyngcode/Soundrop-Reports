Soundrop Reports
================

Intro
-----
[I'm a pianist](https://open.spotify.com/artist/6mdGjVrAY95ecXnVgtefti) with music distributed by Soundrop. Every month, Soundrop sends a comprehensive CSV revenue report where I will filter and copy specific data I'm interested in onto an existing master Excel spreadsheet. However, as my music catalog grew over the years, this manual process started becoming very time-consuming.

One of the reasons I picked up coding in July 2023 was so that I can write scripts to automate tedious / time-consuming tasks. So naturally, this was one of the first few projects I coded.

The Process
-----------
I maintain a list of my full catalog in [catalog_release_titles.py](catalog_release_titles.py) in a specific order that aligns with my master Excel spreadsheet. The titles here should have an identical match to the ones provided in the CSV report.

In [settings.py](settings.py), `output_file` is the output path of the script, which I will then copy over to my master Excel file with a few copy-pastes. 

The regular monthly process looks like this:
1. Change `year` and `month` variables in settings.py to the reporting year and month respectively. `statement_period` is then derived from those variables.
2. Save the given CSV file in the same directory as previous CSV files as `statement_period`.csv.
   ```py
   # Example
   year = '2024'
   month = '01'
   statement_period = f'{year}-{month}' # 2024-01

   soundrop_csv_file = pathlib.Path(
	   f"C:/path/to/my/latest_soundrop_statement/{statement_period}.csv" # 2024-01.csv
   )
   ```
3. Run [Update_Album_Sales_Report.py](Update_Album_Sales_Report.py).

The script will turn this template... ([xl-template.xlsx](xl-template.xlsx))

![xl-template-sheet1](https://github.com/jeremyngcode/Soundrop-Reports/assets/156220343/e4c01d92-36e6-4004-bcee-78484a0bcdf4)

into this... (sheet 1)

![xl-template-sheet1-filled](https://github.com/jeremyngcode/Soundrop-Reports/assets/156220343/6e751ed6-bdb8-4b00-982b-cbb9fd52cb4e)

And also this... (sheet 2)

![xl-template-sheet2](https://github.com/jeremyngcode/Soundrop-Reports/assets/156220343/44127d1a-7844-49b5-8d49-b46694fee9b9) ![xl-template-sheet2-filled](https://github.com/jeremyngcode/Soundrop-Reports/assets/156220343/d805663d-7ce9-4602-928c-105ff8336ea1)

4. Copy-paste columns B and E from sheet 1 and column B from sheet 2 onto the master Excel spreadsheet.
5. Save the file and that's it!

Besides writing data, the script also prints out other info such as total revenue and after-tax income for example. (So I can see my 💸💸💸 immediately upon receiving the CSV file! 🤑)

Extra Thoughts
--------------
- I did consider writing directly onto the separate master Excel spreadsheet since that would eliminate all the copy-pasting which is the most time-consuming part of the process. But I've decided this extra efficiency isn't worth the trade-off of risking having my code mess up something.

- This is my very first repo! Yes, I had only starting learning Git (and Markdown!) a week ago! 😃

- **2024-02-21 Update:** It's been about a month now since I first created this repo. I've just merged the previous two scripts into one file and added some styles. I'm realizing I still enjoy looking at this code and the logic in it. Maybe I love writing loops. 😆

#### Notable libraries used / learned for this project:
- [pandas](https://pypi.org/project/pandas/)
- [openpyxl](https://pypi.org/project/openpyxl/)
