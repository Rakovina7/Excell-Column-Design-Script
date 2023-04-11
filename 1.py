import pandas as pd
import webbrowser
from googlesearch import search
import time
import xlsxwriter

# Read the data from the Excel file
file_name = 'sort.xlsx'  # Replace this with the name of your Excel file
sheet_name = 'Sheet1'  # Replace this with the name of the sheet containing the data
df = pd.read_excel(file_name, sheet_name=sheet_name, header=None)

# Process the data
names = []
urls = []
for _, row in df.iterrows():
    line = row[0]
    if isinstance(line, str):
        try:
            name, url = line.split(': ')
            if url.startswith('https://coinmarketcap.com'):
                names.append(name)
                urls.append(url)
            else:
                raise ValueError
        except ValueError:
            print(f"Searching for missing URL for: {line}")
            search_query = 'coin site:coinmarketcap.com'
            query = f"{line} {search_query}"
            while True:
                try:
                    for google_result in search(query, num_results=1):
                        if 'coinmarketcap.com' in google_result:
                            webbrowser.open(google_result)
                            print(f"Found URL: {google_result}")
                            names.append(line)
                            urls.append(google_result)
                        else:
                            print(f"Skipping search result: {google_result}")
                    break
                except Exception as e:
                    if '429' in str(e):
                        print("Encountered a rate limit error. Waiting for 120 seconds...")
                        time.sleep(120)
                    else:
                        print(f"Encountered an error: {e}")
                        urls.append('https://coinmarketcap.com/currencies/unknown')
                        names.append(line)

    else:
        urls.append('https://coinmarketcap.com/currencies/unknown')
        names.append('')

# Create a new DataFrame
processed_df = pd.DataFrame({'Cryptocurrency Name': names, 'URL': urls})

# Write to a new Excel file with clickable links
writer = pd.ExcelWriter('last.xlsx', engine='xlsxwriter')
processed_df.to_excel(writer, index=False)

# Add hyperlink formatting to the URL column
workbook = writer.book
worksheet = writer.sheets['Sheet1']
url_format = workbook.add_format({'font_color': 'blue', 'underline': 1})
worksheet.set_column('B:B', None, url_format, {'url': True})

# Save the Excel file
writer.save()
