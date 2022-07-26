# Python 3.8.10
# Parces city data from defined URL via Excel spreadsheet 'web_to_spreadsheet.xls'

import bs4, requests, csv
import pandas as pd


# Extract data from Excel spreadsheet
read_file = pd.read_excel(r'web_to_spreadsheet.xls')

# Convert data to CSV formatted file
read_file.to_csv(r'kg_cities_26072022.csv', index=True, header=True)


# Extract links as text for each city
url = 'https://ru.wikipedia.org/wiki/%D0%93%D0%BE%D1%80%D0%BE%D0%B4%D0%B0_%D0%9A%D0%B8%D1%80%D0%B3%D0%B8%D0%B7%D0%B8%D0%B8'
response = requests.get(url)
response.raise_for_status()
bs_obj = bs4.BeautifulSoup(response.text, features='html.parser')

city_links = []
for link in bs_obj.find_all('a'):
	city_links.append(link.get('href'))

with open('kg_cities_26072022.csv', 'r') as f:
	csv_reader = csv.reader(f)
	csv_list = []
	for row in csv_reader:
		csv_list.append(row)

# Add link to each city
csv_list[1].append(city_links[25])
csv_list[2].append(city_links[29])
csv_list[3].append(city_links[33])
csv_list[4].append(city_links[36])
csv_list[5].append(city_links[39])
csv_list[6].append(city_links[43])
csv_list[7].append(city_links[44])
csv_list[8].append(city_links[45])
csv_list[9].append(city_links[49])
csv_list[10].append(city_links[51])
csv_list[11].append(city_links[54])
csv_list[12].append(city_links[57])
csv_list[13].append(city_links[58])
csv_list[14].append(city_links[61])
csv_list[15].append(city_links[62])
csv_list[16].append(city_links[64])
csv_list[17].append(city_links[65])
csv_list[18].append(city_links[66])
csv_list[19].append(city_links[68])
csv_list[20].append(city_links[70])
csv_list[21].append(city_links[73])
csv_list[22].append(city_links[75])
csv_list[23].append(city_links[76])
csv_list[24].append(city_links[79])
csv_list[25].append(city_links[80])
csv_list[26].append(city_links[83])
csv_list[27].append(city_links[85])
csv_list[28].append(city_links[88])
csv_list[29].append(city_links[89])
csv_list[30].append(city_links[90])
csv_list[31].append(city_links[93])

with open('kg_cities_26072022.csv','w', newline='') as file:
	csv_file = csv.writer(file, delimiter='\n')
	csv_file.writerows([csv_list])