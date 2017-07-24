import datetime
import os

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from multiprocessing.dummy import Pool as ThreadPool


from requests.exceptions import ProxyError

url = 'https://www.us-proxy.org'
proxies = {}
counter = 1
soup = BeautifulSoup(requests.get(url).text, "lxml")
table = soup.find('table', {'id': 'proxylisttable'})
rows = table.find_all('tr')
rows = rows[1:]

for r in rows:
    tds = r.find_all('td')
    if len(tds) != 0:
        if tds[6].text == 'yes' and tds[4].text == 'anonymous':
            item = {
                str(counter): "https://" + str(tds[0].text) + ":" + str(tds[1].text)
            }
            proxies.update(item)
            counter += 1
counter = 1
proxies = {'https': proxies[str(counter)]}
headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive"
}
body = {"page": 1, "ResultsPerPage": 25, "resultsPerPage": 16, "specials": [{"key": "europe-and-the-americas"}],
        "IsSpecialOfferPage": "true", "cacheDate": "2017-01-03T04:50:06.4024627-05:00", "cacheName": "null"}
session = requests.session()
session.headers.update(headers)
pool = ThreadPool(10)
url = 'https://www.oceaniacruises.com/api/cruisefinder/getcruises'
page = ''
notSucc = True

while notSucc:
    try:
        proxies = {'https': proxies[str(counter)]}
        page = requests.post(url=url, proxies=proxies)
        notSucc = False
    except ProxyError:
        counter += 1
        notSucc = True

cruise_results = page.json()
counter = 1
to_write = []
ots = []


def calculate_days(sail_date_param, number_of_nights_param):
    try:
        date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))
    except ValueError:
        split = sail_date_param.split('/')
        sail_date_param = split[1] + '/' + split[0] + '/' + split[2]
        date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))

    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def get_vessel_id(ves_name):
    if ves_name == "Insignia":
        return "429"
    if ves_name == "Marina":
        return "700"
    if ves_name == "Nautica":
        return "495"
    if ves_name == "Regatta":
        return "430"
    if ves_name == "Riviera":
        return "770"
    if ves_name == "Sirena":
        return "938"


def match_by_meta(ports_visited):
    bermuda = ['Hamilton', 'St. George']
    hawaii = ['Hilo', 'Honolulu', 'Kahului', 'Nawiliwili']
    panama_canal = ['Colon', 'Fuerte Amador']
    mexico = ['Acapulco', 'Huatulco', 'Cabo San Lucas']
    west_med = ['Ajaccio', 'Alicante', 'Almeria', 'Amalfi', 'Antibes', 'Arrecife', 'Bandol', 'Barcelona', 'Bastia',
                'Belfast', 'Biarritz', 'Bilbao', 'Bordeaux', 'Brest', 'Cagliari', 'Calvi', 'Cannes', 'Cartagena',
                'Casablanca', 'Catania', 'Cinque Terre', 'Cork', 'Corner Brook', 'Dublin', 'Florence', 'Funchal',
                'Gaeta', 'Gibraltar', 'Gijon', 'Huelva', 'Ibiza', 'La Coruna', 'La Rochelle',
                'Las Palmas de Gran Canaria', 'Lisbon', 'London', 'Lorient', 'Mahon', 'Malaga', 'Messina',
                'Monte Carlo', 'Montreal', 'Naples', 'Olbia', 'Oporto', 'Palamos', 'Palermo', 'Palma de Mallorca',
                'Paris', 'Porto Santo Stefano', 'Portofino', 'Port-Vendres', 'Provence', 'Ravenna', 'Rome', 'Roses',
                'Saint-Pierre', 'Saint-Tropez', 'Santa Cruz de La Palma', 'Santa Cruz de Tenerife',
                'Santiago de Compostela', 'Sete', 'Seville', 'Sorrento', 'St. Peter Port', 'Tangier', 'Taormina',
                'Toulon', 'Trois-Rivieres', 'Umbria', 'Valencia', 'Venice', 'Villefranche']
    east_med = ['Argostoli', 'Athens', 'Chania', 'Corfu', 'Dubrovnik', 'Gythion', 'Heraklion', 'Jerusalem', 'Katakolon',
                'Koper', 'Kotor', 'Limassol', 'Monemvasia', 'Mykonos', 'Patmos', 'Rhodes', 'Rijeka', 'Santorini',
                'Split', 'Thessaloniki', 'Tirana', 'Valletta', 'Volos', 'Zadar', 'Zakynthos']
    exotics = ['Aqaba', 'Dubai', 'Luxor', 'Muscat', 'Salalah', 'Sharm El Sheikh']
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in bermuda:
            return 'Bermuda'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in hawaii:
            return 'Hawaii'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in panama_canal:
            return "Panama Canal"
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in mexico:
            return 'Mexico'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in west_med:
            return 'West Mediterranean'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in east_med:
            return 'East Mediterranean'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in exotics:
            return 'Exotics'
    return 'Carib'


def get_destination(param):
    if param == 'South Pacific & Tahiti':
        return ['SoPac&Tahiti', 'I', 'I']
    elif param == 'Australia & New Zealand':
        return ['SoPac-AU', 'P', "AU/NZ"]
    elif param == 'Hawaii':
        return ['Hawaii', 'H', 'H']
    elif param == 'Alaska':
        return ['Alaska', 'A', 'A']
    elif param == 'Panama Canal':
        return ['Panama Canal', 'T', 'T']
    elif param == 'South America & Amazon':
        return ['South America & Amazon', 'S', 'S']
    elif param == 'Bermuda':
        return ['Bermuda', 'BM', 'BM']
    elif param == 'Canada & New England':
        return ['Canada & New England', 'NN', 'NN']
    elif param == 'Baltic & Scandinavia':
        return ['Europe', 'E', 'BL']
    elif param == 'Mexico':
        return ['Mexico', 'M', 'M']
    elif param == 'Asia':
        return ['Exotics', 'O', 'AS']
    elif param == 'Africa':
        return ['Exotics', 'O', "AF"]
    elif param == 'Cuba':
        return ['Cuba', 'C', 'CU']
    elif param == 'East Caribbean':
        return ['East Caribbean', 'EC', 'EC']
    elif param == 'West Caribbean':
        return ['West Caribbean', 'WC', 'WC']
    elif param == 'Carib':
        return ['Caribbean', 'C', 'C']
    elif param == 'Grand Voyages':
        return [param, "OT", 'OT']
    elif param == 'Transoceanic Voyages':
        return [param, "OT", "OT"]
    elif param == 'West Mediterranean':
        return ["West Mediterranean", "E", 'WM']
    elif param == "East Mediterranean":
        return ['East Mediterranean', 'E', 'EM']
    elif param == '180-Day World Cruises':
        return ['180-Day World Cruises', 'WC']
    else:
        return [param, "OT", 'OT']


def split_carib(ports):
    cu = ['Santiago de Cuba', 'Cienfuegos', 'Havana']
    wc = ['Costa Maya', 'Cozumel', 'Falmouth, Jamaica', 'George Town, Grand Cayman',
          'Ocho Rios']

    ec = ['Basseterre, St. Kitts', 'Bridgetown', 'Castries', 'Charlotte Amalie, St. Thomas',
          'Fort De France', 'Kingstown, St. Vincent', 'Philipsburg', 'Ponce, Puerto Rico',
          'Punta Cana, Dominican Rep', 'Roseau', 'San Juan', 'St. Croix, U.S.V.I.',
          "St. George's", "St. John's", 'Tortola, B.V.I']

    bm = ['Kings Wharf, Bermuda']
    result = []
    iscu = False
    isec = False
    iswc = False
    ports_list = []
    for i in range(len(ports)):
        if i == 0:
            pass
        else:
            ports_list.append(ports[i]['name'])
    for element in cu:
        for p in ports_list:
            if p in element:
                iscu = True
    if not iscu:
        for element in wc:
            for p in ports_list:
                if p in element:
                    iswc = True
    if not iswc:
        for element in ec:
            for p in ports_list:
                if p in element:
                    isec = True
    if iscu:
        result.append("Cuba")
        result.append("C")
        result.append("CU")
        return result
    elif iswc:
        result.append("West Carib")
        result.append("C")
        result.append("WC")
        return result
    elif isec:
        result.append("East Carib")
        result.append("C")
        result.append("EC")
        return result
    else:
        result.append("Carib")
        result.append("C")
        result.append("")
        return result


def parse(line):
    brochure_name = line['voyageName']
    sail_date = line['name'].split(' | ')[0]
    number_of_nights = line['cruiseLength']
    vessel_name = line['shipName']
    vessel_id = get_vessel_id(vessel_name)
    cruise_id = '10'
    itinerary_id = ''
    cruise_line_name = "Oceania Cruise Line"
    port_list = line['ports']
    destination_name = line['destinationName']
    if 'Caribbean, Panama Canal & Mexico' in destination_name:
        destination_name = match_by_meta(port_list)
    elif 'Mediterranean' in destination_name:
        destination_name = match_by_meta(port_list)
    destination = get_destination(destination_name)
    destination_name = destination[0]
    destination_code = destination[1]
    if destination_name == 'Caribbean':
        destination = split_carib(port_list)
    if len(destination) > 2:
        subcode = destination[2]
    else:
        subcode = ''
    if destination_code == 'NN':
        ports = []
        for i in range(len(port_list)):
            if i == 0:
                pass
            else:
                ports.append(port_list[i]['name'])
        for p in ports:
            if 'Hamilton' in p or 'St. George' in p:
                destination_code = "BM"
                destination_name = "Bermuda"
                subcode = 'BM'
    details_url = "https://www.oceaniacruises.com" + line['cruiseDetailsUrl']
    details_page = requests.get(details_url, headers=headers, proxies=proxies)
    soup = BeautifulSoup(details_page.text, "lxml")
    intl = soup.find_all("span")
    for row in intl:
        if row.text == "Int'l Date Line East":
            number_of_nights -= 1
        elif row.text == "Int'l Date Line West":
            number_of_nights += 1
    return_date = calculate_days(sail_date, number_of_nights)
    suite_prices = []
    balcony_prices = []
    oceanview_prices = []
    interior_prices = []
    for tmplink in soup.find_all("tbody"):
        children = tmplink.find_all("tr")
        room_type = ''
        for child in children:
            if "class" in child.attrs:
                if child['class'][0] == "category-heading":
                    if child.find('td', {"colspan": "6"}):
                        ch = child.find_all("td", {"colspan": "6"})
                    else:
                        ch = child.find_all("td", {"colspan": "7"})
                    for c in ch:
                        room_type = c.text.strip()
                if child['class'][0] == "category-row":
                    ch = child.find_all("td", {"class": "fare-fare2"})
                    if len(ch) > 1:
                        ch = ch[0]
                    else:
                        ch = ch[0]
                    if room_type == "Suites":
                        suite_prices.append(ch.text)
                    elif room_type == "Veranda":
                        balcony_prices.append(ch.text)
                    elif room_type == "Ocean View":
                        oceanview_prices.append(ch.text)
                    elif room_type == "Inside Staterooms":
                        interior_prices.append(ch.text)
    try:
        interior_bucket_price = (interior_prices[len(interior_prices) - 1]).replace("$", "").replace(',', '')
        balcony_bucket_price = (balcony_prices[len(balcony_prices) - 1]).replace("$", "").replace(',', '')
        oceanview_bucket_price = (oceanview_prices[len(oceanview_prices) - 1]).replace("$", "").replace(',', '')
        suite_bucket_price = (suite_prices[len(suite_prices) - 1]).replace("$", "").replace(',', '')
        temp = [destination_code, destination_name, subcode, vessel_id, vessel_name, cruise_id, cruise_line_name,
                itinerary_id, brochure_name, number_of_nights, sail_date, return_date, interior_bucket_price,
                oceanview_bucket_price, balcony_bucket_price, suite_bucket_price]
    except IndexError:
        temp = [destination_code, destination_name, subcode, vessel_id, vessel_name, cruise_id, cruise_line_name,
                itinerary_id, brochure_name, number_of_nights, sail_date, return_date, "N/A",
                "N/A", "N/A", "N/A"]
    if destination_name == 'OT' or destination_code == 'OT':
        ots.append(temp)
    else:
        to_write.append(temp)
        print(temp)


pool.map(parse, cruise_results['results'])
pool.close()
pool.join()
group = 1
for ot in ots:
    for te in to_write:
        if ot[4] == te[4] and ot[10] == te[10]:
            for te2 in to_write:
                if ot[4] == te2[4] and ot[11] == te2[11]:
                    if ot[7] == '' and te[7] == '' and te2[7] == '':
                        ot[7] = 'Group' + str(group)
                        te[7] = 'Group' + str(group)
                        te2[7] = 'Group' + str(group)
                        group += 1
                    else:
                        if ot[7] != '':
                            te[7] = ot[7]
                            te2[7] = ot[7]
                        elif te[7] != '':
                            ot[7] = te[7]
                            te2[7] = te[7]
                        elif te2[7] != '':
                            te[7] = te2[7]
                            ot[7] = te2[7]
for ot in ots:
    to_write.append(ot)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    print(userhome)
    now = datetime.datetime.now()
    path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + ' Non - Cruise only price Oceania Cruises.xlsx'
    if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 10)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 20)
    worksheet.set_column("H:H", 50)
    worksheet.set_column("I:I", 20)
    worksheet.set_column("J:J", 20)
    worksheet.set_column("K:K", 20)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 25)
    worksheet.set_column("N:N", 20)
    worksheet.set_column("O:O", 20)
    worksheet.set_column("P:P", 20)
    worksheet.write('A1', 'DestinationCode', bold)
    worksheet.write('B1', 'DestinationName', bold)
    worksheet.write('C1', 'DestinationSubcode', bold)
    worksheet.write('D1', 'VesselID', bold)
    worksheet.write('E1', 'VesselName', bold)
    worksheet.write('F1', 'CruiseID', bold)
    worksheet.write('G1', 'CruiseLineName', bold)
    worksheet.write('H1', 'ItineraryID', bold)
    worksheet.write('I1', 'BrochureName', bold)
    worksheet.write('J1', 'NumberOfNights', bold)
    worksheet.write('K1', 'SailDate', bold)
    worksheet.write('L1', 'ReturnDate', bold)
    worksheet.write('M1', 'InteriorBucketPrice', bold)
    worksheet.write('N1', 'OceanViewBucketPrice', bold)
    worksheet.write('O1', 'BalconyBucketPrice', bold)
    worksheet.write('P1', 'SuiteBucketPrice', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship_entry in data_array:
        column_count = 0
        for en in ship_entry:
            if column_count == 0:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 1:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 2:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 3:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 4:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 5:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 6:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 7:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 8:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 9:
                try:
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 10:
                try:
                    try:
                        date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, money_format)
                    except ValueError:
                        split = str(en).split('/')
                        en = split[1] + '/' + split[0] + '/' + split[2]
                        date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 11:
                try:
                    try:
                        date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, money_format)
                    except ValueError:
                        split = str(en).split('/')
                        en = split[1] + '/' + split[0] + '/' + split[2]
                        date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 12:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 13:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 14:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 15:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            column_count += 1
        row_count += 1
    workbook.close()
    pass


def write_ports_to_excell(data_array):
    userhome = os.path.expanduser('~')
    print(userhome)
    now = datetime.datetime.now()
    path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Oceania Cruises ports.xlsx'
    if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.write('A1', 'Ports', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship_entry in data_array:
        column_count = 0
        worksheet.write_string(row_count, column_count, ship_entry, centered)
        row_count += 1
    workbook.close()
    pass


write_file_to_excell(to_write)
