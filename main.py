import os

while True:
    try:
        from playwright.sync_api import sync_playwright, expect
        from configparser import ConfigParser
        import gspread, pathlib, requests
    except ImportError as e:
        package = e.msg.split()[-1][1:-1]
        os.system(f'python -m pip install {package}')
    else:
        break

dir = pathlib.Path(__file__).parent.resolve()

config = ConfigParser()
config.read(os.path.join(dir, 'config.ini'))
sheet_url = config.get('parser', 'url')

def main(start: int, end: int, setup: dict):
    
    print('Подготавливаем всё для работы')
    if start < 3: start = 3

    # ==> ПОДКЛЮЧЕНИЕ ГУГЛ-АККАУНТА
    creds = setup.get('GoogleCredentials')
    google_client = gspread.authorize(creds)
    spreadsheet = google_client.open_by_url(sheet_url)
    sheet = spreadsheet.worksheet('Детали')
    
    # ==> СОЗДАНИЕ ТАБЛИЦЫ ПЕРЕВОДА ЦВЕТА В ID
    colorsheet = spreadsheet.worksheet('colors')
    colorrange = colorsheet.range('A2:B')
    color2id = {}
    for i in range(0, len(colorrange)//2, 2):
        color2id[colorrange[i].value] = colorrange[i+1].value

    # ==> ПОЛУЧЕНИЕ ДАННЫХ С ТАБЛИЦЫ
    articles = sheet.range(f'C{start}:C{end}')
    colors = sheet.range(f'B{start}:B{end}')
    qty_res = []
    price_res = []

    # ==> ПОЛУЧЕНИЕ КУРСА ДОЛЛАРА
    rub = requests.get('https://www.cbr-xml-daily.ru/daily_json.js').json()['Valute']['USD']['Value']

    print('Начинаем работу с браузером')

    # ==> РАБОТА С БРАУЗЕРОМ
    for idx in range(len(articles)):
        try: 
            if articles[idx].value == '' and colors[idx].value == '': break
        except: 
            break
        with sync_playwright() as p:
            driver = p.chromium.launch(proxy={
                'server': 'http://166.0.211.142:7576',
                'username': 'user258866',
                'password': 'pe9qf7'
            })
            page = driver.new_page()
            try:
                print(f'Обработка артикула {articles[idx].value} {colors[idx].value}')
                page.goto(f'https://www.bricklink.com/v2/catalog/catalogitem.page?P={articles[idx].value}#T=P&C={color2id[colors[idx].value]}')
                page.wait_for_selector('table.pcipgMainTable')
                table = page.query_selector('#_idPGContents > table > tbody > tr:nth-child(3) > td:nth-child(4)')
                rows = table.query_selector_all('tr')
                qty_val = int(rows[1].query_selector_all('td')[-1].text_content())
                prc_val = round(float(rows[4].query_selector_all('td')[-1].text_content()[4:]) * rub)
            except Exception as e:
                print(f'Ошибка при работе с артикулом {articles[idx].value} (https://www.bricklink.com/v2/catalog/catalogitem.page?P={articles[idx].value}#T=P&C={color2id[colors[idx].value]})\n{str(e)}')
                qty_res.append([0])
                qty_res.append([0])
            else:
                qty_res.append([qty_val])
                price_res.append([prc_val])
            finally: continue
    print(f'Загружаем информацию на таблицу')
    sheet.update(qty_res, f'H{start}:H{end}')
    sheet.update(price_res, f'G{start}:G{end}')
    print(f'Программа завершила выполнение')
