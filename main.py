import os

while True:
    try:
        from playwright.async_api import async_playwright, Browser
        import gspread, pathlib, requests, asyncio
        from async_lru import alru_cache
    except ImportError as e:
        package = e.msg.split()[-1][1:-1]
        os.system(f'python -m pip install {package}')
    else:
        break

dir = pathlib.Path(__file__).parent.resolve()
rub = requests.get('https://www.cbr-xml-daily.ru/daily_json.js').json()['Valute']['USD']['Value']
color2id = {}


async def main(start: int, end: int, setup: dict):
    
    print('Подготавливаем всё для работы')
    global color2id
    if start < 3: start = 3

    # ==> ПОДКЛЮЧЕНИЕ ГУГЛ-АККАУНТА
    spreadsheet: gspread.spreadsheet.Spreadsheet = setup.get('DetailsSheet')
    sheet = spreadsheet.worksheet('Детали')
    
    # ==> СОЗДАНИЕ ТАБЛИЦЫ ПЕРЕВОДА ЦВЕТА В ID
    colorsheet = spreadsheet.worksheet('colors')
    colorrange = colorsheet.range('A2:B')
    for i in range(0, len(colorrange)//2, 2):
        color2id[colorrange[i].value] = colorrange[i+1].value

    # ==> ПОЛУЧЕНИЕ ДАННЫХ С ТАБЛИЦЫ
    articles = sheet.range(f'C{start}:C{end}')
    colors = sheet.range(f'B{start}:B{end}')
    qty_res = []
    price_res = []
    name_res = []

    print('Начинаем работу с браузером')

    # ==> РАБОТА С БРАУЗЕРОМ
    semaphore = asyncio.Semaphore(4)
    async with async_playwright() as p: 
        browser = await p.chromium.launch(proxy={
            'server': 'http://166.0.211.142:7576',
            'username': 'user258866',
            'password': 'pe9qf7'
        })
        tasks = [parse_item(semaphore, browser, articles[idx].value, colors[idx].value) for idx in range(len(articles))]
        results = await asyncio.gather(*tasks)
    for res in results:
        qty_res.append([res[0]])
        price_res.append([res[1]])
        name_res.append([res[2]])
    print(f'Загружаем информацию на таблицу')
    sheet.update(qty_res, f'H{start}:H{len(qty_res)+start}')
    sheet.update(price_res, f'G{start}:G{len(price_res)+start}')
    sheet.update(name_res, f'A{start}:A{len(name_res)+start}')
    print(f'Программа завершила выполнение')

@alru_cache(None)
async def parse_item(semaphore: asyncio.Semaphore, driver: Browser, art: str, color_name: str):
    global rub, color2id
    if (not art) or (not color_name): return (None, None, None)
    try:
        color = color2id[color_name]
    except:
        return (None, None, None)
    async with semaphore:
        page = await driver.new_page()
        try:
            print(f'Обработка артикула {art} {color}')
            await page.goto(f'https://www.bricklink.com/v2/catalog/catalogitem.page?P={art}#T=P&C={color}')
            await page.wait_for_selector('table.pcipgMainTable')
            table = await page.query_selector('#_idPGContents > table > tbody > tr:nth-child(3) > td:nth-child(4)')
            rows = await table.query_selector_all('tr')
            # Get Qty Value
            rows_1_all_td = await rows[1].query_selector_all('td')
            qty_val = int(await rows_1_all_td[-1].text_content())
            # Get Price Value
            rows_4_all_td = await rows[4].query_selector_all('td')
            rows_4_text_content = await rows_4_all_td[-1].text_content()
            prc_val = round(float(rows_4_text_content[4:]) * rub)
            # Get Name Value
            item_name_title_div = await page.query_selector('#item-name-title')
            name_val = await item_name_title_div.text_content()
        except Exception as e:
            await page.close()
            return (None, None, None)
        else:
            await page.close()
            return (qty_val, prc_val, name_val)

if __name__ == '__main__':
    from Setup.setup import setup
    asyncio.run(main(3, 10_000, setup))