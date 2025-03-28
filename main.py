from pip._internal import main as pip
import os

while True:
    try:
        from playwright.sync_api import sync_playwright
        from configparser import ConfigParser
        import gspread, pathlib
    except ImportError as e:
        package = e.msg.split()[-1][1:-1]
        pip(['install', package])
        if package == 'playwright':
            os.system('playwright install')
    else:
        break

dir = pathlib.Path(__file__).parent.resolve()

config = ConfigParser()
config.read(os.path.join(dir, 'config.ini'))
sheet_url = config.get('parser', 'url')

def main(start: int, end: int, setup: dict):

    creds = setup.get('GoogleCredentials')
    google_client = gspread.authorize(creds)
    spreadsheet = google_client.open_by_url(sheet_url)
    sheet = spreadsheet.worksheet('Детали')
    
    colorsheet = spreadsheet.worksheet('colors')
    colorrange = colorsheet.range('A2:B')
    color2id = {}
    for i in range(0, len(colorrange)//2, 2):
        color2id[colorrange[i].value] = colorrange[i+1].value
    print(color2id)

    articles = sheet.range(f'C{start}:C{end}')
    colors = sheet.range(f'B{start}:B{end}')

    with sync_playwright() as p:
        driver = p.chromium.launch(proxy={
            'server': 'http://166.0.211.142:7576',
            'username': 'user258866',
            'password': 'pe9qf7'
        }, headless=False)
        page = driver.new_page()
        for idx in range(len(articles)):
            try:
                if articles[idx] == None:
                    break
            except:
                break
            page.goto(f'https://www.bricklink.com/v2/catalog/catalogitem.page?P={articles[idx].value}#T=C&C={color2id[colors[idx].value]}')

# ================ ДЛЯ ТЕСТИРОВАНИЯ ================

from google.oauth2.service_account import Credentials 
setup = {
    "GoogleCredentials": Credentials.from_service_account_file(
        os.path.join(dir, 'credentials.json'),
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
}

# ================ ДЛЯ ТЕСТИРОВАНИЯ ================

main(2, 10, setup)