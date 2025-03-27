from pip._internal import main as pip
import os

while True:
    try:
        from playwright.sync_api import sync_playwright
        import gspread
    except ImportError as e:
        package = e.msg.split()[-1][1:-1]
        pip(['install', package])
        if package == 'playwright':
            os.system('playwright install')
    else:
        break

# def main(start: int, end: int, setup: dict):
def main():

    # worksheet: gspread.spreadsheet.Spreadsheet = setup.get('GoogleSheet')
    # sheet = worksheet.worksheet('')

    art = '62605pb01c01'
    color = '86'

    with sync_playwright() as p:
        driver = p.chromium.launch(proxy={
            'server': 'http://166.0.211.142:7576',
            'username': 'user258866',
            'password': 'pe9qf7'
        }, headless=False)
        page = driver.new_page()
        page.goto(f'https://www.bricklink.com/v2/catalog/catalogitem.page?P={art}#T=C&C={color}')

main()