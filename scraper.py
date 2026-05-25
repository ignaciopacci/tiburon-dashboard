import os
import time
import requests
from datetime import datetime
import calendar
from playwright.sync_api import sync_playwright
import dropbox

USUARIO = os.environ['AMS_USUARIO']
PASSWORD = os.environ['AMS_PASSWORD']
EMPRESA_1 = os.environ['AMS_EMPRESA_1']
EMPRESA_2 = os.environ['AMS_EMPRESA_2']
DROPBOX_TOKEN = os.environ['DROPBOX_TOKEN']

URL_LOGIN = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/login.aspx'
BASE = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/'

def get_url_reporte():
    hoy = datetime.now()
    primer_dia = hoy.replace(day=1).strftime('%Y%m%d')
    ultimo = calendar.monthrange(hoy.year, hoy.month)[1]
    ultimo_dia = hoy.replace(day=ultimo).strftime('%Y%m%d')
    return f'{BASE}alstinfcompcosto.aspx?{primer_dia},{ultimo_dia},PES,,A,SCR'

def login(page, empresa):
    print(f'Entrando como: {empresa}')
    page.goto(URL_LOGIN)
    page.wait_for_load_state('networkidle')
    page.wait_for_selector('#vUSUARIOCOD', timeout=15000)
    page.fill('#vUSUARIOCOD', USUARIO)
    page.press('#vUSUARIOCOD', 'Tab')
    page.wait_for_timeout(2000)
    page.evaluate(f'''
        var sel = document.querySelector('#vPERFILCGO_MPAGE');
        for (var i = 0; i < sel.options.length; i++) {{
            if (sel.options[i].text.trim() === "{empresa}") {{
                sel.selectedIndex = i;
                sel.dispatchEvent(new Event('change'));
                break;
            }}
        }}
    ''')
    page.wait_for_timeout(1000)
    page.fill('#vUSUARIOPASS', PASSWORD)
    page.wait_for_timeout(500)
    page.evaluate('''
        var btns = document.querySelectorAll("input[type=button], input[type=submit], button");
        for (var i = 0; i < btns.length; i++) {
            if (btns[i].value && btns[i].value.toLowerCase().includes("ingres")) {
                btns[i].click(); break;
            }
        }
    ''')
    page.wait_for_load_state('networkidle')
    page.wait_for_timeout(2000)
    print(f'Login OK: {empresa}')

def descargar_pdf(page, context, empresa_nombre):
    url_reporte = get_url_reporte()
    print(f'URL del reporte: {url_reporte}')

    # Usar requests directamente con las cookies de sesión
    cookies = context.cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])

    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Referer': BASE + 'InfCompCosto.aspx',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    })

    response = session.get(url_reporte, allow_redirects=True)
    print(f'Status: {response.status_code} | Tipo: {response.headers.get("Content-Type")} | Tamaño: {len(response.content)} bytes')

    filename = f'{empresa_nombre.replace(" ", "_").replace(".", "")}.pdf'
    with open(filename, 'wb') as f:
        f.write(response.content)
    print(f'Descargado: {filename}')
    return filename

def subir_dropbox(filepath, empresa_nombre):
    dbx = dropbox.Dropbox(DROPBOX_TOKEN)
    ext = filepath.split('.')[-1]
    dropbox_path = f'/AMS_Data/{empresa_nombre.replace(" ", "_").replace(".", "")}.{ext}'
    with open(filepath, 'rb') as f:
        dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)
    print(f'Subido a Dropbox: {dropbox_path}')

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        for empresa in [EMPRESA_1, EMPRESA_2]:
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            try:
                login(page, empresa)
                archivo = descargar_pdf(page, context, empresa)
                subir_dropbox(archivo, empresa)
                print(f'✓ {empresa} completado')
            except Exception as e:
                print(f'✗ Error en {empresa}: {e}')
                raise
            finally:
                context.close()
            time.sleep(3)

        browser.close()
    print('Proceso completo:', datetime.now().strftime('%d/%m/%Y %H:%M'))

if __name__ == '__main__':
    main()
