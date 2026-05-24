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
URL_REPORTE = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/InfCompCosto.aspx'
BASE = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/'

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

def descargar_excel(page, context, empresa_nombre):
    print(f'Descargando reporte: {empresa_nombre}')
    page.goto(URL_REPORTE)
    page.wait_for_load_state('networkidle')
    page.wait_for_selector('#vEXPORTAREXCEL', timeout=15000)

    hoy = datetime.now()
    primer_dia = hoy.replace(day=1).strftime('%d/%m/%y')
    ultimo = calendar.monthrange(hoy.year, hoy.month)[1]
    ultimo_dia = hoy.replace(day=ultimo).strftime('%d/%m/%y')

    page.evaluate(f'''
        var inputs = document.querySelectorAll("input[type=text]");
        if (inputs[0]) {{ inputs[0].value = "{primer_dia}"; inputs[0].dispatchEvent(new Event('change')); }}
        if (inputs[1]) {{ inputs[1].value = "{ultimo_dia}"; inputs[1].dispatchEvent(new Event('change')); }}
    ''')
    page.wait_for_timeout(1500)

    # Capturar URL del archivo via respuestas de red
    archivo_url = []

    def capturar_respuesta(response):
        url = response.url
        if 'PublicTempStorage' in url and ('.xlsx' in url or '.xls' in url):
            print(f'Archivo encontrado: {url}')
            archivo_url.append(url)

    page.on('response', capturar_respuesta)

    # Click en Excel
    page.click('#vEXPORTAREXCEL')
    
    # Esperar hasta 15 segundos a que aparezca la URL
    for i in range(30):
        if archivo_url:
            break
        page.wait_for_timeout(500)
    
    page.remove_listener('response', capturar_respuesta)

    if not archivo_url:
        # Intentar leer la URL desde el DOM o window
        url_desde_dom = page.evaluate('''
            () => {
                // Buscar links o iframes con el archivo
                var links = document.querySelectorAll("a[href*='PublicTempStorage'], iframe[src*='PublicTempStorage']");
                if (links.length > 0) return links[0].href || links[0].src;
                return null;
            }
        ''')
        if url_desde_dom:
            archivo_url.append(url_desde_dom)

    if not archivo_url:
        raise Exception('No se encontró la URL del archivo generado')

    url_final = archivo_url[0]
    print(f'Descargando desde: {url_final}')

    # Descargar con cookies de sesión
    cookies = context.cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])

    response = session.get(url_final)
    filename = f'{empresa_nombre.replace(" ", "_").replace(".", "")}.xlsx'
    with open(filename, 'wb') as f:
        f.write(response.content)
    print(f'Descargado: {filename} ({len(response.content)} bytes)')
    return filename

def subir_dropbox(filepath, empresa_nombre):
    dbx = dropbox.Dropbox(DROPBOX_TOKEN)
    dropbox_path = f'/AMS_Data/{empresa_nombre.replace(" ", "_").replace(".", "")}.xlsx'
    with open(filepath, 'rb') as f:
        dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)
    print(f'Subido a Dropbox: {dropbox_path}')

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        for empresa in [EMPRESA_1, EMPRESA_2]:
            context = browser.new_context()
            page = context.new_page()
            try:
                login(page, empresa)
                archivo = descargar_excel(page, context, empresa)
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
