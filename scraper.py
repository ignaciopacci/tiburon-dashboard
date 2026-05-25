import os
import time
import json
import re
import requests
from datetime import datetime
import calendar
import pdfplumber
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
    cookies = context.cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Referer': BASE + 'InfCompCosto.aspx',
    })
    response = session.get(url_reporte, allow_redirects=True)
    filename = f'{empresa_nombre.replace(" ", "_").replace(".", "")}.pdf'
    with open(filename, 'wb') as f:
        f.write(response.content)
    print(f'PDF descargado: {filename} ({len(response.content)} bytes)')
    return filename

def parsear_pdf(filepath):
    datos = []
    rubro_actual = None
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not any(row):
                        continue
                    # Detectar fila de rubro
                    if row[0] and 'Rubro' in str(row[0]):
                        rubro_actual = str(row[0]).strip()
                        continue
                    # Fila de artículo: tiene cantidad numérica
                    try:
                        cantidad = float(str(row[1] or '').replace('.', '').replace(',', '.'))
                        costo_unit = float(str(row[2] or '').replace('.', '').replace(',', '.'))
                        total_costo = float(str(row[3] or '').replace('.', '').replace(',', '.'))
                        total_fac = float(str(row[4] or '').replace('.', '').replace(',', '.'))
                        articulo = str(row[0] or '').strip()
                        if articulo and rubro_actual:
                            datos.append({
                                'rubro': rubro_actual,
                                'articulo': articulo,
                                'cantidad': cantidad,
                                'ultimoCosto': costo_unit,
                                'totalCosto': total_costo,
                                'totalFac': total_fac
                            })
                    except:
                        continue
    print(f'Artículos parseados: {len(datos)}')
    return datos

def generar_json(datos, empresa_nombre):
    hoy = datetime.now()
    rubros = {}
    for d in datos:
        r = 'Pinceles' if 'Pinceles' in d['rubro'] or '001' in d['rubro'] else 'Accesorios'
        if r not in rubros:
            rubros[r] = {'unidades': 0, 'totalFac': 0, 'totalCosto': 0, 'articulos': []}
        rubros[r]['unidades'] += d['cantidad']
        rubros[r]['totalFac'] += d['totalFac']
        rubros[r]['totalCosto'] += d['totalCosto']
        rubros[r]['articulos'].append(d)

    total_fac = sum(r['totalFac'] for r in rubros.values())
    total_costo = sum(r['totalCosto'] for r in rubros.values())

    resultado = {
        'empresa': empresa_nombre,
        'fechaActualizacion': hoy.strftime('%d/%m/%Y %H:%M'),
        'mes': hoy.strftime('%m/%Y'),
        'totalFac': total_fac,
        'totalCosto': total_costo,
        'ganancia': total_fac - total_costo,
        'margen': round((total_fac - total_costo) / total_fac * 100, 1) if total_fac else 0,
        'rubros': rubros
    }
    filename = f'{empresa_nombre.replace(" ", "_").replace(".", "")}.json'
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)
    print(f'JSON generado: {filename}')
    return filename

def subir_dropbox(filepath, dropbox_path):
    dbx = dropbox.Dropbox(DROPBOX_TOKEN)
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
                pdf = descargar_pdf(page, context, empresa)
                datos = parsear_pdf(pdf)
                json_file = generar_json(datos, empresa)
                nombre = empresa.replace(' ', '_').replace('.', '')
                subir_dropbox(pdf, f'/AMS_Data/{nombre}.pdf')
                subir_dropbox(json_file, f'/AMS_Data/{nombre}.json')
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
