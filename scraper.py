import os
import time
import json
import requests
from datetime import datetime
import calendar
import openpyxl
from io import BytesIO
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

def get_fechas():
    hoy = datetime.now()
    primer_dia = hoy.replace(day=1).strftime('%d/%m/%y')
    ultimo = calendar.monthrange(hoy.year, hoy.month)[1]
    ultimo_dia = hoy.replace(day=ultimo).strftime('%d/%m/%y')
    return primer_dia, ultimo_dia

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
    print(f'Navegando al reporte...')
    page.goto(URL_REPORTE)
    page.wait_for_load_state('networkidle')
    page.wait_for_selector('#vEXPORTAREXCEL', timeout=15000)

    primer_dia, ultimo_dia = get_fechas()
    page.evaluate(f'''
        var inputs = document.querySelectorAll("input[type=text]");
        if (inputs[0]) {{ inputs[0].value = "{primer_dia}"; inputs[0].dispatchEvent(new Event('change')); }}
        if (inputs[1]) {{ inputs[1].value = "{ultimo_dia}"; inputs[1].dispatchEvent(new Event('change')); }}
    ''')
    page.wait_for_timeout(1500)

    # Interceptar URL del Excel via requests de red
    excel_url = []

    def on_request(request):
        url = request.url
        if 'PublicTempStorage' in url and '.xlsx' in url.lower():
            print(f'Excel URL interceptada (request): {url}')
            excel_url.append(url)

    def on_response(response):
        url = response.url
        if 'PublicTempStorage' in url and '.xlsx' in url.lower():
            print(f'Excel URL interceptada (response): {url}')
            excel_url.append(url)

    page.on('request', on_request)
    page.on('response', on_response)

    # Click en Excel
    print('Clickeando Excel...')
    page.click('#vEXPORTAREXCEL')
    
    # Esperar hasta 20 segundos
    for i in range(40):
        if excel_url:
            break
        page.wait_for_timeout(500)
        if i % 10 == 0:
            print(f'Esperando URL... {i/2}s')

    page.remove_listener('request', on_request)
    page.remove_listener('response', on_response)

    if not excel_url:
        # Buscar en el DOM
        url_dom = page.evaluate('''
            () => {
                var links = document.querySelectorAll("a, iframe");
                for (var l of links) {
                    var href = l.href || l.src || '';
                    if (href.includes('PublicTempStorage')) return href;
                }
                return null;
            }
        ''')
        if url_dom:
            excel_url.append(url_dom)

    if not excel_url:
        raise Exception('No se pudo interceptar la URL del Excel')

    url_final = excel_url[0]
    print(f'Descargando Excel desde: {url_final}')

    cookies = context.cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    session.headers.update({'Referer': BASE})

    response = session.get(url_final)
    print(f'Excel descargado: {len(response.content)} bytes')
    return response.content

def parsear_excel(contenido):
    wb = openpyxl.load_workbook(BytesIO(contenido))
    ws = wb.active
    datos = []
    rubro_actual = None
    header_row = None

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if row[0] == 'Rubro':
            header_row = i
            continue
        if header_row is None:
            continue
        if not row[0]:
            continue
        # Detectar fila de rubro
        if isinstance(row[0], str) and ('Pinceles' in row[0] or 'Accesorios' in row[0] or 'Rodillos' in row[0]):
            rubro_actual = row[0]
            continue
        # Fila de artículo
        try:
            cantidad = float(row[2] or 0)
            costo_unit = float(row[3] or 0)
            total_costo = float(row[4] or 0)
            total_fac = float(row[5] or 0)
            articulo = str(row[1] or '').strip()
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

    return {
        'empresa': empresa_nombre,
        'fechaAc
