import os
import time
import json
import re
import requests
from datetime import datetime, timedelta, timezone
import calendar
import pdfplumber
from io import BytesIO
from playwright.sync_api import sync_playwright
import dropbox
import base64
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import xlrd

ARG = timezone(timedelta(hours=-3))

USUARIO = os.environ['AMS_USUARIO']
PASSWORD = os.environ['AMS_PASSWORD']
EMPRESA_1 = os.environ['AMS_EMPRESA_1']
EMPRESA_2 = os.environ['AMS_EMPRESA_2']
APP_KEY = os.environ['DROPBOX_APP_KEY']
APP_SECRET = os.environ['DROPBOX_APP_SECRET']
REFRESH_TOKEN = os.environ['DROPBOX_REFRESH_TOKEN']
GH_TOKEN = os.environ['GH_TOKEN_PUSH']
GH_REPO = 'ignaciopacci/tiburon-dashboard'
GOOGLE_CREDENTIALS = os.environ['GOOGLE_CREDENTIALS']
GANANCIAS_FILE_ID = '1b2rxkDVO9ujMurm19gqOd9C1myfDt_b_'

URL_LOGIN = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/login.aspx'
BASE = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/'

def get_hora_arg():
    try:
        r = requests.get('https://worldtimeapi.org/api/timezone/America/Argentina/Buenos_Aires', timeout=5)
        dt_str = r.json()['datetime'][:19]
        return datetime.fromisoformat(dt_str)
    except:
        return datetime.now(ARG).replace(tzinfo=None)

def get_url_reporte():
    hoy = get_hora_arg()
    primer_dia = hoy.replace(day=1).strftime('%Y%m%d')
    ultimo = calendar.monthrange(hoy.year, hoy.month)[1]
    ultimo_dia = hoy.replace(day=ultimo).strftime('%Y%m%d')
    return f'{BASE}alstinfcompcosto.aspx?{primer_dia},{ultimo_dia},PES,,A,SCR'

def login(page, empresa, valor):
    print(f'Entrando como: [{empresa}] (value={valor})')
    page.goto(URL_LOGIN)
    page.wait_for_load_state('networkidle')
    page.wait_for_selector('#vUSUARIOCOD', timeout=15000)
    page.fill('#vUSUARIOCOD', USUARIO)
    page.press('#vUSUARIOCOD', 'Tab')
    page.wait_for_timeout(2000)
    page.evaluate(f'''
        var sel = document.querySelector('#vEMPRESACGO');
        sel.value = '{valor}';
        sel.dispatchEvent(new Event('change', {{bubbles: true}}));
    ''')
    page.wait_for_timeout(1000)
    seleccionado = page.evaluate("document.querySelector('#vEMPRESACGO').value")
    print(f'Empresa seleccionada value: {seleccionado}')
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
    print(f'Login OK: [{empresa}]')

def descargar_pdf(page, context):
    url_reporte = get_url_reporte()
    print(f'Descargando PDF: {url_reporte}')
    cookies = context.cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Referer': BASE + 'InfCompCosto.aspx',
    })
    response = session.get(url_reporte, allow_redirects=True)
    print(f'PDF: {len(response.content)} bytes')
    return response.content

def parsear_pdf(pdf_bytes):
    datos = []
    rubro_actual = None
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue
            for linea in texto.split('\n'):
                linea = linea.strip()
                if not linea:
                    continue
                if 'Rubro Pinceles' in linea or 'Pinceles(001)' in linea:
                    rubro_actual = 'Pinceles'
                    continue
                if 'Rubro Accesorios' in linea or 'Accesorios(002)' in linea:
                    rubro_actual = 'Accesorios'
                    continue
                if 'Rubro Rodillos' in linea or 'Rodillos(003)' in linea:
                    rubro_actual = 'Rodillos'
                    continue
                if not rubro_actual:
                    continue
                match = re.search(
                    r'^(.+?\([A-Z0-9\-]+\))\s+([\d\-]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s*$',
                    linea
                )
                if match:
                    try:
                        articulo = match.group(1).strip()
                        cantidad = float(match.group(2).replace('.', '').replace(',', '.'))
                        costo_unit = float(match.group(3).replace('.', '').replace(',', '.'))
                        total_costo = float(match.group(4).replace('.', '').replace(',', '.'))
                        total_fac = float(match.group(5).replace('.', '').replace(',', '.'))
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
    hoy = get_hora_arg()
    rubros = {}
    for d in datos:
        r = d['rubro']
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
        'fechaActualizacion': hoy.strftime('%d/%m/%Y %H:%M'),
        'mes': hoy.strftime('%m/%Y'),
        'anio': hoy.year,
        'nroMes': hoy.month,
        'totalFac': total_fac,
        'totalCosto': total_costo,
        'ganancia': total_fac - total_costo,
        'margen': round((total_fac - total_costo) / total_fac * 100, 1) if total_fac else 0,
        'rubros': rubros
    }

def subir_github(contenido_str, path):
    url = f'https://api.github.com/repos/{GH_REPO}/contents/{path}'
    headers = {
        'Authorization': f'token {GH_TOKEN}',
        'Accept': 'application/vnd.github.v3+json'
    }
    r = requests.get(url, headers=headers)
    sha = r.json().get('sha') if r.status_code == 200 else None
    contenido_b64 = base64.b64encode(contenido_str.encode('utf-8')).decode('utf-8')
    data = {
        'message': f'Actualización {datetime.now(ARG).strftime("%d/%m/%Y %H:%M")}',
        'content': contenido_b64
    }
    if sha:
        data['sha'] = sha
    r = requests.put(url, headers=headers, json=data)
    if r.status_code in [200, 201]:
        print(f'✓ GitHub: {path}')
    else:
        raise Exception(f'Error GitHub {r.status_code}: {path}')

def leer_ganancias_drive():
    print('Leyendo Excel de ganancias desde Google Drive...')
    creds_dict = json.loads(GOOGLE_CREDENTIALS)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    service = build('drive', 'v3', credentials=creds)
    request = service.files().get_media(fileId=GANANCIAS_FILE_ID)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)

    wb = xlrd.open_workbook(file_contents=fh.read())
    print(f'Hojas: {wb.sheet_names()}')

    resultado = {'anios': {}, 'mensual': []}

    # Hoja GANANCIA — resumen anual
    # Columnas: 0=Año, 1=GananciaBruta, 2=Gasto, 3=GananciaLimpia, 4=Pct, 5=Pct, 6=Ventas
    if 'GANANCIA' in wb.sheet_names():
        ws = wb.sheet_by_name('GANANCIA')
        for i in range(ws.nrows):
            row = ws.row_values(i)
            if row[0] and str(row[0]).strip().replace('.0', '').isdigit():
                try:
                    anio = int(float(row[0]))
                    resultado['anios'][anio] = {
                        'anio': anio,
                        'ventas': float(row[6]) if row[6] else 0,
                        'costo': float(row[2]) if row[2] else 0,
                        'gananciaBruta': float(row[1]) if row[1] else 0,
                        'gananciaLimpia': float(row[3]) if row[3] else 0,
                    }
                except:
                    continue

    # Hoja VENTAS — datos mensuales (estructura vertical por año)
    # Columnas: 0=Mes, 1=Ventas, 2=Costo, 3=GananciaBruta, 4=Gastos, 5=GananciaLimpia
    if 'VENTAS' in wb.sheet_names():
        ws = wb.sheet_by_name('VENTAS')
        meses_map = {
            'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4,
            'MAYO': 5, 'JUNIO': 6, 'JULIO': 7, 'AGOSTO': 8,
            'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12
        }
        meses_nombre = {
            1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
            7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'
        }
        anio_actual = None
        for i in range(ws.nrows):
            row = ws.row_values(i)
            if not row[0]:
                continue
            celda = str(row[0]).strip().upper()
            # Detectar fila de año: "AÑO 2021", "AÑO 2022", etc.
            if 'AÑO' in celda or 'ANO' in celda:
                partes = celda.split()
                for p in partes:
                    p_clean = p.replace('.0', '')
                    if p_clean.isdigit() and len(p_clean) == 4:
                        anio_actual = int(p_clean)
                        break
                continue
            # Detectar fila de mes
            if anio_actual and celda in meses_map:
                nro_mes = meses_map[celda]
                try:
                    ventas = float(row[1]) if row[1] else 0
                    costo = float(row[2]) if row[2] else 0
                    ganancia_bruta = float(row[3]) if row[3] else 0
                    ganancia_limpia = float(row[5]) if len(row) > 5 and row[5] else 0
                    if ventas:
                        resultado['mensual'].append({
                            'anio': anio_actual,
                            'mes': nro_mes,
                            'mesNombre': meses_nombre[nro_mes],
                            'ventas': ventas,
                            'costo': costo,
                            'gananciaBruta': ganancia_bruta,
                            'gananciaLimpia': ganancia_limpia,
                        })
                except:
                    continue

    print(f'✓ Ganancias: {len(resultado["anios"])} años, {len(resultado["mensual"])} registros mensuales')
    return resultado

def main():
    hoy = get_hora_arg()
    mes_key = hoy.strftime('%Y-%m')
    print(f'Hora Argentina: {hoy.strftime("%d/%m/%Y %H:%M")}')

    empresas = [
        (EMPRESA_1, '1'),
        (EMPRESA_2, '2'),
    ]

    resultados = {}

    for empresa, valor in empresas:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            try:
                login(page, empresa, valor)
                pdf_bytes = descargar_pdf(page, context)
                datos = parsear_pdf(pdf_bytes)
                resultado = generar_json(datos, empresa)
                nombre = empresa.replace(' ', '_').replace('.', '')
                resultados[nombre] = resultado
                print(f'✓ {empresa} — {len(datos)} artículos')
            except Exception as e:
                print(f'✗ Error: {e}')
                raise
            finally:
                context.close()
                browser.close()
        time.sleep(3)

    # Subir JSONs de ventas a GitHub
    for nombre, resultado in resultados.items():
        contenido = json.dumps(resultado, ensure_ascii=False, indent=2)
        subir_github(contenido, f'data/{nombre}.json')
        subir_github(contenido, f'data/historico/{mes_key}_{nombre}.json')

    # Leer y subir ganancias históricas desde Drive
    try:
        ganancias = leer_ganancias_drive()
        subir_github(
            json.dumps(ganancias, ensure_ascii=False, indent=2),
            'data/ganancias_historicas.json'
        )
    except Exception as e:
        print(f'✗ Error leyendo ganancias Drive: {e}')

    print('Completo:', hoy.strftime('%d/%m/%Y %H:%M'))

if __name__ == '__main__':
    main()
