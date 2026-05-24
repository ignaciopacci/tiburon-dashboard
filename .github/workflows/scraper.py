import os
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
import dropbox

# Config desde secrets
USUARIO = os.environ['AMS_USUARIO']
PASSWORD = os.environ['AMS_PASSWORD']
EMPRESA_1 = os.environ['AMS_EMPRESA_1']
EMPRESA_2 = os.environ['AMS_EMPRESA_2']
DROPBOX_TOKEN = os.environ['DROPBOX_TOKEN']

URL_LOGIN = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/login.aspx'
URL_REPORTE = 'https://apps1.mahonsistemas.com.ar/WebCorporateTiburon/InfCompCosto.aspx'

def login(page, empresa):
    print(f'Entrando como: {empresa}')
    page.goto(URL_LOGIN)
    page.wait_for_load_state('networkidle')
    page.fill('input[name*="Usuario"], input[id*="Usuario"]', USUARIO)
    page.fill('input[name*="Password"], input[id*="Password"], input[type="password"]', PASSWORD)
    # Seleccionar empresa en dropdown
    page.select_option('select[name*="Empresa"], select[id*="Empresa"]', label=empresa)
    page.click('input[value="Ingresar"], button:has-text("Ingresar")')
    page.wait_for_load_state('networkidle')
    print(f'Login OK: {empresa}')

def descargar_excel(page, empresa_nombre):
    print(f'Descargando reporte: {empresa_nombre}')
    page.goto(URL_REPORTE)
    page.wait_for_load_state('networkidle')

    # Setear fechas del mes actual
    hoy = datetime.now()
    primer_dia = hoy.replace(day=1).strftime('%d/%m/%y')
    
    # Calcular último día del mes
    if hoy.month == 12:
        ultimo_dia = hoy.replace(day=31).strftime('%d/%m/%y')
    else:
        import calendar
        ultimo = calendar.monthrange(hoy.year, hoy.month)[1]
        ultimo_dia = hoy.replace(day=ultimo).strftime('%d/%m/%y')

    # Completar fechas
    fecha_desde = page.locator('input[id*="FechaDesde"], input[name*="FechaDesde"]').first
    fecha_hasta = page.locator('input[id*="FechaHasta"], input[name*="FechaHasta"]').first
    fecha_desde.fill('')
    fecha_desde.type(primer_dia)
    fecha_hasta.fill('')
    fecha_hasta.type(ultimo_dia)

    # Click en ícono Excel (imagen verde)
    with page.expect_download() as download_info:
        page.click('img[src*="excel"], img[alt*="xls"], img[title*="xcel"], a[href*="excel"]')
    
    download = download_info.value
    filename = f'{empresa_nombre.replace(" ", "_").replace(".", "")}.xlsx'
    download.save_as(filename)
    print(f'Descargado: {filename}')
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
                archivo = descargar_excel(page, empresa)
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
