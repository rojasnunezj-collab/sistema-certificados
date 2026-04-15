import sys
sys.path.append('.')
from src.services.google_service import obtener_servicios
from src.config.settings import ID_SHEET_REPOSITORIO
drive, sheets = obtener_servicios()
if sheets:
    r = sheets.spreadsheets().values().get(spreadsheetId=ID_SHEET_REPOSITORIO, range="'Guias_recibidas'!B:B").execute()
    vals = r.get('values', [])
    print('Total rows in B:', len(vals))
    print('First 15 rows in B:', vals[:15])
else:
    print('No sheets client')
