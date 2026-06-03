import os
import io
from googleapiclient.discovery import build
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials

def obtener_servicios():
    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    elif os.path.exists("secretoslocal.json"):
        creds = service_account.Credentials.from_service_account_file("secretoslocal.json", scopes=scopes)
    
    if creds:
        return build('drive', 'v3', credentials=creds), build('sheets', 'v4', credentials=creds)
    return None, None

def main():
    servicio_drive, servicio_sheets = obtener_servicios()
    if not servicio_drive:
        print("No creds")
        return
        
    # Get recent pending files from sheet
    ID_SHEET_GUIAS = "14As5bCpZi56V5Nq1DRs0xl6R1LuOXLvRRoV26nI50NU"
    r = servicio_sheets.spreadsheets().values().get(spreadsheetId=ID_SHEET_GUIAS, range="'Guias_recibidas'!A2:H").execute()
    v = r.get('values', [])
    
    import sys
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from src.services.google_service import descargar_guias_drive
    print("Recent items in sheets:")
    for fila in reversed(v[-10:]):
        if len(fila) >= 6:
            f_archivo = str(fila[5]).strip()
            print(f"File name or URL in sheet: {f_archivo}")
            archivos = descargar_guias_drive(servicio_drive, [f_archivo])
            print(f"Downloaded files: {archivos}")
            break

if __name__ == '__main__':
    main()
