from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import gdown as gd
import datetime
import os

def download_excel_from_gdrive(googleFile):

    #Descarga de ExcelCloud.xlsx
    #Es necesario que el fichero a descargar estÃ© compartido para TODOS los que tengan link
    print('***** Download ' +googleFile+'.xlsx')
    url = 'https://docs.google.com/spreadsheets/d/1q1pTRNJOd4a-tdwLi1UXmuqMi15GCB3B/edit?usp=sharing&ouid=102456934783095133027&rtpof=true&sd=true'
    gd.download(url, googleFile+'.xlsx', quiet=False,fuzzy=True)

def upload_excel_to_gdrive_and_remove(googleFile):
    #Añadir fecha al final del nombre del fichero

    # Obtener la fecha y hora actual 
    now = datetime.datetime.now()
    # Formatear la fecha y hora como una cadena 
    timestamp = now.strftime("%Y%m%d_%H%M%S")


    # Autenticación
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()  # Esto abrirá una ventana del navegador para autenticarse
    drive = GoogleDrive(gauth)

    # Crear y subir el archivo
    folder_id = '11x5pymyWpBsXSH_07gKPpxl2YjTxB0xA'
    OutFile = googleFile + '_' + timestamp + '.xlsx'
    print('Uploading..... ' + OutFile + ' file to Google Drive')
    file = drive.CreateFile({'title': OutFile, 'parents': [{'id': folder_id}]}) # Cambia 'nombre_del_archivo.xlsx' por el nombre que quieras darle al archivo en Google Drive
    file.SetContentFile(googleFile+'.xlsx')  # Cambia 'ruta_al_archivo.xlsx' por la ruta al archivo Excel en tu sistema
    file.Upload()
    print('File uploaded')

    #Eliminar fichero local.xlsx
    # Cerrar el archivo antes de eliminarlo 
    file = None
    os.remove(googleFile+'.xlsx')
