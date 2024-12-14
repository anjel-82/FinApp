from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import gdown as gd
import datetime
import os

def download_excel_from_gdrive(googleFile):

    #Descarga de ExcelCloud.xlsx
    #Es necesario que el fichero a descargar estÃ© compartido para TODOS los que tengan link
    print('***** Download ' +googleFile+'.xlsx')
    url = 'https://docs.google.com/spreadsheets/d/1TRvRLzvtsVOseiilzcn-O9bZUex79ofh/edit?usp=sharing&ouid=102456934783095133027&rtpof=true&sd=true'
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
    file_id =   '1TRvRLzvtsVOseiilzcn-O9bZUex79ofh'
    OutFile = googleFile + '_' + timestamp + '.xlsx'
    print('Uploading..... ' + OutFile + ' file to Google Drive')
  
    #Se sobreescribe el fichero descargado con los cambios
    file = drive.CreateFile({'id': file_id}) # Modifica el fichero con ID especifico de google drive (para sobreescribir)
    file.SetContentFile(googleFile+'.xlsx')  # Cambia 'ruta_al_archivo.xlsx' por la ruta al archivo Excel en tu sistema
    file.Upload()
    print('File uploaded')

    #Se hace un backup del fichero original (xxx_Backup.xlxs) en la carpeta correspondiente
    file_backup = drive.CreateFile({'title': OutFile, 'parents': [{'id': folder_id}]}) # Cambia 'nombre_del_archivo.xlsx' por el nombre que quieras darle al archivo en Google Drive
    file_backup.SetContentFile(googleFile+'_Backup.xlsx')  # Cambia 'ruta_al_archivo.xlsx' por la ruta al archivo Excel en tu sistema
    file_backup.Upload()
    print('File Backup uploaded')

    #Eliminar fichero local.xlsx
    # Cerrar el archivo antes de eliminarlo 
    file = None
    file_backup = None
    os.remove(googleFile+'.xlsx')
    os.remove(googleFile+'_Backup.xlsx')
    print('Local files removed')
