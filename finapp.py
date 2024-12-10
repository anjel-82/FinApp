import lib.libgoogle as libgoogle

#Download Fondos.xlsx from GoogleDrive
libgoogle.download_excel_from_gdrive('Fondos')

#Añadir datos en Diciembre en la hoja de Anjel y Maitane
import pandas as pd
df_anjel = pd.read_excel('Fondos.xlsx',sheet_name='Anjel')
df_maitane = pd.read_excel('Fondos.xlsx',sheet_name='Maitane')

index = df_anjel[df_anjel['Date'] == '/12/2024'].index.tolist() 
df_anjel.loc[13,'IE00B03HCZ61'] = 666
df_anjel.to_excel('Fondos_copia.xlsx')

#libgoogle.upload_excel_to_gdrive_and_remove('Fondos')