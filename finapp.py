import lib.libgoogle as libgoogle

#Download Fondos.xlsx from GoogleDrive
libgoogle.download_excel_from_gdrive('Fondos')

#Añadir datos en Diciembre en la hoja de Anjel y Maitane
import pandas as pd
df_anjel = pd.read_excel('Fondos.xlsx',sheet_name='Anjel')
df_maitane = pd.read_excel('Fondos.xlsx',sheet_name='Maitane')

#Se hace una copia de seguridad para subirlo como backup
with pd.ExcelWriter('Fondos_Backup.xlsx', engine='openpyxl') as writer: 
    df_anjel.to_excel(writer, sheet_name='Anjel', index=False) 
    df_maitane.to_excel(writer, sheet_name='Maitane', index=False)

#Se modifica el fichero con lo que se necesite
index = df_anjel[df_anjel['Date'] == '/12/2024'].index.tolist() 
df_anjel.loc[index,'IE00B03HCZ61'] = 666
#df_anjel.to_excel('Fondos_copia.xlsx')
#df_anjel.to_excel('Fondos_copia.xlsx', index=False, sheet_name='Anjel', engine='openpyxl')
#df_maitane.to_excel('Fondos_copia.xlsx', index=False, sheet_name='Maitane', engine='openpyxl')
with pd.ExcelWriter('Fondos.xlsx', engine='openpyxl') as writer: 
    df_anjel.to_excel(writer, sheet_name='Anjel', index=False) 
    df_maitane.to_excel(writer, sheet_name='Maitane', index=False)

#Se sobreescribe el fichero modificado y se genera un backup con fecha actual. Todo en Google Drive
libgoogle.upload_excel_to_gdrive_and_remove('Fondos')