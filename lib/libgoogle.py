def download_excel_from_gdrive():
    from pydrive2.auth import GoogleAuth
    from pydrive2.drive import GoogleDrive

    import gdown as gd

    #Descarga de ExcelCloud.xlsx
    #Es necesario que el fichero a descargar estÃ© compartido para TODOS los que tengan link
    print('***** Download Fondox.xlsx')
    url = 'https://docs.google.com/spreadsheets/d/1q1pTRNJOd4a-tdwLi1UXmuqMi15GCB3B/edit?usp=sharing&ouid=102456934783095133027&rtpof=true&sd=true'
    googleFile = 'Fondos'
    gd.download(url, googleFile+'.xlsx', quiet=False,fuzzy=True)
