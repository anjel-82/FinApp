import lib.libgoogle as libgoogle

def UpdateFondos():
    #Download Fondos.xlsx from GoogleDrive
    libgoogle.download_excel_from_gdrive('Fondos')

    #Añadir datos en Diciembre en la hoja de Anjel y Maitane
    import pandas as pd
    global df_anjel, df_maitane
    df_anjel = pd.read_excel('Fondos.xlsx',sheet_name='Anjel')
    df_maitane = pd.read_excel('Fondos.xlsx',sheet_name='Maitane')

    #Se hace una copia de seguridad para subirlo como backup
    with pd.ExcelWriter('Fondos_Backup.xlsx', engine='openpyxl') as writer: 
        df_anjel.to_excel(writer, sheet_name='Anjel', index=False) 
        df_maitane.to_excel(writer, sheet_name='Maitane', index=False)

    names = df_anjel.columns.tolist()
    names.pop(0)

    combobox.delete(0, 'end')  # Limpiar el menú existente
    combobox['values'] = names  # Actualizar los valores del Combobox


def ModifyExcel():
    import pandas as pd
    Month=Monthcombobox.get()
    Year=Yearcombobox.get()
    Name = Namecombobox.get()
    MonthYear = '/'+Month+'/'+Year
    if Name=='Anjel':
        index = df_anjel[df_anjel['Date'] == MonthYear].index.tolist() 
        df_anjel.loc[index,combobox.get()] = float(EditBox.get())
    elif Name=='Maitane':
        index = df_maitane[df_maitane['Date'] == MonthYear].index.tolist() 
        df_maitane.loc[index,combobox.get()] = float(EditBox.get())
    else:
        print('No se ha encontrado el nombre')
    with pd.ExcelWriter('Fondos.xlsx', engine='openpyxl') as writer: 
        df_anjel.to_excel(writer, sheet_name='Anjel', index=False) 
        df_maitane.to_excel(writer, sheet_name='Maitane', index=False)

def UploadFondos():
    libgoogle.upload_excel_to_gdrive_and_remove('Fondos')        


#USER INTERFACE
import tkinter as tk
from tkinter import ttk

root = tk.Tk() # place a label on the root window
#root.geometry("500x500")
#root.maxsize(500, 500)  # width x height
root.wm_attributes('-transparentcolor', 'black')
root.configure(background="grey") #skyblue
#root.config(bg="#6FAFE7")  # set background color of root window
root.title("FinApp")


#Crear botones
button = tk.Button(root, text="Update Data from GDrive", command=UpdateFondos)
button.grid(row=1, column=0)

UploadButton = tk.Button(root, text="Upload to GDrive", command=UploadFondos)
UploadButton.grid(row=5, column=1)

SaveButton = tk.Button(root, text="Update Data", command=ModifyExcel)
SaveButton.grid(row=4, column=1)

# Crear un Combobox
combobox = ttk.Combobox(root)
combobox.grid(row=1, column=2)
combobox.bind("<<ComboboxSelected>>") 


#Crear zona para incluir importe de Fondo
EditBox = tk.Entry(root)
EditBox.grid(row=2, column=1)

#Añadimos dos combobox  para seleccionar el mes y el año
# Crear un Combobox
Months=['1','2','3','4','5','6','7','8','9','10','11','12']
Years=['2024','2025','2026','2027','2028','2029','2030']
Names = ['Anjel','Maitane']
Monthlabel = tk.Label(root, text="Select Month",bg="grey")
Monthlabel.grid(row=6, column=0)
Yearlabel = tk.Label(root, text="Select Year",bg="grey")
Yearlabel.grid(row=8, column=0)
Namelabel = tk.Label(root, text="Select Name",bg="grey")
Namelabel.grid(row=10, column=1)
Monthcombobox = ttk.Combobox(root)
Monthcombobox['values'] = Months  # Actualizar los valores del Combobox
Monthcombobox.grid(row=6, column=2)
Monthcombobox.bind("<<ComboboxSelected>>") 
Yearcombobox = ttk.Combobox(root)
Yearcombobox['values'] = Years  # Actualizar los valores del Combobox
Yearcombobox.grid(row=8, column=2)
Yearcombobox.bind("<<ComboboxSelected>>") 
Namecombobox = ttk.Combobox(root)
Namecombobox['values'] = Names  # Actualizar los valores del Combobox
Namecombobox.grid(row=11, column=1)
Namecombobox.bind("<<ComboboxSelected>>") 

# Añadir opciones al menú 
#menu.add_command(label="Opción 1", command=lambda: print("Opción 1 seleccionada")) 
#menu.add_command(label="Opción 2", command=lambda: print("Opción 2 seleccionada")) 
#menu.add_command(label="Opción 3", command=lambda: print("Opción 3 seleccionada")) 
#menu.insert_command(2, label='New Option')
#menu.delete(3)


root.mainloop()





#with pd.ExcelWriter('Fondos.xlsx', engine='openpyxl') as writer: 
 #   df_anjel.to_excel(writer, sheet_name='Anjel', index=False) 
  #  df_maitane.to_excel(writer, sheet_name='Maitane', index=False)

#Se sobreescribe el fichero modificado y se genera un backup con fecha actual. Todo en Google Drive
#libgoogle.upload_excel_to_gdrive_and_remove('Fondos')