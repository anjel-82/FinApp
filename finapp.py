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

# Función para actualizar el Label con el nombre seleccionado
def actualizar_label(event):
    nombre_seleccionado = combobox.get()
    label.config(text=nombre_seleccionado)

def ModifyExcel():
    import pandas as pd
    index = df_anjel[df_anjel['Date'] == '/12/2024'].index.tolist() 
    df_anjel.loc[index,'IE00B03HCZ61'] = float(EditBox.get())
    with pd.ExcelWriter('Fondos.xlsx', engine='openpyxl') as writer: 
        df_anjel.to_excel(writer, sheet_name='Anjel', index=False) 
        df_maitane.to_excel(writer, sheet_name='Maitane', index=False)


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

# Crear un Label para mostrar el nombre seleccionado
label = tk.Label(root, text="", bg="sky blue", width=20, height=2, font=("Arial", 14))
label.pack(pady=10)

#Crear botones
button = tk.Button(root, text="Update Data from GDrive", command=UpdateFondos)
button.pack(side="left",padx=2)

SaveButton = tk.Button(root, text="Update Data To GDrive", command=ModifyExcel)
SaveButton.pack(side="right",padx=2)

# Crear un Combobox
combobox = ttk.Combobox(root)
combobox.pack(pady=10)
combobox.bind("<<ComboboxSelected>>", actualizar_label) 


#Crear zona para incluir importe de Fondo
EditBox = tk.Entry(root)
EditBox.pack()



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