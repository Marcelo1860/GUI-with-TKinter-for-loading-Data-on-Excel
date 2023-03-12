import tkinter as tk
import pandas as pd
import io


class Formulario(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Carga de datos a Excel")
        
        # Label para indicar el archivo de destino
        self.lbl_destino = tk.Label(self, text="Archivo de destino:")
        self.lbl_destino.pack()
        
        # Entry para ingresar el nombre del archivo de destino
        self.ent_destino = tk.Entry(self)
        self.ent_destino.pack()
        
        # Label para indicar la hoja de destino
        self.lbl_hoja = tk.Label(self, text="Hoja de destino:")
        self.lbl_hoja.pack()
        
        # Entry para ingresar el nombre de la hoja de destino
        self.ent_hoja = tk.Entry(self)
        self.ent_hoja.pack()
        
        # Label para indicar los datos a cargar
        self.lbl_datos = tk.Label(self, text="Datos a cargar:")
        self.lbl_datos.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos = tk.Text(self)
        self.txt_datos.pack()
        
        # Botón para cargar los datos a Excel
        self.btn_cargar = tk.Button(self, text="Cargar", command=self.cargar_datos)
        self.btn_cargar.pack()

        # Muestra estado en un label
        self.lbl_estado = tk.Label(self, text="")
        self.lbl_estado.pack()
    
    def cargar_datos(self):
        # Obtener el nombre del archivo de destino y la hoja de destino
        archivo_destino = self.ent_destino.get()
        hoja_destino = self.ent_hoja.get()
        
        # Obtener los datos ingresados en el campo de texto
        datos = self.txt_datos.get("1.0", "end-1c")
        
        # Convertir los datos en un DataFrame de pandas
        df = pd.read_csv(io.StringIO(datos), sep='\t')
        
        # Guardar el DataFrame en el archivo de destino
        writer = pd.ExcelWriter(archivo_destino, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=hoja_destino, index=False)
        writer.save()
        
        # Después de cargar los datos en el DataFrame
        self.lbl_estado.config(text="Los datos se cargaron correctamente.")

formulario = Formulario()
formulario.mainloop()