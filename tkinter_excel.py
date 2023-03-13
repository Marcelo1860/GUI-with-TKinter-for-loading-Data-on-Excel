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
        self.lbl_datos1 = tk.Label(self, text="Fecha de ingreso")
        self.lbl_datos1.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos1 = tk.Text(self, height=1, width=100)
        self.txt_datos1.pack()

        # Label para indicar los datos a cargar
        self.lbl_datos2 = tk.Label(self, text="Nombre Cliente")
        self.lbl_datos2.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos2 = tk.Text(self, height=1, width=100)
        self.txt_datos2.pack()

        # Label para indicar los datos a cargar
        self.lbl_datos3 = tk.Label(self, text="Modelo PBX")
        self.lbl_datos3.pack()

        # Text para ingresar los datos a cargar
        self.txt_datos3 = tk.Text(self, height=1, width=100)
        self.txt_datos3.pack()

        # Label para indicar los datos a cargar
        self.lbl_datos4 = tk.Label(self, text="S/N")
        self.lbl_datos4.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos4 = tk.Text(self, height=1, width=100)
        self.txt_datos4.pack()

        # Label para indicar los datos a cargar
        self.lbl_datos5 = tk.Label(self, text="Falla acusada")
        self.lbl_datos5.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos5 = tk.Text(self, height=3, width=100)
        self.txt_datos5.pack()

        # Label para indicar los datos a cargar
        self.lbl_datos6 = tk.Label(self, text="Diagnostico")
        self.lbl_datos6.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos6 = tk.Text(self, height=3, width=100)
        self.txt_datos6.pack()

        # Label para indicar los datos a cargar
        self.lbl_datos7 = tk.Label(self, text="Resolucion")
        self.lbl_datos7.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos7 = tk.Text(self, height=3, width=100)
        self.txt_datos7.pack()

        # Label para indicar los datos a cargar
        self.lbl_datos8 = tk.Label(self, text="PV")
        self.lbl_datos8.pack()
        
        # Text para ingresar los datos a cargar
        self.txt_datos8 = tk.Text(self, height=1, width=100)
        self.txt_datos8.pack()

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
        lista_datos = [self.txt_datos1.get("1.0", "end-1c"),self.txt_datos2.get("1.0", "end-1c"),self.txt_datos3.get("1.0", "end-1c"),self.txt_datos4.get("1.0", "end-1c"),self.txt_datos5.get("1.0", "end-1c"),self.txt_datos6.get("1.0", "end-1c"),self.txt_datos7.get("1.0", "end-1c"),self.txt_datos8.get("1.0", "end-1c")]
        print(lista_datos)

        df = pd.read_excel(archivo_destino, hoja_destino)

        new_row = pd.DataFrame({'Fecha': [0], 'Nombre cliente': [0], 'Modelo PBX': [0],'S/N': [0], 'Falla acusada': [0], 'Diagnostico': [0],'Resolucion': [0], 'PV': [0] })
        df = pd.concat([df, new_row], ignore_index=True)

        for i in range(8):
            df.iloc[-1,i]=lista_datos[i]

        # Guardar el DataFrame en el archivo de destino
        writer = pd.ExcelWriter(archivo_destino, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=hoja_destino, index=False)
        writer.save()
        
        # Después de cargar los datos en el DataFrame
        self.lbl_estado.config(text="Los datos se cargaron correctamente.")

formulario = Formulario()
formulario.mainloop()