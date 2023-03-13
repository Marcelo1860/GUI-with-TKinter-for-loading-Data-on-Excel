import tkinter as tk
import pandas as pd
import io
from PIL import Image, ImageDraw, ImageFont
import textwrap


class Formulario(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Carga de datos a Excel")

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

        lista_datos = [self.txt_datos1.get("1.0", "end-1c"),self.txt_datos2.get("1.0", "end-1c"),self.txt_datos3.get("1.0", "end-1c"),self.txt_datos4.get("1.0", "end-1c"),self.txt_datos5.get("1.0", "end-1c"),self.txt_datos6.get("1.0", "end-1c"),self.txt_datos7.get("1.0", "end-1c"),self.txt_datos8.get("1.0", "end-1c")]
        print(lista_datos)

        df = pd.read_excel('datos.xlsx', sheet_name='Hoja 1')

        new_row = pd.DataFrame({'Fecha': [0], 'Nombre cliente': [0], 'Modelo PBX': [0],'S/N': [0], 'Falla acusada': [0], 'Diagnostico': [0],'Resolucion': [0], 'PV': [0] })
        df = pd.concat([df, new_row], ignore_index=True)

        for i in range(8):
            df.iloc[-1,i]=lista_datos[i]

        # Guardar el DataFrame en el archivo de destino
        writer = pd.ExcelWriter('datos.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Hoja 1', index=False)
        writer.save()
        
        # Después de cargar los datos en el DataFrame
        self.lbl_estado.config(text="Los datos se cargaron correctamente.")

        # Crear una imagen en blanco
        imagen = Image.new('RGB', (800, 1600), color='white')
    
        # Crear un objeto Draw para dibujar en la imagen
        draw = ImageDraw.Draw(imagen)
    
        font_descripcion = ImageFont.truetype('arial.ttf', 20)
        font_column = ImageFont.truetype('arial.ttf', 16)
        font_value = ImageFont.truetype('arial.ttf', 14)

        # Dibujar los textos en la imagen
        x_column, x_value = 50, 200
        y_start = 100
        y_step = 30
        prev_lines = 0

        draw.text((x_column, y_start-50), 'Ing. Simonella - Comprobante tecnico', fill='green', font=font_descripcion)
        for i, column in enumerate(df.columns):
            # Dibujar el nombre de la columna
            draw.text((x_column, y_start + y_step * 2 * i + prev_lines*y_step), column, fill='blue', font=font_column)
        
            # Obtener el valor de la última fila de la columna
            value = str(df.iloc[-1][column])
        
            # Dividir el valor en varias líneas si es necesario
            lines = textwrap.wrap(value, width=90)
            num_lines = len(lines)
        
            # Dibujar cada línea de texto
            for j, line in enumerate(lines):
                draw.text((x_column, y_start + y_step * (2*i+1) +(prev_lines+j)*y_step), line, fill='black', font=font_value)

            # Actualizar el número de líneas previas
            prev_lines = prev_lines +  num_lines -1    

        # Guardar la imagen en un archivo temporal
        imagen_path = '{}_{}_{}.pdf'.format(df.iloc[-1,0],df.iloc[-1,1],df.iloc[-1,2])
        #imagen_path = 'tata.pdf'
        imagen.save(imagen_path)

formulario = Formulario()
formulario.mainloop()

