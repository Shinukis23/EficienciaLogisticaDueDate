# Script para llenado de formatos PDF 3520 y NTHSA para exportacion de mercancia
# La informacion necesaria para llenar los formatos se obtienen al leer la factura en Excel
# Julio 25/ 2023

import pandas as pd
from datetime import datetime
#from PyPDF4 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject
from PyPDF4.generic import NameObject
from PyPDF4.generic import TextStringObject
from PyPDF4.pdf import PdfFileReader
from PyPDF4.pdf import PdfFileWriter
from fdfgen import forge_fdf 
from gooey import Gooey, GooeyParser
import subprocess
import json
import random
import sys
import os



# Ruta del archivo Excel
#archivo_excel = "FACT-1115 25JULIO23.xlsx"
infile = "O_NTHSA_unlocked.pdf"
infile2 = "O_form3520-1-unlocked.pdf"
work_path = os.getcwd()

@Gooey(program_name="Programa generador de Formatos PDf para exportacion Ver.2")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    #print(args_file)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            #print("ffff",data_file)
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Genera formatos PDF')
    parser.add_argument('Factura_partes',
                        action='store',
                        default=stored_args.get('cust_file'),
                        widget='FileChooser',
                        help='Ej. FACT-1115 diaMesAño.xlsx')
    #parser.add_argument('Archivo_Rutas',
    #                    action='store',
    #                    default=stored_args.get('cust_file'),
    #                    widget='FileChooser',
    #                    help='Ej. Rutas pendientes.xls')
    parser.add_argument('Directorio_de_trabajo',
                        action='store',
                        default=stored_args.get('data_directory'),
                        widget='DirChooser',
                        help="Directorio de salida ")
    
    #parser.add_argument('Fecha', help='Seleccione Fecha del Reporte',
    #                    default=stored_args.get('Fecha'),
    #                    widget='DateChooser')
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args


def unir_archivos_pdf(file,folder_path, output_filename,valor_factura):
    pdf_writer = PdfFileWriter()

    for filename in os.listdir(work_path):
        if filename.startswith(file) and filename.endswith('.pdf'):
            file_path = os.path.join(work_path, filename)
            pdf_reader = PdfFileReader(file_path)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                pdf_writer.addPage(page)

            # Copiar los campos de formulario del archivo original al nuevo archivo
            #if pdf_reader.getFields():
            #    pdf_writer.updatePageFormFieldValues(pdf_writer.getPageNumber(page), pdf_reader.getFields())

    output_path = os.path.join(folder_path, output_filename+"_"+str(valor_factura)+".pdf")
    with open(output_path, 'wb') as output_file:
        pdf_writer.write(output_file)
    for filename in os.listdir(work_path):
        if filename.startswith(file) and filename.endswith('.pdf'):
            os.remove(filename)
    print(f"Archivos PDF y formularios unidos exitosamente en '{output_filename}'.")

#def Principal(Directorio_de_trabajo,ReporteProduccionDB,Rutas_pendientes):
def Principal(Factura_partes,folder_path):
    # Leer el archivo Excel
    dini = pd.read_excel(Factura_partes)
    valor_celda = dini.iloc[2, 6]
    valor_factura = dini.iloc[1, 6]
    valor_celda = datetime.strftime(valor_celda, "%m/%d/%Y")
    df = dini[dini.iloc[:, 11].notnull()]
    #df = dini.notnull(dini.iloc[:, 11])

    df_filtrado = df[df.iloc[:, 6] == "ENGINE ASSY"]
    print(len(df_filtrado))
    print(len(df))
    print(df)
    print(folder_path)
    # Leer todas las filas después de la fila 14 hasta que no exista otra fila con datos
    datos_finales = []
    j = 0
    #writer2.add_page(page)
    for i in range(1, len(df)):
        fila = df.iloc[i]
        if pd.isnull(fila[0]):
            break
    #for i in range(0, len(df_filtrado)):
    #    fila = df_filtrado.iloc[i]
    #    if pd.isnull(fila[0]):
    #        break
        print(j)
        dato_m = fila[12].split('/')[0].strip()  # Separar el dato de la columna 'M' antes del "/"
        dato_n = fila[13].split('/')[0].strip()  # Separar el dato de la columna 'N' después del "/"
        dato_o = fila[12].split('/')[1].strip()  # Separar el dato de la columna 'M' antes del "/"
        dato_p = fila[13].split('/')[1].strip()  # Separar el dato de la columna 'N' después del "/"
        fila[8] = dato_m  # Actualizar el valor de la columna 'M'
        fila[12] = dato_n  # Actualizar el valor de la columna 'N'
        fila[9] = dato_o  # Actualizar el valor de la columna 'M'
        fila[13] = dato_p  # Actualizar el valor de la columna 'N'
        V1 = fila[13] + "/" + fila[12]  # Concatenar los datos de las columnas 13 y 12 en V1
        V2 = fila[9] + "/ENGINE ASSEMBLY"  # Concatenar los datos de las columnas 9 y 6 en V2
        V3 = fila[8]  # Asignar el valor de la columna 8 a V3
        if len(str(fila[11])) > 17 and  (fila[6] == "ENGINE" or fila[6] == "TRANSMISSION"):
            #V6 = fila[11]
            V6 = fila[15] # Asignar el valor de la columna 15 a V6
        else:     
            #V6 = fila[15] # Asignar el valor de la columna 15 a V6
            V6 = fila[11]
        V7 = fila[6]
        V8 = fila[12]
        V9 = fila[9]
        #datos_finales.append([V1, V2, V3, V6, valor_celda])


        fields = [
                ('MAKE OF VEHICLE', V3),
                ('VEHICLE IDENTIFICATION NUMBER VIN', V6),
                ('MODEL', V9),
                ('YEAR', V8),
                ('REGISTERED IMPORTER NAME AND NHTSA REGISTRATION NUMBER Required when Box 3 is checked', "TAPATIO AUTO AND TRUCK WRECKING"),
                ('DESCRIPTION OF MERCHANDISE IF MOTOR VEHICLE EQUIPMENT',V7),
                ('NAME OF IMPORTER Please type',"TAPATIO AUTO AND TRUCK WRECKING"),
                ('IMPORTERS ADDRESS Street City State ZIP Code',"2327 Siempre Viva Ct, San Diego,CA 92154"),
                ('NAME OF DECLARANT Please type'," MANUEL RAZO"),
                ('DECLARANTS ADDRESS',"2327 Siempre Viva Ct, San Diego,CA 92154"),
                ('DECLARANTS CAPACITY',"IMPORT/EXPORT"),
                ('DATE SIGNED',valor_celda)

           ] 
    
        fdf = forge_fdf("",fields,[],[],[])
        fdf_file = open("data.fdf","wb") 
        fdf_file.write(fdf) 
        fdf_file.close()
        comando = f"pdftk {infile} fill_form data.fdf output NATHASA_{j}.pdf flatten"
        subprocess.run(comando, shell=True)
        j += 1    

    j = 0 
    for i in range(0, len(df_filtrado)):
        fila = df_filtrado.iloc[i]
        if pd.isnull(fila[0]):
            break  
    #for i in range(13, len(df)):
    #    fila = df.iloc[i]
    #    if pd.isnull(fila[0]):
    #        break
        dato_m = fila[12].split('/')[0].strip()  # Separar el dato de la columna 'M' antes del "/"
        dato_n = fila[13].split('/')[0].strip()  # Separar el dato de la columna 'N' después del "/"
        dato_o = fila[12].split('/')[1].strip()  # Separar el dato de la columna 'M' antes del "/"
        dato_p = fila[13].split('/')[1].strip()  # Separar el dato de la columna 'N' después del "/"
        fila[8] = dato_m  # Actualizar el valor de la columna 'M'
        fila[12] = dato_n  # Actualizar el valor de la columna 'N'
        fila[9] = dato_o  # Actualizar el valor de la columna 'M'
        fila[13] = dato_p  # Actualizar el valor de la columna 'N'
        V1 = fila[13] + "/" + fila[12]  # Concatenar los datos de las columnas 13 y 12 en V1
        V2 = fila[9] + "/ENGINE ASSEMBLY"  # Concatenar los datos de las columnas 9 y 6 en V2
        V3 = fila[8]  # Asignar el valor de la columna 8 a V3
        if len(str(fila[11])) < 17 :
            V6 = fila[11]
        else:     
            V6 = fila[15] # Asignar el valor de la columna 15 a V6
        #V6 = fila[15]  # Asignar el valor de la columna 15 a V6
        V7 = fila[6]
        V8 = fila[12]
        V9 = fila[9]
        #datos_finales.append([V1, V2, V3, V6, valor_celda])

        fields2 = [ 
                ('4 Vehicle Identification Number VIN or engine serial number', V6),
                ('5 Manufacture date mmyyyy', V1),
                ('6 Manufacture make', V3),
                ('7 Model', V2),
                ('10 Owner', "\r\r2327 Siempre Viva Ct, San Diego,CA 92154"),
                ('11 Storage contact', "\r\r2327 Siempre Viva Ct, San Diego,CA 92154"),
                ('14 Date mmddyyyy', valor_celda)
            ]   

        fdf2 = forge_fdf("",fields2,[],[],[])
        fdf_file2 = open("data2.fdf","wb") 
        fdf_file2.write(fdf2) 
        fdf_file2.close()
        #salida = folder_path+f"\\forma3520_{j}.pdf"
        #print(salida)
        comando2 = f"pdftk {infile2} fill_form data2.fdf output forma3520_{j}.pdf flatten"
        subprocess.run(comando2, shell=True)
        j += 1    

        
    # Imprimir los valores de las variables V1, V2, V3, V4, V5 y V6
    print("Valor de la celda: ", valor_celda)
    print("Datos finales:")
    #for datos in datos_finales:
    #    print(datos)
    #for f in files:
    #        if (f.startswith("Apple") or f.startswith("apple")):
    #            shutil.move(f, dest1) 
    output_filename = "NTHSA"  # Nombre del archivo PDF resultante
    unir_archivos_pdf("NATHASA",folder_path, output_filename,valor_factura)
    output_filename = "EPA" 
    unir_archivos_pdf("forma3520",folder_path, output_filename,valor_factura)    

if __name__ == '__main__':
  conf = parse_args()
  #Principal(conf.Directorio_de_trabajo,conf.Archivo_Produccion,conf.Archivo_Rutas)
  Principal(conf.Factura_partes,conf.Directorio_de_trabajo)