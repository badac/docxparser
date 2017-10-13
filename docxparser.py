# -*- coding: latin-1 -*-


import sys, os
import csv
from docx import Document


dir_path = sys.argv[1]
dest_file = sys.argv[2] or "fichas.csv"
processed_count = 0
unprocessed_count = 0
unprocessed_files = []

print("Scanning " + dir_path)

filenames = []
table_headers = [
    "Titulo",
    "Fecha",
    "Tecnica",
    "Dimensiones",
    "Imagen.: Formato/ Nombre/Autor",
    "Creditos",
    "Propietarios",
    "Exposiciones",
    "Publicaciones",
    "Referente de prensa",
    "Archivo de origen"
    ]
print(table_headers)

for root, dirs, files in os.walk(dir_path, topdown=True):
    for name in files:
        if name.endswith(".docx") and not name.startswith("~$"):
            path = os.path.join(root, name)
            print(path)
            filenames.append(path)


with open('output/' + dest_file, 'w') as csvfile:
    fichas_writer = csv.writer(csvfile, delimiter='\t')
    fichas_writer.writerow(table_headers)
    for filename in filenames:
        try:
            doc = Document(filename)
            table = doc.tables[0]
            col = table.columns[1]
            values = []
            for cell in col.cells:
                txt = cell.text
                txt = txt.replace('\r\n',',')
                txt = txt.replace('\n',',')
                txt = txt.replace('\r',',')
                values.append(txt)
            #print(values)
            values.append(filename)
            fichas_writer.writerow(values)
            processed_count += 1
        except Exception:
            print("Error procesando archivo " + filename + ". Ignorando")
            unprocessed_count += 1
            unprocessed_files.append(filename)
            pass


    with open('output/docxparse_process.log', 'w') as log:
        log.write('Registro de procesamiento de archivos'  + '\n\n')
        log.write('Cantidad de archivos procesados: ' + str(processed_count) + '\n')
        log.write('Cantidad de archivos sin procesar: ' + str(unprocessed_count)  + '\n' )
        log.write('Lista de archivos sin procesar:'  + '\n')
        for f in unprocessed_files:
            log.write(f + '\n')
