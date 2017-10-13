# -*- coding: latin-1 -*-


import sys, os
import csv
from docx import Document


dir_path = sys.argv[1]
dest_file = sys.argv[2] ¡ "fichas.csv"

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
    "Referente de prensa"
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

        print(values)
        fichas_writer.writerow(values)
