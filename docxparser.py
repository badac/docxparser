# -*- coding: latin-1 -*-


import sys, os
import re
import regex
import glob
import traceback
import csv
from docx import Document


dir_path = sys.argv[1]
dest_file = sys.argv[2] or "fichas.csv"
processed_count = 0
unprocessed_count = 0
unprocessed_files = []

print("Scanning " + dir_path)

filenames = []
paths = []
directories = []

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
    "image_file",
    "img_file",
    "Tipo",
    "Archivo de origen"
    ]
print(table_headers)
#los strings en esta lista deben estar ordenados del más largo, plural al más sencillo y singular
trim_strings = [
    "Diapositiva/ Foto",
    "Diapositiva/Foto",
    "Diapositivas/Fotos",
    "imagen escaneada",
    "foto mala",
    "placa-diapo",
    "diapo-placa",
    "fotografía",
    "fotos",
    "foto",
    "diapositivas",
    "diapositiva",
    "scanner",
    "placa"
]

no_img_strings = [
    "sin registro",
    "sin imagen",
    "sin foto",
    "no hay foto",
    "no hay foto,",
    "no hay imagen",
    "Foto,pendiente",
    "pendiente"
]

expressions = []

for string in trim_strings:
    exp = re.compile(re.escape(string), re.IGNORECASE)
    expressions.append(exp)

for root, dirs, files in os.walk(dir_path, topdown=True):
    for name in files:
        if name.endswith(".docx") and not name.startswith("~$"):
            path = os.path.join(root, name)
            #print(path)
            filenames.append(name)
            paths.append(path)
            directories.append(root)


with open('output/' + dest_file, 'w') as csvfile:
    fichas_writer = csv.writer(csvfile, delimiter='\t')
    fichas_writer.writerow(table_headers)
    for path, filename, directory in zip(paths, filenames, directories):
        try:
            doc = Document(path)
            table = doc.tables[0]
            col = table.columns[1]
            values = []
            img_filename = "" #nombre de archivo en la celda, procesado
            img_file = "" #nombre de archivo encontrado en la carpeta JPG
            for index, cell in enumerate(col.cells):

                txt = cell.text # nombre de archivo en la celda, sin procesar

                # 4 es el campo de archivo de imagen
                txt = txt.replace('\r\n',',')
                txt = txt.replace('\n',',')
                txt = txt.replace('\r',',')
                txt = txt.replace('\t',',')

                if index == 4:
                    is_img = True
                    img_filename = txt
                    #Verificamos si el campo contiene indicacion de que no hay imagen
                    for string in no_img_strings:
                        if string.casefold() in img_filename.casefold():
                            is_img = False
                    #Si hay nombre de archivo de imagen, procesamos el texto
                    #sacando las partes que no son parte del nombre de archivo.
                    if is_img:

                        #img_filename = img_filename.replace('\r\n','')
                        #img_filename = img_filename.replace('\n','')
                        #img_filename = img_filename.replace('\r','')

                        #elimina palabras que no son parte del nombre de archivo
                        for exp in expressions:
                            img_filename = exp.sub('',img_filename)

                        #elimina caracteres no alfanumericos al ppio del archivo
                        e = re.compile("^\W+")
                        img_filename = e.sub('',img_filename)
                        print(path)
                        print(img_filename)

                        #listamos los archivos jpg del directorio de imagenes
                        img_dir_path = os.path.join(directory, "JPG" )
                        #print(img_dir_path)
                        if os.path.exists(img_dir_path):
                            #num_files = len([name for name in os.listdir(img_dir_path) if os.path.isfile(name)])
                            img_file = ""
                            img_list = []
                            types = ('*.jpg', '*.jpeg','*.JPG','*.JPEG')
                            for f in types:
                                img_list.extend(glob.glob(os.path.join(img_dir_path,f)))

                            #print(img_list)
                            if len(img_list) == 1:
                                img_file = img_list[0]
                            if len(img_list) > 1:
                                #fuzzy string matching
                                errs = []
                                for im in img_list:
                                    res = regex.fullmatch(r"(?:%s){i,d,s}" % img_filename ,"%s" % im ).fuzzy_counts
                                    s = sum(res)
                                    errs.append(s)
                                min_index = errs.index(min(errs))
                                img_file = img_list[min_index]

                # Agregamos los resultado a los valores.
                values.append(txt)

            values.append(img_filename)
            values.append(img_file)
            #print(values)
            #agrega tipo de entrada obra/boceto
            if "boceto".casefold() in filename.casefold() or "boceto".casefold() in values[0].casefold() :
                values.append("boceto")
            else:
                values.append("obra")
            #agrega la ruta al archivo .docx del que se extrae la info
            values.append(path)

            fichas_writer.writerow(values)
            processed_count += 1
        except Exception:
            e = sys.exc_info()[0]
            print("Error procesando archivo " + path + ". Ignorando: " )
            print(e)
            print(traceback.format_exc())
            unprocessed_count += 1
            unprocessed_files.append(path)
            pass

    #creacion de logs
    with open('output/docxparse_process.log', 'w') as log:
        log.write('Registro de procesamiento de archivos'  + '\n\n')
        log.write('Cantidad de archivos procesados: ' + str(processed_count) + '\n')
        log.write('Cantidad de archivos sin procesar: ' + str(unprocessed_count)  + '\n' )
        log.write('Lista de archivos sin procesar:'  + '\n')
        for f in unprocessed_files:
            log.write(f + '\n')
