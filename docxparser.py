# -*- coding: latin-1 -*-


import sys, os, getopt
import re
import regex
import glob
import traceback
import csv
from shutil import copyfile
from docx import Document



class DocxParser:
    def __init__(self, _dir_path, _dest_file, _img_out, _url_prefix):
        self.dir_path = _dir_path
        self.dest_file = _dest_file or "fichas.csv"
        self.img_out = _img_out or None
        self.copy = False
        self.url_prefix = _url_prefix or None
        self.processed_count = 0
        self.unprocessed_count = 0
        self.unprocessed_files = []
        self.filenames = []
        self.paths = []
        self.directories = []

        self.table_headers = [
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
            "img_filename",
            "img_file",
            "img_url",
            "Tipo",
            "Archivo de origen"
            ]

        #los strings en esta lista deben estar ordenados del más largo, plural al más sencillo y singular
        self.trim_strings = [
            "Diapositiva/ Foto",
            "Diapositiva/Foto",
            "Dipositiva/Foto",
            "Diapositivas/Fotos",
            "foto-diapo",
            "foto/placa",
            "imagen escaneada",
            "foto mala",
            "placa-diapo",
            "diapo-placa",
            "fotografía",
            "fotos",
            "foto",
            "diapositivas",
            "diapositiva",
            "dipositiva",
            "diapositivo",
            "scanner",
            "placa"
        ]

        self.no_img_strings = [
            "sin registro",
            "sin imagen",
            "sin foto",
            "no hay foto",
            "no hay foto,",
            "no hay imagen",
            "Foto,pendiente",
            "pendiente"
        ]

        self.expressions = []

        for string in self.trim_strings:
            exp = re.compile(re.escape(string), re.IGNORECASE)
            self.expressions.append(exp)
    #Escanea el directorio extrayendo nombres de subdirectorios y archivos
    def scanDir(self):
        print("Scanning " + self.dir_path)

        for root, dirs, files in os.walk(self.dir_path, topdown=True):
            for name in files:
                if name.endswith(".docx") and not name.startswith("~$"):
                    path = os.path.join(root, name)
                    #print(path)
                    self.filenames.append(name)
                    self.paths.append(path)
                    self.directories.append(root)


    #extrae la información de los archivos y directorios
    def extractInfo(self):
        with open(self.dest_file, 'w') as csvfile:
            fichas_writer = csv.writer(csvfile, delimiter='|')
            fichas_writer.writerow(self.table_headers)
            for path, filename, directory in zip(self.paths, self.filenames, self.directories):
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
                            for string in self.no_img_strings:
                                if string.casefold() in img_filename.casefold():
                                    is_img = False
                                    img_filename = ""
                                    break
                            #Si hay nombre de archivo de imagen, procesamos el texto
                            #sacando las partes que no son parte del nombre de archivo.
                            if is_img:

                                #img_filename = img_filename.replace('\r\n','')
                                #img_filename = img_filename.replace('\n','')
                                #img_filename = img_filename.replace('\r','')

                                #elimina palabras que no son parte del nombre de archivo
                                for exp in self.expressions:
                                    img_filename = exp.sub('',img_filename)

                                #elimina caracteres no alfanumericos al ppio del archivo
                                e = re.compile("^\W+")
                                img_filename = e.sub('',img_filename)
                                #print(path)
                                #print(img_filename)

                                #listamos los archivos jpg del directorio de imagenes
                                dir_names = ["JPEG", "JPG"]
                                for dir_name in dir_names:
                                    img_dir_path = os.path.join(directory, dir_name )

                                    #print(img_dir_path)
                                    if os.path.exists(img_dir_path):
                                        #num_files = len([name for name in os.listdir(img_dir_path) if os.path.isfile(name)])
                                        #img_file = ""
                                        print("path exsits")
                                        img_list = []
                                        types = ('*.jpg', '*.jpeg','*.JPG','*.JPEG')
                                        for f in types:
                                            img_list.extend(glob.glob(os.path.join(img_dir_path,f)))

                                        print(len(img_list))
                                        if len(img_list) == 1:
                                            img_file = img_list[0]
                                            print("add one file: " + img_file)

                                        if len(img_list) > 1:
                                            #fuzzy string matching
                                            errs = []
                                            for im in img_list:
                                                res = regex.fullmatch(r"(?:%s){i,d,s}" % img_filename ,"%s" % im ).fuzzy_counts
                                                s = sum(res)
                                                errs.append(s)
                                            #print(img_list)
                                            #print(errs)
                                            min_index = errs.index(min(errs))
                                            #print(min_index)
                                            img_file = img_list[min_index]
                                            print("chosen file: " + img_file)
                                    else:
                                        print("this dir doesn't exists: " + img_dir_path)
                        # Agregamos los resultado a los valores.
                        values.append(txt)

                    values.append(img_filename)
                    values.append(img_file)

                    #copia imagenes a la carpeta de destino
                    img_url = ""
                    if self.img_out:
                        if img_file:
                            basename = os.path.basename(img_file)
                            dest = os.path.join(self.img_out, basename)
                            copyfile(img_file, dest)
                            # si tenemos prefijo de url, construímos la url
                            if self.url_prefix:
                                img_url = self.url_prefix + dest

                    print("add file: " + img_file)

                    values.append(img_url)
                    #print(values)
                    #agrega tipo de entrada obra/boceto
                    if "boceto".casefold() in filename.casefold() or "boceto".casefold() in values[0].casefold() :
                        values.append("boceto")
                    else:
                        values.append("obra")
                    #agrega la ruta al archivo .docx del que se extrae la info
                    values.append(path)

                    fichas_writer.writerow(values)
                    self.processed_count += 1

                except Exception:
                    e = sys.exc_info()[0]
                    print("Error procesando archivo " + path + ". Ignorando: " )
                    print(e)
                    print(traceback.format_exc())
                    self.unprocessed_count += 1
                    self.unprocessed_files.append(path)
                    pass

    def logInfo(self):
        #creacion de logs
        with open('output/docxparse_process.log', 'w') as log:
            log.write('Registro de procesamiento de archivos'  + '\n\n')
            log.write('Cantidad de archivos procesados: ' + str(self.processed_count) + '\n')
            log.write('Cantidad de archivos sin procesar: ' + str(self.unprocessed_count)  + '\n' )
            log.write('Lista de archivos sin procesar:'  + '\n')
            for f in self.unprocessed_files:
                log.write(f + '\n')

    def run(self):
        self.scanDir()
        self.extractInfo()
        self.logInfo()



#mensaje de uso
def usage():
    usage_msg = "Usage: python docxparse.py -h -i <input_dir> -o <output_file> -c <copy_images_dir>"
    print(usage_msg)

def main(argv):
    dir_path = ""
    dest_file = "fichas.csv"
    processed_count = 0
    img_url = ""
    unprocessed_count = 0
    unprocessed_files = []
    img_out  = "output/img/"

    try:
        opts, args = getopt.getopt(argv, "hi:o:c:u:",["help", "indir=","outfile=","copy=", "urlprefix"])
    except getopt.getopt.GetoptError:
        usage()
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()
            sys.exit()
        elif opt in ("-i", "--indir"):
            print("arg: " + arg)
            dir_path = arg

        elif opt in ("-o", "--outfile"):
            dest_file = arg

        elif opt in ("-c", "--copy"):
            img_out = arg
        elif opt in ("-u", "urlprefix"):
            img_url = arg

        docxparser = DocxParser(dir_path, dest_file, img_out, img_url)
        docxparser.run()


#main loop
if __name__ == "__main__":
    main(sys.argv[1:])
