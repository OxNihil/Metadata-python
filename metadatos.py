# -*- encoding: utf-8 -*-
#python 3.6


#Modulos
from PyPDF2 import PdfFileReader, PdfFileWriter
import os 
import exifread
import eyed3
import sys
import  docx
from eyed3 import id3

#Funcion principal

def main():
    print('##############################################')
    print('#         Metadata Recursive Analyzer        #')
    print('##############################################')
    print('Ext: Docs,PDF,JPG,TIFF,MP3 \n ')
    if len(sys.argv) > 1:
        directorio = sys.argv[1]
    else:
        directorio = input('[-]Introduce la ruta del directorio o unidad: ')
    if os.path.exists(directorio) == False:
        print('[!]El directorio introducido no es valido \n')
        main()
    elif len(directorio) == 2 and len(directorio.split(':')[0]) == 1 :
        print('[!]Porfavor introduce la unidad en un formato adecuado,ejemplo > H:\ \n')
        main()
    printMeta(directorio)

def printMeta(directorio):
    try:
        print('Analizando el directorio: ' + directorio, '\n')
        for dirpath, dirnames, files in os.walk(directorio):
            for name in files:
                ext = name.lower().rsplit('.', 1)[-1]
                #Documentos
                if ext in ['pdf']:
                    print ("[*] Metadatos del archivo: %s " %(dirpath+os.path.sep+name))
                    print ('----------------------------------------------------------')
                    try:
                        pdfFile = PdfFileReader(open(dirpath+os.path.sep+name, 'rb'))#abrimos el fichero
                        docInfo = pdfFile.getDocumentInfo() #creamos un diccionario con la info recolectada
                    
                        for metaItem in docInfo:
                            print ('[+]' + metaItem + ':' + str(docInfo[metaItem]))
                            
                        docInfoextra = {pdfFile.getNumPages():'Numero de paginas: ',
                                        pdfFile.getPageMode(): 'Modo de la pagina: ',
                                        pdfFile.isEncrypted: 'Encriptacion: ',
                                        pdfFile.getFields(): 'Campos de texto: '}
                        for element in docInfoextra:
                            if element != None:
                                print('[+]/'+docInfoextra[element]+str(element))
                        xmpinfo = pdfFile.getXmpMetadata()
                    except:
                        pass
                    if xmpinfo != None:
                        if hasattr(xmpinfo,'dc_contributor'):
                            print ('[+]/'+'dc_contributor', xmpinfo.dc_contributor)
                        elif hasattr(xmpinfo,'dc_identifier'):
                            print ('[+]/'+'dc_identifier', xmpinfo.dc_identifier)
                        elif hasattr(xmpinfo,'dc_date'):
                            print ('[+]/'+'dc_date', xmpinfo.dc_date)
                        elif hasattr(xmpinfo,'dc_source'):
                            print ('[+]/'+'dc_source', xmpinfo.dc_source)
                        elif hasattr(xmpinfo,'dc_subject'):
                            print ('[+]/'+'dc_subject', xmpinfo.dc_subject)
                        elif hasattr(xmpinfo,'xmp_modifyDate'):
                            print ('[+]/'+'xmp_modifyDate', xmpinfo.xmp_modifyDate)
                        elif hasattr(xmpinfo,'xmp_metadataDate'):
                            print ('[+]/'+'xmp_metadataDate'), xmpinfo.xmp_metadataDate
                        elif hasattr(xmpinfo,'xmpmm_documentId'):
                            print ('[+]/'+'xmpmm_documentId', xmpinfo.xmpmm_documentId)
                        elif hasattr(xmpinfo,'xmpmm_instanceId'):
                            print('[+]/'+'xmpmm_instanceId', xmpinfo.xmpmm_instanceId)
                        elif hasattr(xmpinfo,'pdf_keywords'):
                            print ('[+]/'+'pdf_keywords', xmpinfo.pdf_keywords)
                        elif hasattr(xmpinfo,'pdf_pdfversion'):
                            print ('[+]/'+'pdf_pdfversion', xmpinfo.pdf_pdfversion)
                    print("\n")
                #Imagenes
                elif ext in ['jpg','tiff']:
                    print ("[*] Metadatos del archivo: %s " %(dirpath+os.path.sep+name))
                    print ('----------------------------------------------------------')
                    f = open(dirpath+os.path.sep+name,'rb')
                    tags = exifread.process_file(f)
                    if len(tags) == 0 :
                        print('[!]No hay metadatos')
                    for tag in tags.keys():
                        if tag not in ('JPEGThumbnail', 'TIFFThumbnail', 'Filename', 'EXIF MakerNote'):
                            print ("[+]: %s, valor %s" % (tag, tags[tag]))
                    print("\n")
                #Musica
                elif ext in ['mp3']:
                    print ("[*] Metadatos del archivo: %s " %(dirpath+os.path.sep+name))
                    print ('----------------------------------------------------------')
                    tag = eyed3.id3.Tag()
                    tag.parse(dirpath+os.path.sep+name)
                    if tag.artist is not None:
                        print('Artista: ',tag.artist)
                    if tag.album is not None:
                        print('Album: ',tag.album)
                    if tag.title is not None:
                        print('Titulo: ',tag.title)
                    if tag.track_num[0] is not None:
                        print('Track: ',tag.track_num[0])
                    else:
                        print('[!]No hay metadatos')
                #Docs
                elif ext in ['docs']: 
                    print ("[*] Metadatos del archivo: %s " %(dirpath+os.path.sep+name))
                    print ('----------------------------------------------------------')
                    document = docx.Document(docx = dirpath+os.path.sep+name)
                    core_properties = document.core_properties
                    print(core_properties.author)
                    print(core_properties.created)
                    print(core_properties.last_modified_by)
                    print(core_properties.last_printed)
                    print(core_properties.modified)
                    print(core_properties.revision)
                    print(core_properties.title)
                    print(core_properties.category)
                    print(core_properties.comments)
                    print(core_properties.identifier)
                    print(core_properties.keywords)
                    print(core_properties.language)
                    print(core_properties.subject)
                    print(core_properties.version)
                    print(core_properties.keywords)
                    print(core_properties.content_status)
                    
        print('[+]Ejecucion finalizada')
        
    except(KeyboardInterrupt, SystemExit):
        print('[!]Se ha interrumpido la ejecucion')
    except:
        print("Unexpected error:", sys.exc_info()[0])
main()
