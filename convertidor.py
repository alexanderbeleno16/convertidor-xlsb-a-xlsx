# from datetime import datetime
import openpyxl
from pyxlsb import open_workbook as open_xlsb
from pyxlsb import convert_date
import sys
import os.path

# ------ PIP ------
# pip install pyxlsb https://pypi.org/project/pyxlsb/
# pip install openpyxl https://pypi.org/project/openpyxl/


def extraerColumnasFechas(Row):
    columnas_fecha = []
    columna = 0
    for r in Row:
        if ( 'FECHA' in r.decode("utf-8").upper()):
            columnas_fecha.append(columna)
            
        columna += 1
    
    return columnas_fecha

def convertirXLSB(nombre_excelXLSB):
    wbNew = openpyxl.Workbook()
    hoja = wbNew.active
    with open_xlsb(nombre_excelXLSB) as wb:
        with wb.get_sheet(2) as sheet:
            fila = 1
            columnas_fecha = []
            for row in sheet.rows():
                # print(fila)
                
                row_aux = [None if item.v == None else str(item.v).replace('', ' ').encode(encoding="utf-8",errors="xmlcharrefreplace") for item in row]

                if(fila==1):
                    columnas_fecha = extraerColumnasFechas(row_aux)
                    # print('columna fecha: ',columnas_fecha)
                else:
                    conversionFechasFila(columnas_fecha, row_aux)

                try:
                    hoja.append(row_aux)
                except openpyxl.utils.exceptions.IllegalCharacterError:
                    print('error de escritura para la fila:'+str(fila))
                    quit()
                
                fila+=1

    guardarArchivoNuevo(wbNew, nombre_excelXLSB)

def conversionFechasFila(columnafecha=[], row=[]):
    if(len(columnafecha) > 0):
        for cf in columnafecha:
            
            if( row[cf] != None and row[cf].decode('utf-8').replace('.', '').isnumeric()):
                
                row[cf]=convert_date(float(row[cf].decode('utf-8'))).strftime("%Y-%m-%d").encode(encoding="utf-8") 


def convertirReparto(nombre_excelXLSB):
    wbNew = openpyxl.Workbook()
    hoja = wbNew.active
    with open_xlsb(nombre_excelXLSB) as wb:
        with wb.get_sheet(2) as sheet:
            fila = 1
            columnas_fecha = []
            for row in sheet.rows():
                row_aux = [str(item.v).encode(encoding="utf-8",errors="xmlcharrefreplace") for item in row]

                # row_aux = [None if item.v == None else str(item.v).replace('', ' ').encode(encoding="utf-8",errors="xmlcharrefreplace") for item in row]

                row_aux.pop()

                if(fila==1):
                    columnas_fecha = extraerColumnasFechas(row_aux)
                    # print(columnas_fecha)
                else:
                    conversionFechasFila(columnas_fecha, row_aux)

                try:
                    hoja.append(row_aux)
                except openpyxl.utils.exceptions.IllegalCharacterError:
                    print('error de escritura para la fila:'+str(fila))
                    quit()
                
                fila+=1

    guardarArchivoNuevo(wbNew, nombre_excelXLSB)

def guardarArchivoNuevo(wbNew, nombre=""):
    nombre_guardar = os.path.basename(nombre).replace('.xlsb', '.xlsx')
    
    try:
        if(sys.argv[3]):
            nombre_guardar = sys.argv[3]
    except IndexError:
        pass

    try:
        wbNew.save(nombre_guardar)            
        print("-----------------------------------------------")
        print("Conversión de xlsb a xlsx exitosa!")
        print("-----------------------------------------------")
    except PermissionError:
        print("")
        print("###############################################################")
        print("Tiene en ejecución un archivo excel con el mismo nombre, cerrar")
        print("Nombre del archivo en ejecución: {}".format(nombre_guardar))
        print("###############################################################")
        print("")
    except FileNotFoundError:
        print("No se pudo acceder al directorio '{}' ".format(os.path.dirname(nombre_guardar)))
    

try:
    if(sys.argv[1] == ''):
        raise IndexError('Nombre del archivo vacio')
    funcion='N/D'
    file_name = sys.argv[1]
    try:
        if(sys.argv[2] != '' and sys.argv[2] != '-d'):
            funcion = sys.argv[2]
    except IndexError:
        pass

    if(funcion.upper() == 'REPARTO'):
        # print('REPARTO')
        convertirReparto(sys.argv[1])
    elif(funcion.upper() == 'N/D'):
        # print('generico')
        convertirXLSB(sys.argv[1])
    else:
        print("")
        print("")
        print("##############################################")
        print("Advertencia: Opción digitada no se encuentra")
        print("Por favor, especifique la acción a realizar...")
        print("##############################################\n")
        print("-----------------------------------------------")
        print("Opciones:                                     |")
        print("----------------------------------------------|")
        print("REPARTO                                       |")
        print("-----------------------------------------------")
        print("")
    
    print('Nombre Archivo: '+file_name,', ','Opción: '+funcion)
    print("")
except IndexError:
    print('Se requiere el nombre del archivo',IndexError)

