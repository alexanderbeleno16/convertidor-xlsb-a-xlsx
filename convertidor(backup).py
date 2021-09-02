from datetime import datetime
import openpyxl
from pyxlsb import open_workbook as open_xlsb
from pyxlsb import convert_date
# import re

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

def convertirXLSB():
    wbNew = openpyxl.Workbook()
    hoja = wbNew.active
    with open_xlsb('Prueba_Nueva.xlsb') as wb:
        with wb.get_sheet(2) as sheet:
            fila = 1
            columnas_fecha = []
            for row in sheet.rows():
                print(fila)
                
                # format('')
                row_aux = [None if item.v == None else str(item.v).replace('', ' ').encode(encoding="utf-8",errors="xmlcharrefreplace") for item in row]

                if(fila==1):
                    columnas_fecha = extraerColumnasFechas(row_aux)
                    print('columna fecha: ',columnas_fecha)
                else:
                    if(len(columnas_fecha) > 0):
                        for cf in columnas_fecha:
                            # print('fila:'+str(fila), 'columna:'+str(cf), row_aux[cf].decode('utf-8'))
                            
                            if( row_aux[cf] != None and row_aux[cf].decode('utf-8').replace('.', '').isnumeric()):
                                # print('fila:'+str(fila), 
                                    #     'columna:'+str(cf), 
                                    #     'fecha_excel:', row_aux[cf].decode('utf-8'), 
                                    #     'conversion:',convert_date(float(row_aux[cf].decode('utf-8'))),
                                    #     'caracteres: '+convert_date(float(row_aux[cf].decode('utf-8'))).strftime("%Y-%m-%d")
                                #     )
                                row_aux[cf]=convert_date(float(row_aux[cf].decode('utf-8'))).strftime("%Y-%m-%d").encode(encoding="utf-8") 

                try:
                    hoja.append(row_aux)
                except openpyxl.utils.exceptions.IllegalCharacterError:
                    print('error de escritura para la fila:'+str(fila))
                    quit()
                fila+=1

    wbNew.save('Prueba_Nueva5.xlsx')


def convertirReparto():
    wbNew = openpyxl.Workbook()
    hoja = wbNew.active
    with open_xlsb('REPARTO1955.xlsb') as wb:
        with wb.get_sheet(2) as sheet:
            fila = 1
            columnas_fecha = []
            for row in sheet.rows():
                row_aux = [str(item.v).encode(encoding="utf-8",errors="xmlcharrefreplace") for item in row]
                row_aux.pop()

                if(fila==1):
                    columnas_fecha = extraerColumnasFechas(row_aux)
                    print(columnas_fecha)
                else:
                    if(len(columnas_fecha) > 0):
                        for ff in columnas_fecha:
                            print('fila:'+str(fila), 'columna:'+str(ff), row_aux[ff].decode('utf-8'))

                hoja.append(row_aux)
                fila+=1


            # fila = 1
            # columnas_fecha = []
            # for row in sheet.rows():
            #     print(fila)
                
            #     # format('')
            #     row_aux = [None if item.v == None else str(item.v).replace('', ' ').encode(encoding="utf-8",errors="xmlcharrefreplace") for item in row]

            #     if(fila==1):
            #         columnas_fecha = extraerColumnasFechas(row_aux)
            #         print('columna fecha: ',columnas_fecha)
            #     else:
            #         if(len(columnas_fecha) > 0):
            #             for cf in columnas_fecha:
            #                 # print('fila:'+str(fila), 'columna:'+str(cf), row_aux[cf].decode('utf-8'))
                            
            #                 if( row_aux[cf] != None and row_aux[cf].decode('utf-8').replace('.', '').isnumeric()):
            #                     # print('fila:'+str(fila), 
            #                         #     'columna:'+str(cf), 
            #                         #     'fecha_excel:', row_aux[cf].decode('utf-8'), 
            #                         #     'conversion:',convert_date(float(row_aux[cf].decode('utf-8'))),
            #                         #     'caracteres: '+convert_date(float(row_aux[cf].decode('utf-8'))).strftime("%Y-%m-%d")
            #                     #     )
            #                     row_aux[cf]=convert_date(float(row_aux[cf].decode('utf-8'))).strftime("%Y-%m-%d").encode(encoding="utf-8") 

            #     try:
            #         hoja.append(row_aux)
            #     except openpyxl.utils.exceptions.IllegalCharacterError:
            #         print('error de escritura para la fila:'+str(fila))
            #         quit()
            #     fila+=1

    wbNew.save('prueba8.xlsx')


convertirXLSB()



            
