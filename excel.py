# -*- coding: utf-8 -*-
from datetime import date, time
from decimal import Decimal
import xlwt
import sqlite3


def procesar_codigos(datos):
    conn = sqlite3.connect('database.db')
    lista = []
    for detalle in datos:
        r = [" ".join(record)
             for record in conn.execute("select ape_pat, ape_mat, nombres from clientes where codigo='%s'" % detalle[0])]
        if len(r) > 0:
            depositante = r[0]
        else:
            depositante = "DESCONOCIDO"
        lista.append(depositante)
    conn.close()
    return lista


def generar_excel(dirname, datos):
    cabecera = datos[0]
    detalles = datos[1]
    # Incluir el nombre del depositante en los detalles
    lista_depositantes = procesar_codigos(detalles)
    for i, detalle in enumerate(detalles):
        detalle[1] = lista_depositantes[i]
    
    ruta_archivo = dirname + "/"
    nombre_archivo = "reporte.xls"
    
    style = xlwt.XFStyle()
    styles = {'default': xlwt.easyxf('font: name Liberation Mono, height 180;',
                                     num_format_str='@'),
              'headers': xlwt.easyxf('font: name Liberation Mono, height 180, colour white, bold 1;'
                                     'pattern: pattern solid, fore_colour black;',
                                     num_format_str='@'),
              'h_default': xlwt.easyxf('font: name Liberation Mono, height 180; align: horiz right',
                                       num_format_str='@'),
              'decimal': xlwt.easyxf('font: name Liberation Mono, height 180;',
                                     num_format_str='0.00'),
              'time': xlwt.easyxf('font: name Liberation Mono, height 180;',
                                  num_format_str='HH:MM:SS AM/PM'),
              'date': xlwt.easyxf('font: name Liberation Mono, height 180;',
                                  num_format_str='dd/mm/yyyy')}
    book = xlwt.Workbook(encoding='utf8')
    sheet = book.add_sheet('Hoja 1')
    i = 1
    # Datos de Cabecera
    encabezado_1 = ["Cuenta",
                    "Código Interno",
                    "Fecha de Proceso",
                    "Total de Registros",
                    "Monto Total",
                    "Código Teletransfer"]
    
    encabezado_2 = ["Código",
                    "Depositante",
                    "Retorno",
                    "Dato Adicional",
                    "Fecha de Pago",
                    "Fecha de Vencimiento",
                    "Monto Pagado",
                    "Mora Pagada",
                    "Monto Total",
                    "Sucursal",
                    "Nº de Operación",
                    "Referencia",
                    "Terminal",
                    "Medio de atención",
                    "Hora de atención",
                    "Número de cheque",
                    "Código del banco"]
    
    cell_style = styles['headers']
    
    colx = 0
    rowx = 0
    for value in encabezado_1:
        if rowx == 3:
            colx = 4
            rowx = 0
        sheet.write(rowx+i, colx, value, style=cell_style)
        sheet.col(colx).width = 5010
        rowx += 1
        
    for colx, value in enumerate(encabezado_2):
        sheet.write(rowx+i, colx, value, style=cell_style)
        # Ajustar el ancho
        if colx == 2:
            sheet.col(colx).width = 7000
        elif colx not in [0, 4]:
            sheet.col(colx).width = 320 * len(value)
        
    colx = 1
    rowx = 0
    for value in cabecera:
        if type(value) is Decimal:
            cell_style = styles['decimal']
        else:
            cell_style = styles['h_default']
            
        # Establezco el ancho de la 2da y 6ta columna
        if rowx == 0:
            sheet.col(colx).width = 5200
        elif rowx == 3:
            colx = 5
            rowx = 0
            sheet.col(colx).width = 4090
        sheet.write(rowx+i, colx, value, style=cell_style)
        
        rowx += 1
        
    # Datos de los Detalles
    i = rowx + i + 1
    for rowx, row in enumerate(detalles):
        for colx, value in enumerate(row):
            if type(value) is Decimal:
                cell_style = styles['decimal']
            elif type(value) is date:
                cell_style = styles['date']
            elif type(value) is time:
                cell_style = styles['time']
            else:
                cell_style = styles['default']
            sheet.write(rowx+i, colx, value, style=cell_style)
            
    full_path = ruta_archivo + nombre_archivo
    book.save(full_path)
    