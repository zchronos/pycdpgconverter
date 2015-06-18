# -*- coding: utf-8 -*-
from datetime import date, time
from decimal import Decimal
from zipfile import ZipFile
import os
import wx
import xlwt
import sqlite3


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
                                  num_format_str='dd/mm/yyyy'),}
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
    
class Empresas:
    def __init__(self, empresa=None):
        self.database = "database.db"
        self.abreviatura = empresa
        if empresa is None:
            self.todas()
        else:
            self.obtener_cuenta()
            self.obtener_razon_social()
        
    def obtener_cuenta(self, *args, **kwargs):
        conn = sqlite3.connect(self.database)
        r = [record[0] for record in conn.execute("select cuenta from empresas where abreviatura='%s'" % self.abreviatura)]
        conn.close()
        self.cuenta = r[0]
    
    def obtener_razon_social(self, *args, **kwargs):
        conn = sqlite3.connect(self.database)
        r = [record[0] for record in conn.execute("select razon_social from empresas where abreviatura='%s'" % self.abreviatura)]
        conn.close()
        self.razon_social = r[0]
        
    def todas(self, *args, **kwargs):
        conn = sqlite3.connect(self.database)
        r = [record[0] for record in conn.execute("select abreviatura from empresas")]
        conn.close()
        return r
    
def procesar_cdpg(cdpgfile, empresa_alias):
    cabecera = []
    detalles = []
    for i, line in enumerate(cdpgfile):
        if i == 0:
            # Cabecera
            cuenta = Empresas(empresa_alias).cuenta
            cuenta_txt = line[2:13]
            #comparar
            if cuenta_txt[:3] != cuenta[:3] or cuenta_txt[4:11] != cuenta[4:11]:
                print "Error: Cuenta Incorrecta"
                return ["Error", "Cuenta Incorrecta"]
            
            fecha_proceso = date(
                              int(line[14:18]),
                              int(line[18:20]),
                              int(line[20:22])).strftime('%d/%m/%Y')
            total_registros = int(line[22:31])
            total_monto = Decimal(line[31:44] + "." + line[44:46])
            cod_interno = int(line[46:50])
            cod_teletransfer = line[50:56]
            cabecera.extend([cuenta, cod_interno, fecha_proceso,
                             total_registros, total_monto, cod_teletransfer])
            if total_registros == 0:
                break
        else:
            # Detalle
            codigo = line[13:27].replace(" ", "")
            depositante = " "
            retorno = line[27:52]
            dato_adicional = line[52:57]
            fecha_pago = date(
                              int(line[57:61]),
                              int(line[61:63]),
                              int(line[63:65]))
            fecha_vencimiento = date(
                                     int(line[65:69]),
                                     int(line[69:71]),
                                     int(line[71:73]))
            monto_pagado = Decimal(line[73:86] + "." + line[86:88])
            mora_pagado = Decimal(line[88:101] + "." + line[101:103])
            monto_total = Decimal(line[103:116] + "." + line[116:118])
            sucursal = int(line[118:124].rjust(6).replace(" ", "0"))
            num_operacion = int(line[124:130].rjust(6).replace(" ", "0"))
            referencia = line[130:152]
            terminal_id = line[152:156]
            atencion_medio = line[156:168].rjust(12).replace(" ", "")
            atencion_hora = time(
                                int(line[168:170]),
                                int(line[170:172]),
                                int(line[172:174]))
            num_cheque = line[174:184].rjust(10).replace(" ", "")
            cod_banco = line[184:186].rjust(2).replace(" ", "")
            
            # Convertir números (mejora la estética)
            if codigo.isdigit():
                codigo = int(codigo)
            if dato_adicional.isdigit():
                dato_adicional = int(dato_adicional)
            if atencion_medio.isdigit():
                atencion_medio = int(atencion_medio)
            
            detalles.extend([[codigo, depositante, retorno, dato_adicional, fecha_pago,
                              fecha_vencimiento, monto_pagado, mora_pagado,
                              monto_total, sucursal, num_operacion, referencia,
                              terminal_id, atencion_medio, atencion_hora,
                              num_cheque, cod_banco]])
    if len(detalles) == 0:
        return ["Error", "No hay datos"]
    else:
        return [cabecera, detalles]
    
def procesar_codigos(datos):
    conn = sqlite3.connect('database.db')
    lista = []
    for detalle in datos:
        r = [(record[0] + " " + record[1] + " " + record[2]) for record in conn.execute("select ape_pat, ape_mat, nombres from clientes where codigo='%s'" % detalle[0])]
        if len(r) > 0:
            depositante = r[0]
        else:
            print "Depositante Deconocido: ", detalle[0]
            depositante = "DESCONOCIDO"
        lista.append(depositante)
    conn.close()
    return lista
    
def abrir_archivo(ruta_archivo, nombre_archivo):
    ext_archivo = str(nombre_archivo).lower()[-3:]
    archivo = None
    
    if ext_archivo == "txt":
        archivo = open('%s/%s' % (ruta_archivo, nombre_archivo), 'rb')
    elif ext_archivo == "zip":
        # El archivo CDPG___.txt siempre tiene el mismo número del ZIP
        archivo = ZipFile('%s/%s' % (ruta_archivo, nombre_archivo),
                          'r').open(nombre_archivo[:-3] + 'TXT', 'rU')
        
    return archivo

    
    
class MainWindow(wx.Frame):
    def __init__(self, parent, title):
        self.dirname = ''
        wx.Frame.__init__(self, parent, title=title, size=(310,200))
        self.CreateStatusBar()
        
        # ComboBox
        empresaD = {}
        empresas = Empresas()
        for empresa in empresas.todas():
            empresaD[empresa] = empresa
        
        empresaList = sorted(empresaD.keys())
        
        wx.StaticText(self, -1, "Seleccione una empresa: ", (10, 10))
        self.combo1 = wx.ComboBox(self, -1, value=".......", pos=wx.Point(10, 30),
                                  size=wx.Size(150, 28), choices=empresaList)
        
        # Menu
        filemenu = wx.Menu()
        menuOpen = filemenu.Append(wx.ID_OPEN, "&Abrir", " Abrir un archivo")
        menuExit = filemenu.Append(wx.ID_EXIT, "&Salir", " Finaliza el programa")
        #
        helpmenu = wx.Menu()
        menuAbout = helpmenu.Append(wx.ID_ABOUT, "A&cerca de", " Información sobre el programa")
        
        # MenuBar
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, "&Archivo")
        menuBar.Append(helpmenu, "&Ayuda")
        self.SetMenuBar(menuBar)
        
        # Button
        self.button1 = wx.Button(self, -1, label="Convertir",
            pos=wx.Point(170, 30), size=wx.Size(130, 28))
        
        # Set events.
        self.Bind(wx.EVT_MENU, self.OnOpen, menuOpen)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        
        self.button1.Bind(wx.EVT_BUTTON, self.button1Click, self.button1)
        
        self.Show(True)
        
    def button1Click(self,event):
        self.empresa_alias = self.combo1.GetValue()
        f = abrir_archivo(self.dirname, self.filename)
        datos = procesar_cdpg(f, self.empresa_alias)
        f.close()
        
        if datos[0] == "Error":
            dlg = wx.MessageDialog(self, "Error: " + datos[1], "Mensaje", wx.OK)
        else:
            generar_excel(self.dirname, datos)
            dlg = wx.MessageDialog(self, "Archivo generado correctamente", "Mensaje", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()

    def OnAbout(self,e):
        dlg = wx.MessageDialog(self, "Convierte CDPG del banco a Excel \n\n\t V0.2",
                               "PyCDPGConverter", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()

    def OnExit(self,e):
        self.Close(True)
        
    def OnOpen(self,e):
        dlg = wx.FileDialog(self, "Selecciona un archivo", self.dirname, "", "*.*", wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            # Obtengo la ruta y el nombre del archivo para usarlos en button1Click
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            dlg.Destroy()
        

app = wx.App(False)
frame = MainWindow(None, "Converter")
app.MainLoop()