# -*- coding: utf-8 -*-
from datetime import date, time
from decimal import Decimal
from zipfile import ZipFile
from excel import generar_excel
import os
import wx
import sqlite3


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
        r = [record[0]
             for record in conn.execute("select cuenta from empresas where abreviatura='%s'" % self.abreviatura)]
        conn.close()
        if r:
            self.cuenta = r[0]
        else:
            self.cuenta = ""
    
    def obtener_razon_social(self, *args, **kwargs):
        conn = sqlite3.connect(self.database)
        r = [record[0]
             for record in conn.execute("select razon_social from empresas where abreviatura='%s'" % self.abreviatura)]
        conn.close()
        if r:
            self.razon_social = r[0]
        else:
            self.cuenta = ""
        
    def todas(self, *args, **kwargs):
        conn = sqlite3.connect(self.database)
        try:
            r = [record[0]
                 for record in conn.execute("select abreviatura from empresas")]
        except:
            return ""
        conn.close()
        return r

    
def procesar_cdpg(cdpgfile, empresa_alias):
    cabecera = []
    detalles = []
    for i, line in enumerate(cdpgfile):
        if i == 0:
            # Cabecera
            cuenta = Empresas(empresa_alias).cuenta
            cuenta_txt = line[2:5] + "-" + line[6:13]
            if cuenta_txt != cuenta[:11]:
                return ["Error", "Cuenta %s no pertenece a %s" % (cuenta_txt, empresa_alias)]
            
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
            
            detalles.extend([[codigo, depositante, retorno, dato_adicional,
                              fecha_pago, fecha_vencimiento, monto_pagado,
                              mora_pagado, monto_total, sucursal, num_operacion,
                              referencia, terminal_id, atencion_medio,
                              atencion_hora, num_cheque, cod_banco]])
    if len(detalles) == 0:
        return ["Error", "No hay datos"]
    else:
        return [cabecera, detalles]
    
    
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
        wx.Frame.__init__(self, parent, title=title, size=(310, 200))
        self.CreateStatusBar()
        
        # ComboBox
        empresaD = {}
        empresas = Empresas()
        for empresa in empresas.todas():
            empresaD[empresa] = empresa
        
        empresaList = sorted(empresaD.keys())
        
        wx.StaticText(self, -1, "Seleccione una empresa: ", (10, 10))
        self.combo1 = wx.ComboBox(self, -1, value="", pos=wx.Point(10, 30),
                                  size=wx.Size(150, 28), choices=empresaList)
        
        # Menu
        filemenu = wx.Menu()
        menuOpen = filemenu.Append(wx.ID_OPEN, "&Abrir", " Abrir un archivo")
        menuExit = filemenu.Append(wx.ID_EXIT, "&Salir", " Finaliza el programa")
        #
        empmenu = wx.Menu()
        menuAdmEmp = empmenu.Append(wx.ID_HOME, "&Administrar", " Administrar Empresas")
        #
        helpmenu = wx.Menu()
        menuAbout = helpmenu.Append(wx.ID_ABOUT, "A&cerca de", " Información sobre el programa")
        
        # MenuBar
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, "&Archivo")
        menuBar.Append(empmenu, "&Empresas")
        menuBar.Append(helpmenu, "&Ayuda")
        self.SetMenuBar(menuBar)
        
        # Button
        self.button1 = wx.Button(self, -1, label="Convertir",
                                 pos=wx.Point(170, 30), size=wx.Size(130, 28))
        
        # FileDisplay
        self.FileDisplay = wx.StaticText(self, wx.ID_ANY, label="Debe seleccionar un archivo (zip o txt)", pos=wx.Point(10, 90))
        
        # Set events.
        self.Bind(wx.EVT_MENU, self.OnOpen, menuOpen)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnEmp, menuAdmEmp)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        
        self.button1.Bind(wx.EVT_BUTTON, self.button1Click, self.button1)
        
        self.Show(True)
        
    def button1Click(self, e):
        self.empresa_alias = self.combo1.GetValue()
        try:
            f = abrir_archivo(self.dirname, self.filename)
        except:
            f = None
            datos = ["Error", "No ha seleccionado el archivo a convertir"]
        
        if f:
            datos = procesar_cdpg(f, self.empresa_alias)
            f.close()
        
        if datos[0] == "Error":
            dlg = wx.MessageDialog(self, "Error: " + datos[1], "Mensaje", wx.ICON_HAND)
        else:
            generar_excel(self.dirname, datos)
            dlg = wx.MessageDialog(self, "Archivo generado correctamente", "Mensaje", wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()

    def OnAbout(self, e):
        dlg = wx.MessageDialog(self, "Convierte CDPG del banco a Excel \n\n\t V0.3",
                               "PyCDPGConverter", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()

    def OnExit(self, e):
        self.Close(True)
        
    def OnOpen(self, e):
        dlg = wx.FileDialog(self, "Selecciona un archivo (zip o txt)", self.dirname, "", "*.*", wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            # Obtengo la ruta y el nombre del archivo para usarlos en button1Click
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            dlg.Destroy()
            self.FileDisplay.SetLabel("Archivo: %s" % self.filename)
            
    def OnEmp(self, e):
        secondWindow = Window2(None, "Administrar")
        secondWindow.Show()
        

class Window2(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, None, wx.ID_ANY, title="Administrar Empresas", size=(500, 200))
        panel = wx.Panel(self, wx.ID_ANY)
        self.index = 0
        
        self.list_ctrl = wx.ListCtrl(panel, size=(390, 100),
                                     style=wx.LC_REPORT | wx.BORDER_SUNKEN)
        self.list_ctrl.InsertColumn(0, 'Cuenta', width=120)
        self.list_ctrl.InsertColumn(1, 'Abreviatura', width=120)
        self.list_ctrl.InsertColumn(2, 'Razón Social', width=230)
        
        empresaD = {}
        empresas = Empresas()
        for empresa in empresas.todas():
            empresaD[empresa] = empresa
        
        empresaList = sorted(empresaD.keys())
        items = empresaList
        index = 0
        for key, empresa_alias in enumerate(items):
            cuenta = Empresas(empresa_alias).cuenta
            razon_social = Empresas(empresa_alias).razon_social
            data = [cuenta, empresa_alias, razon_social]
            self.list_ctrl.InsertStringItem(index, data[0])
            self.list_ctrl.SetStringItem(index, 1, data[1])
            self.list_ctrl.SetStringItem(index, 2, data[2])
            self.list_ctrl.SetItemData(index, key)
            index += 1
        
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(self.list_ctrl, 0, wx.ALL | wx.EXPAND, 5)
        btnSizer = wx.BoxSizer(wx.VERTICAL)
        btn = wx.Button(panel, label="Nueva")
        btn2 = wx.Button(panel, label="Eliminar")
        
        btnSizer.Add(btn, 0, wx.ALL | wx.UP, 5)
        btnSizer.Add(btn2, 0, wx.ALL | wx.UP, 5)
        
        sizer.Add(btnSizer, 0, wx.CENTER|wx.ALL, 5)
        panel.SetSizer(sizer)

app = wx.App(False)
frame = MainWindow(None, "Converter")
app.MainLoop()