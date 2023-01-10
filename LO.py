#-------------------------------------------------------------------------------
# Name:        Licencias v4
# Purpose: Emisión de PDF con Licencias en funcion de un PDF mayor
# con ciertas caracteristicas, al cual, hay que borrarle y agregarle cosas en una
# pagina y agregarle paginas modelo.
#
# Author:      Ing. Esp. Leonardo Barenghi
#
# Created:     14/09/2021
# Modificated:     10/01/2023
# Copyright:   (c) lbarenghi 2021-2023 https://github.com/leogbar/Licencias
# Licence:     BSD 3
#-------------------------------------------------------------------------------

import PyPDF2 as pypdf #https://mstamy2.github.io/PyPDF2/
from unidecode import unidecode  #(c) Tomaz Solc (GPLv2+) (GPL) https://pypi.org/project/Unidecode/ https://github.com/avian2/unidecode
from reportlab.lib.pagesizes import A4 #(c)  Andy Robinson, Robin Becker, the ReportLab team and the community (BSD) https://pypi.org/project/reportlab/
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
import datetime
import PySimpleGUI as sg #(LGPLv3+) PySimpleGUI https://github.com/PySimpleGUI/PySimpleGUI
from sys import platform

if platform == "linux" or platform == "linux2":
    separador = "/"
elif platform == "win32" or platform == "win64":
    separador = "\\"


salida=[]
pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Bold', 'Arial.ttf'))

def current_date_format(date):
    months = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    messsage = "{} de {} del {}".format(day, month, year)
    return messsage

now = datetime.datetime.now()


event, values = sg.Window('Generacion de Licencias v3', [[sg.Text('(c) Ing. Esp. Leonardo Barenghi - 2021 - Licencia bajo BSD 3')],[sg.Text(' ')],[sg.Text('Ingrese cantidad de dias a partir de hoy para la emisión:')],[sg.InputText()],
[sg.Radio('Autorización de Operación Provisoria', "RADIO1", default=False, key="Licencia0")],
[sg.Radio('Licencia de Operación', "RADIO1", default=False, key="Licencia1")],
[sg.Radio('Registro', "RADIO1", default=False, key="Licencia2")],
[sg.Text(' ')],
[sg.Radio('Una pagina por Documento', "RADIO2", default=False, key="cantidad0")],
[sg.Radio('Un Documento con multiples paginas', "RADIO2", default=False, key="cantidad1")],
 [sg.Text('Cargue el Archivo con el/los documento/s de entrada')], [sg.Input(), sg.FileBrowse()],
 [sg.Text('Cargue el Archivo que contiene los nombres de las Entidades Responsables:')], [sg.Input(), sg.FileBrowse()],
[sg.Text('Seleccione carpeta de salida:')],  [sg.Input(),sg.FolderBrowse()], [sg.OK(), sg.Cancel()]

 ]).read(close=True)

encabezadof=""
adjuntof=""

if values["Licencia0"] == True:
    encabezadof = "EncLicProv.pdf"
    adjuntof ="AdjLicProv.pdf"
    titulo = "Aut_Op_Prov"
else:
    if  values["Licencia1"] == True:
            encabezadof ="EncLicOp.pdf"
            adjuntof ="AdjLicOp.pdf"
            titulo = "Lic_Op"
    else:

        if values["Licencia2"] == True:
                encabezadof = "EncReg.pdf"
                adjuntof ="AdjReg.pdf"
                titulo = "Reg"
        else:
            print("Operación invalida")
            exit()


explorar =  values["Browse"]
nombres = values["Browse0"]
canvas = Canvas("firma.pdf", pagesize=A4)
if len(values[0])<1:
    print("No ingresó la cantidad de días...")
    exit()
else:
    dias = int(values[0])

AnchoA4 =A4[0]

FirmaIni=open("firma.ini","r", encoding="utf-8")
NombreFirma=FirmaIni.readline()
PuestoFirma1=FirmaIni.readline()
PuestoFirma2=FirmaIni.readline()


dias=datetime.timedelta(days=dias)
lugar = "BUENOS AIRES, "
fecha = current_date_format(now+dias)
lugarFecha = lugar + fecha
canvas.setFont("Arial", 10)

tamanoTexto = stringWidth(lugarFecha, "Arial", 10)
ubicacionTexto=(AnchoA4*3/4-tamanoTexto)/2
canvas.drawString(ubicacionTexto, 60, lugarFecha )

tamanoTexto = stringWidth(NombreFirma, "Arial", 10)
ubicacionTexto=(AnchoA4*3/4-tamanoTexto)/2
canvas.drawString(ubicacionTexto, 50, NombreFirma[:-1]) #Nombre de quien firma

canvas.setFont("Times-Roman", 8)

tamanoTexto = stringWidth(PuestoFirma1, "Times-Roman", 8)
ubicacionTexto=(AnchoA4*3/4-tamanoTexto)/2
canvas.drawString(ubicacionTexto, 40, PuestoFirma1[:-1]) #Area a la que pertenece

tamanoTexto = stringWidth(PuestoFirma2, "Times-Roman", 8)
ubicacionTexto=(AnchoA4*3/4-tamanoTexto)/2
canvas.drawString(ubicacionTexto, 30, PuestoFirma2[:-1]) #Area a la que pertenece

canvas.save()

encabezado=open(encabezadof, "rb")
firma=open("firma.pdf", "rb")
borrar=open("borrar.pdf", "rb")
linea=open("linea.pdf", "rb")
adjunto=open(adjuntof, "rb")
anexo=open("anexo.pdf", "rb")
anexoBorrar=open("anexoBorrar.pdf", "rb")
carpeta= values["Browse1"]
if len(explorar)<1:
    print("No ingresó el archivo con las Licencias...")
    exit()
else:
    recorrer=open(explorar, "rb")

if len(nombres)<1:
    nombres="Licencia"
else:
    nombres=open(nombres,"r", encoding="utf-8")
if len(carpeta)<1:
    print("No ingresó el archivo con la carpeta de salida...")
    exit()

borrarPdf = pypdf.PdfFileReader(borrar).getPage(0)
encabezadoPdf = pypdf.PdfFileReader(encabezado).getPage(0)
firmaPdf = pypdf.PdfFileReader(firma).getPage(0)
adjuntoPdf = pypdf.PdfFileReader(adjunto).getPage(0)
lineaPdf = pypdf.PdfFileReader(linea).getPage(0)
anexoPdf = pypdf.PdfFileReader(anexo).getPage(0)
anexoBorrarPdf = pypdf.PdfFileReader(anexoBorrar).getPage(0)
original = pypdf.PdfFileReader(recorrer)
x=1

if values["cantidad0"] == True:
    for i in range(original.getNumPages()):
        archivo=nombres.readline()
        licencia1 = pypdf.PdfFileWriter()
        licencia1=original.getPage(i)

        licencia1.mergePage(borrarPdf)
        licencia1.mergePage(encabezadoPdf)
        licencia1.mergePage(lineaPdf)
        licencia1.mergePage(firmaPdf)

        licencia = pypdf.PdfFileWriter()
        licencia.addPage(licencia1)
        licencia.addPage(adjuntoPdf)

        archivo = unidecode(archivo)
        archivo=archivo[:-1]
        salida=carpeta+separador+str(x)+"_"+archivo+".pdf"
        x=x+1
        print("Generando "+salida+ "..."+"\n")
        with open(salida, "wb") as outFile:
            licencia.write(outFile)
    print("Proceso Terminado..."+"\n")
    exit()
else:
    if values["cantidad1"] == True:
        fecha='{:%Y%m%d-%H%M%S}'.format(now)
        archivo=titulo+"_"+fecha
        licencia = pypdf.PdfFileWriter()

        for i in range(original.getNumPages()):
            if(x==1):
                licencia1 = pypdf.PdfFileWriter()
                licencia1=original.getPage(i)
                licencia1.mergePage(borrarPdf)
                licencia1.mergePage(encabezadoPdf)
                licencia1.mergePage(lineaPdf)
                licencia1.mergePage(firmaPdf)
                x=x+1
            else:
                licencia1 = pypdf.PdfFileWriter()
                licencia1=original.getPage(i)
                licencia1.mergePage(anexoBorrarPdf)
                licencia1.mergePage(anexoPdf)
                x=x+1
            licencia.addPage(licencia1)
        licencia.addPage(adjuntoPdf)
        salida=carpeta+separador+archivo+".pdf"
        print("Generando "+salida+ "..."+"\n")

        with open(salida, "wb") as outFile:
            licencia.write(outFile)
        print("Proceso Terminado..."+"\n")
        #exit()
    else:
        print("No ingresó la cantidad de hojas...")
        #exit()

