from PIL import Image, ImageDraw, ImageFont
import tempfile
import xlrd
import time
import os
import shutil
from weasyprint import HTML
import PyPDF2
import subprocess

fnWatermarkTemplate ="data/watermark.txt"
fnXLS = "data/datos.xls"
fnHtmlTemplate = "html/certificado.html"
fnFont = "html/font.ttf"
fnTextTemplate = "data/plantilla_marca_agua.txt"

fnWatermarkTemp = "temp/wm.pdf"
fnWatermarkTempPNG = "temp/wm.png"
fnWatermarkedOut = "temp/wmo.pdf"
fnHtmlPdfOut = "temp/htpdo.pdf"
fnHtmlOut = "temp/ht.html"

dFnPDF = {'RF30': 'pdf/rf30s.pdf', 'RF60 SIMPLE': 'pdf/rf60s.pdf', 'RF60 DOBLE': 'pdf/rf60d.pdf', 'RF90': 'pdf/rf120s.pdf', 'RF120': 'pdf/rf120s.pdf', 'RF60 EURO SIMPLE': 'pdf/rf60s_euro_inti.pdf', 'RF60 EURO DOBLE': 'pdf/rf60d_euro.pdf', 'RF60 SIMPLE VIDRIADA': 'pdf/rf60sv_euro.pdf', 'RF120 EURO SIMPLE': 'pdf/rf120s_euro.pdf', 'RF120 EURO DOBLE': 'pdf/rf120d_euro.pdf'}
dOrden = {'RF30': 'Nro 101/14809, con fecha del 25/02/2008', 'RF60 SIMPLE': 'Nro 101/15978, con fecha del 30/08/2011', 'RF60 DOBLE': 'Nro 101/8896, con fecha del 30/03/2005', 'RF90': 'Nro 101/4268, con fecha del 17/10/2000','RF120': 'Nro 101/4268, con fecha del 17/10/2000', 'RF60 EURO SIMPLE': 'Nro 101/22989, con fecha del 10/09/2013', 'RF60 EURO DOBLE': 'Nro 13097-2, con fecha del 14/09/2006', 'RF60 SIMPLE VIDRIADA': 'Nro 28485-2, con fecha del 25/08/2011', 'RF120 EURO SIMPLE': 'Nro 22076-5, con fecha del 19/06/2009', 'RF120 EURO DOBLE': 'Nro 24922-2, con fecha del 16/06/2010'}

def deleteFile(fnIn):
        try:
            os.unlink(fnIn)
        except IOError as e:
            print("Error al eliminar {0}").format(fnIn)
            exit()

def load_dictionaries_xls():

    book = xlrd.open_workbook(fnXLS)
    sh = book.sheet_by_index(0)

    lData = []
    dColumnToKey = {}

    for cx in range(sh.ncols):
        dColumnToKey[cx] = sh.cell_value(rowx=0, colx=cx)

    for rx in range(sh.nrows):
        certificado = {}
        for cx in range(sh.ncols):
            certificado[dColumnToKey[cx]] = sh.cell_value(rowx=rx, colx=cx)

        lData.append(certificado)

    lData.pop(0) #the first entry on the list is the first row of the xls, which is the keys to the dictionary
    return lData

def generateHtmlPdf(dCertificado):
    with open(fnHtmlTemplate, 'r') as fTemplate:
        sOut = fTemplate.read()

    for key in dCertificado.keys():
        sOut = sOut.replace(key, str(dCertificado[key]))

    sOut = sOut.replace("#FECHA#", time.strftime("%d-%m-%Y"))
    sOut = sOut.replace("#ORDEN#", dOrden[dCertificado['#MODELO#']])

    with open(fnHtmlOut, 'w') as fOut:
        fOut.write(sOut)

    HTML(fnHtmlOut).write_pdf(fnHtmlPdfOut)

def generate_watermark_text(dCertificado):
    fTemplate = open(fnTextTemplate, 'r')

    string_out = fTemplate.read()

    for key in dCertificado.keys():
        string_out = string_out.replace(key, str(dCertificado[key]))

    string_out = string_out.replace("#FECHA#", time.strftime("%d-%m-%Y"))

    return string_out.upper()

def text_to_watermark(text):
    myFont = ImageFont.truetype(fnFont, 36)

    bg = Image.new("RGBA", (595,842), (255,255,255,0))
    wm = Image.new("RGBA", (800,800), (255, 255, 255, 0))

    draw = ImageDraw.Draw(wm)

    draw.multiline_text((0,0), text, fill=(255, 0, 0, 200), font=myFont, align='center')
    wm = wm.rotate(45, expand=True, resample=Image.BICUBIC)
    imageBox = wm.getbbox()
    wm = wm.crop(imageBox)

    middle = getPasteMiddleCoord(wm.size, bg.size)

    bg.paste(wm, middle, wm)

    bg.save(fnWatermarkTempPNG, "PNG")

    bg.close()
    wm.close()

    convert = subprocess.Popen(['/usr/bin/convert', fnWatermarkTempPNG, fnWatermarkTemp])
    convert.wait()

def concatenate_pdf(fnPdf1, fnPdf2, fnOut):
    fPdf1 = open(fnPdf1, 'rb')
    fPdf2 = open(fnPdf2, 'rb')

    pdfReader1 = PyPDF2.PdfFileReader(fPdf1)
    pdfReader2 = PyPDF2.PdfFileReader(fPdf2)

    pdfWriter = PyPDF2.PdfFileWriter()

    for pageNum in range(pdfReader1.numPages):
        pageObj = pdfReader1.getPage(pageNum)
        pdfWriter.addPage(pageObj)

    for pageNum in range(pdfReader2.numPages):
        pageObj = pdfReader2.getPage(pageNum)
        pdfWriter.addPage(pageObj)

    out = open(fnOut, 'wb')
    pdfWriter.write(out)

    fPdf1.close()
    fPdf2.close()
    out.close()

def merge_pdf_watermark(fnBg, fnWatermark):
    fBg = open(fnBg, 'rb')
    fWatermark = open(fnWatermark, 'rb')

    pdfReader = PyPDF2.PdfFileReader(fBg)
    pdfWriter = PyPDF2.PdfFileWriter()

    pdfWatermarkReader = PyPDF2.PdfFileReader(fWatermark)
    pobjWatermark = pdfWatermarkReader.getPage(0)

    for pageNum in range(0, pdfReader.numPages):
      pobjCurrent = pdfReader.getPage(pageNum)
      pobjCurrent.mergePage(pobjWatermark)
      pdfWriter.addPage(pobjCurrent)


    with open(fnWatermarkedOut, 'wb') as fOut:
      pdfWriter.write(fOut)

    fBg.close()
    fWatermark.close()

def getPasteMiddleCoord((toPasteX, toPasteY), (backgroundX, backgroundY)):
    middle = (backgroundX/2 - toPasteX/2, backgroundY/2 - toPasteY/2)
    return middle


lDicc = load_dictionaries_xls()

for dCertificado in lDicc:
    os.makedirs('temp')
    generateHtmlPdf(dCertificado)
    watermarkText = generate_watermark_text(dCertificado)
    text_to_watermark(watermarkText)
    merge_pdf_watermark(dFnPDF[dCertificado['#MODELO#']], fnWatermarkTemp)

    fnOut = "out/{}_{}.pdf".format(dCertificado['#EMPRESA#'], dCertificado['#MODELO#'])

    concatenate_pdf(fnWatermarkedOut, fnHtmlPdfOut, fnOut)
    shutil.rmtree("temp")
