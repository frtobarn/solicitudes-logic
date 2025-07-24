#!/usr/bin/env python3

import argparse
import os
import shutil
import re
import pandas as pd
import gspread
from google.auth import default
from datetime import datetime
import pytz
import qrcode
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, KeepTogether, Image, PageBreak
from reportlab.lib.units import mm

# Rutas relativas dentro de Colab
RUTA_LOGO = '/content/drive/MyDrive/Domicilios/logos_black.png'
CARPETA_DESTINO_PDF = '/content/drive/MyDrive/Domicilios/Tiquetes'

# Función para limpiar los títulos de materiales
def limpiar_titulo(tokens, truncar=True):
    if not tokens:
        return ""
    lit_primarios = {"A","C","N","P","T","M-L","CA","A-C","A-N","A-P"}
    pat_num = re.compile(r'^\d{1,3}(?:\.\d+)?$')
    pat_gen = re.compile(r'^[A-Z]\d+[A-Za-z]$')
    pat_lit = re.compile(r'^[A-Z0-9]{4,5}$')
    pat_dist = re.compile(r'^[A-Z0-9]{2,}$')

    skip = 0
    if tokens and tokens[0] in lit_primarios:
        skip = 1
        if len(tokens) > 1 and pat_lit.match(tokens[1]):
            skip = 2
    elif tokens and pat_num.match(tokens[0]):
        skip = 1
        if len(tokens) > 1 and pat_gen.match(tokens[1]):
            skip = 2
    elif tokens and tokens[0] == 'DG':
        skip = 1
        if len(tokens) > 1 and pat_dist.match(tokens[1]):
            skip = 2

    palabras_final = {"Primera","Segunda","Tercera","Cuarta","Quinta","Sexta",
                      "Séptima","Octava","Novena","Décima","Edición","edición",
                      "NUEVO","ESTA","EN","COLECCIONES","PRESTADO","CAMBIADO","POR"}

    fin = len(tokens)
    for i in range(skip, len(tokens)):
        if tokens[i] in palabras_final:
            fin = i
            break

    titulo_tokens = [t for t in tokens[skip:fin] if t not in {'-', ':'}]
    titulo = " ".join(titulo_tokens)
    if truncar and len(titulo) > 25:
        titulo = titulo[:25] + '...'
    return titulo

# Función que genera un solo tiquete y lo añade a la lista de elementos
def generar_tiquet(elements, datos, nombre_bib, fecha_fmt, banner):
    qr_data = f"{datos['Cedula']}\t{datos['Nombre']}\t{datos['Direccion']} {datos['Localidad']} {datos['Barrio']}\t{datos['Telefono']}"
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
    qr.add_data(qr_data)
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="black", back_color="white")
    qr_path = f"/content/{datos['Cedula']}_qr.png"
    img_qr.save(qr_path)

    logo = Image(RUTA_LOGO)
    scale = 1.6
    logo.drawHeight = (4.5 * mm) * scale
    logo.drawWidth = (36.2 * mm) * scale
    img = Image(qr_path)
    img.drawHeight = 75
    img.drawWidth = 75

    tabla = [
        [logo, "", "Servicio de préstamo a domicilio", "", ""],
        ["Datos\nde\norigen", "Fecha de alistamiento", fecha_fmt, "", img],
        ["", "Biblioteca que envía", nombre_bib, "", ""],
        ["", "Biblioteca que recibe", datos['Biblioteca'], "", ""],
        ["Datos del\nusuario\nsolicitante", "Nombre", datos['Nombre'], "", ""],
        ["", "Dirección", datos['Direccion'], "", ""],
        ["", "Barrio", datos['Barrio'], "Localidad", datos['Localidad']],
        ["", "No. de documento", datos['Cedula'], "No. de teléfono", datos['Telefono']],
        ["Total\nsolicitud", "Materiales", datos['Materiales'], "Cantidad", datos['Cantidad']],
        ["", banner, "", "", ""]
    ]
    col_w = [2.1*28.35, 4.1*28.35, 6*28.35, 3*28.35, 5*28.35]
    table = Table(tabla, colWidths=col_w)

    style = TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('WORDWRAP',(0,0),(-1,-1)),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('FONT',(0,0),(-1,-1),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),10),
        ('SPAN',(0,0),(1,0)),('ALIGN',(0,0),(2,0),'CENTER'),
        ('SPAN',(2,0),(4,0)),('ALIGN',(2,0),(4,0),'CENTER'),('BACKGROUND',(2,0),(4,0),colors.grey),('TEXTCOLOR',(2,0),(4,0),colors.white),
        ('SPAN',(0,1),(0,3)),('ALIGN',(0,1),(0,3),'LEFT'),('BACKGROUND',(0,1),(0,3),colors.grey),('TEXTCOLOR',(0,1),(0,3),colors.white),
        ('SPAN',(4,1),(4,4)),('ALIGN',(4,1),(4,4),'CENTER'),
        ('SPAN',(2,1),(3,1)),('ALIGN',(2,1),(3,1),'CENTER'),
        ('SPAN',(2,2),(3,2)),('ALIGN',(2,2),(3,2),'CENTER'),
        ('SPAN',(2,3),(3,3)),('ALIGN',(2,3),(3,3),'CENTER'),
        ('SPAN',(0,4),(0,7)),('ALIGN',(0,4),(0,7),'LEFT'),('BACKGROUND',(0,4),(0,7),colors.grey),('TEXTCOLOR',(0,4),(0,7),colors.white),
        ('SPAN',(2,4),(3,4)),('ALIGN',(2,4),(3,4),'CENTER'),
        ('SPAN',(2,5),(4,5)),('ALIGN',(2,5),(4,5),'CENTER'),
        ('BACKGROUND',(0,8),(0,9),colors.grey),('TEXTCOLOR',(0,8),(0,9),colors.white),
        ('SPAN',(1,9),(4,9)),
        ('ALIGN',(4,8),(4,8),'CENTER'),('FONT',(4,8),(4,8),'Helvetica-Bold'),('FONTSIZE',(4,8),(4,8),14),('VALIGN',(4,8),(4,8),'MIDDLE'),
        ('FONT',(2,8),(2,8),'Helvetica'),('FONTSIZE',(2,8),(2,8),8),('ALIGN',(2,8),(2,8),'CENTER'),('TEXTCOLOR',(2,8),(2,8),colors.darkviolet),('WORDWRAP',(2,8),(2,8)),
        ('TOPPADDING',(1,9),(4,9),5),('BOTTOMPADDING',(1,9),(4,9),5)
    ])
    table.setStyle(style)
    elements.append(KeepTogether(table))

# Función principal
def main():
    parser = argparse.ArgumentParser(description='Generar tiquetes PDF.')
    parser.add_argument('--inicio',    type=int,   required=True, help='Registro inicial (fila)')
    parser.add_argument('--fin',       type=int,   required=True, help='Registro final (fila)')
    parser.add_argument('--nombre',    type=str,   required=True, help='Nombre de la biblioteca')
    parser.add_argument('--hoja_id',   type=str,   required=True, help='ID de la hoja de Google Sheets')
    parser.add_argument('--banner',    type=str,   required=True, help='Texto de banner o novedades')
    args = parser.parse_args()

    tz = pytz.timezone('America/Bogota')
    now = datetime.now(tz)
    dias = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
    meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    dia_sem = dias[now.weekday()]
    fecha_fmt = f"{dia_sem} {now.day} de {meses[now.month-1]} de {now.year} a las {now.strftime('%I').lstrip('0')}:{now.strftime('%M')} {now.strftime('%p').lower()}"
    fname = f"{now.year}_{meses[now.month-1]}_{now.day}_{dia_sem}_{now.strftime('%I').lstrip('0')}_{now.strftime('%M')}_{now.strftime('%p').lower()}"

    creds, _ = default()
    gc = gspread.authorize(creds)
    sheet = gc.open_by_key(args.hoja_id).sheet1
    vals = sheet.get_all_values()
    df = pd.DataFrame(vals[args.inicio-1:args.fin], columns=vals[1])

    resultados = []
    for _, row in df.iterrows():
        materiales = ""
        count = 0
        for i in range(1, 10):
            key = f'MATERIAL {i}'
            if key in row and row[key].strip():
                count += 1
                materiales += limpiar_titulo(row[key].split(), True) + ' | '
                if i % 2 == 1:
                    materiales += '\n'
        resultados.append({
            'Nombre':     row['NOMBRE DE USUARIO'],
            'Cedula':     row['N° Identificación'],
            'Direccion':  row['DIRECCIÓN'],
            'Telefono':   row['N° Telefono'],
            'Localidad':  row['LOCALIDAD'],
            'Barrio':     row['BARRIO'],
            'Biblioteca': row.get('MATERIAL 9',''),
            'Materiales': materiales.strip('\n '),
            'Cantidad':   count
        })

    df_res = pd.DataFrame(resultados)

    doc = SimpleDocTemplate(f"{fname}_tiquetes.pdf", pagesize=letter,
                            leftMargin=0.2*28.35, rightMargin=0.2*28.35,
                            topMargin=0.2*28.35, bottomMargin=0.2*28.35)
    elements = []
    for idx, data in enumerate(df_res.iterrows()):
        generar_tiquet(elements, data[1], args.nombre, fecha_fmt, args.banner)
        # Insertar salto de página después de cada 4 tiquetes
        if (idx + 1) % 4 == 0:
            elements.append(PageBreak())

    doc.build(elements)

    generado = f"/content/{fname}_tiquetes.pdf"
    shutil.copy(generado, CARPETA_DESTINO_PDF)
    print("✅ Proceso completado. Archivo generado:", generado)

if __name__ == '__main__':
    main()
