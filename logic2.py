#!/usr/bin/env python3
import sys
import os
import argparse
import pandas as pd
import shutil
import gspread
import qrcode
import random
from datetime import datetime
import pytz
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, KeepTogether, Image
)
from google.auth import default
import re

# -- funciones auxiliares -------------------

def limpiar_titulo(lista, truncar):
    # (copia aquí tu función limpiar_titulo tal cual)
    # …
    return titulo_limpio

def generar_tiquet(pdf_elements, datos, img, fecha_str, banner):
    # (copia aquí tu código de generación de tiquet,
    # recibe: lista pdf_elements, fila de df, objeto Image img, fecha_formateada, banner)
    # …
    pdf_elements.append(KeepTogether(table))

def main(args):
    # 1) Montaje asumido hecho por Colab
    # 2) Preparar credenciales de gspread
    creds, _ = default()
    gc = gspread.authorize(creds)

    # 3) Variables desde CLI
    registro_inicial  = args.registro_inicial
    registro_final    = args.registro_final
    nombre_biblioteca = args.nombre_biblioteca
    id_hoja           = args.id_hoja
    banner            = args.banner

    # 4) Rango y DataFrame
    hoja   = gc.open_by_key(id_hoja).sheet1
    datos  = hoja.get_all_values()
    df     = pd.DataFrame(
        datos[registro_inicial-1:registro_final],
        columns=datos[1]
    )

    # 5) Copiar logo y preparar imagen
    ruta_logo = '/content/drive/MyDrive/Domicilios/logos_black.png'
    shutil.copy(ruta_logo, '/content/logos_black.png')
    img = Image("/content/logos_black.png")
    scale = 1.6
    img.drawHeight = (4.5 * mm) * scale
    img.drawWidth  = (36.2 * mm) * scale

    # 6) Prepara fecha y nombre de PDF
    tz     = pytz.timezone("America/Bogota")
    now    = datetime.now(tz)
    dias   = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
    meses  = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto",
              "septiembre","octubre","noviembre","diciembre"]
    dia_nom = dias[now.weekday()]
    fecha_str = f"{dia_nom} {now.day} de {meses[now.month-1]} de {now.year} a las {now.strftime('%I').lstrip('0')}:{now.strftime('%M')} {now.strftime('%p').lower()}"
    archivo_pdf = now.strftime(f"%Y_%m_%d_{dia_nom}_%I_%M_%p_tiquetes.pdf")

    # 7) Procesar filas
    resultados = []
    for _, fila in df.iterrows():
        # (corta la parte de generar lista de materiales y demás)
        resultados.append({
            'Nombre':       fila['NOMBRE DE USUARIO'],
            'Cedula':       fila['N° Identificación'],
            'Direccion':    fila['DIRECCIÓN'],
            'Localidad':    fila['LOCALIDAD'],
            'Barrio':       fila['BARRIO'],
            'Telefono':     fila['N° Telefono'],
            'Biblioteca':   fila['MATERIAL 9'],
            'Materiales':   materiales,
            'Cantidad':     cantidad
        })
    df_res = pd.DataFrame(resultados)

    # 8) Crear PDF
    doc = SimpleDocTemplate(archivo_pdf, pagesize=letter,
                            leftMargin=0.2*28.35, rightMargin=0.2*28.35,
                            topMargin=0.2*28.35, bottomMargin=0.2*28.35)
    elements = []
    for _, row in df_res.iterrows():
        generar_tiquet(elements, row, img, fecha_str, banner)
    doc.build(elements)

    # 9) Copiar a Drive
    dest = '/content/drive/MyDrive/Domicilios/Tiquetes/' + archivo_pdf
    shutil.copy(archivo_pdf, dest)
    print("✅ PDF guardado en:", dest)

if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("registro_inicial",  type=int,  help="Primer registro")
    p.add_argument("registro_final",    type=int,  help="Último registro")
    p.add_argument("nombre_biblioteca",   help="Nombre de la biblioteca")
    p.add_argument("id_hoja",             help="Google Sheet ID")
    p.add_argument("banner",              help="Texto de banner")
    args = p.parse_args()
    main(args)
