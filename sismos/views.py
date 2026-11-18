import base64
import json
import os
import sqlite3
import sys
from io import BytesIO
from xml.sax.saxutils import escape

import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from django.conf import settings
from django.http import FileResponse, JsonResponse
from django.shortcuts import render

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle


# carga desde SQLite (se carga automaticamente desde el CSV la primera vez)
def _cargar_datos_sqlite():
    columnas = ["time", "latitude", "longitude", "depth", "mag", "place", "mag_category", "type"]
    if any(cmd in sys.argv for cmd in ["makemigrations", "migrate"]):
        return pd.DataFrame(columns=columnas + ["year"])

    ruta_db = str(settings.DATABASES["default"]["NAME"])
    ruta_csv = os.path.join(settings.BASE_DIR, "data", "earthquakes_1900_2026.csv")

    with sqlite3.connect(ruta_db) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='sismos_sismo';"
        )
        tabla_existe = cursor.fetchone() is not None
        cantidad = 0
        if tabla_existe:
            cursor.execute("SELECT COUNT(1) FROM sismos_sismo;")
            cantidad = cursor.fetchone()[0]

        if cantidad == 0 and os.path.exists(ruta_csv):
            primer_bloque = True
            for chunk in pd.read_csv(ruta_csv, usecols=columnas, chunksize=150000):
                fechas = pd.to_datetime(chunk["time"], format="ISO8601", errors="coerce")
                chunk["year"] = fechas.dt.year
                chunk.to_sql(
                    "sismos_sismo",
                    conn,
                    if_exists="replace" if primer_bloque else "append",
                    index=False,
                )
                primer_bloque = False

        datos_db = pd.read_sql_query(
            "SELECT time, latitude, longitude, depth, mag, place, mag_category, type, year FROM sismos_sismo",
            conn,
        )

    datos_db["time"] = pd.to_datetime(datos_db["time"], format="ISO8601", errors="coerce")
    if "year" not in datos_db.columns or datos_db["year"].isna().all():
        datos_db["year"] = datos_db["time"].dt.year
    return datos_db


df = _cargar_datos_sqlite()
# Subconjunto pre-limpiado en memoria para que el filtrado en tiempo real tenga un tiempo de respuesta de bajo ms
df_dinamico = df.dropna(subset=["year", "mag", "depth"]).copy()
if not df_dinamico.empty:
    df_dinamico["year"] = df_dinamico["year"].astype(int)
    df_dinamico["place_str"] = df_dinamico["place"].fillna("Sin especificar").astype(str)


def index(request):
    return render(request, "sismos/index.html")


# Funciones aux para exportar tablas en PDF y Excel

def _generar_pdf_reporte(titulo, explicacion, encabezados, filas_datos, nombre_archivo):
    buffer = BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A3),
        rightMargin=24,
        leftMargin=24,
        topMargin=24,
        bottomMargin=24,
        title=titulo,
    )
    estilos = getSampleStyleSheet()
    celda = ParagraphStyle(
        "Celda", fontName="Helvetica", fontSize=9, leading=12, wordWrap="CJK"
    )
    celda_header = ParagraphStyle(
        "CeldaHeader", fontName="Helvetica-Bold", fontSize=9, leading=12, wordWrap="CJK"
    )

    filas = [[Paragraph(escape(str(texto)), celda_header) for texto in encabezados]]
    for fila in filas_datos:
        filas.append([Paragraph(escape(str(valor)), celda) for valor in fila])

    if not filas_datos:
        filas.append(
            [Paragraph("No existen registros para los criterios seleccionados.", celda)]
            + [""] * (len(encabezados) - 1)
        )

    ancho_util = 1140
    anchos = [int(ancho_util / len(encabezados))] * len(encabezados)
    tabla = LongTable(filas, colWidths=anchos, repeatRows=1, hAlign="LEFT")

    formato = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE6F1")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BBBBBB")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]
    if not filas_datos:
        formato.append(("SPAN", (0, 1), (-1, 1)))

    tabla.setStyle(TableStyle(formato))
    elementos = [
        Paragraph(escape(titulo), estilos["Title"]),
        Paragraph(escape(explicacion), estilos["Normal"]),
        Spacer(1, 8),
        Paragraph(f"Total de registros en el reporte: {len(filas_datos)}", estilos["Normal"]),
        Spacer(1, 12),
        tabla,
    ]
    documento.build(elementos)
    buffer.seek(0)
    return FileResponse(buffer, as_attachment=True, filename=nombre_archivo)


def _generar_excel_reporte(titulo_hoja, encabezados, filas_datos, nombre_archivo):
    libro = Workbook()
    hoja = libro.active
    hoja.title = titulo_hoja[:31]

    hoja.append(encabezados)
    for fila in filas_datos:
        hoja.append(list(fila))
        for celda in hoja[hoja.max_row]:
            if isinstance(celda.value, str):
                celda.data_type = "s"

    borde = Side(style="thin", color="D9D9D9")
    for fila in hoja:
        for celda in fila:
            celda.border = Border(left=borde, right=borde, top=borde, bottom=borde)
            celda.alignment = Alignment(vertical="top", wrap_text=True)

    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")

    for numero in range(1, len(encabezados) + 1):
        hoja.column_dimensions[get_column_letter(numero)].width = 25

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions

    buffer = BytesIO()
    libro.save(buffer)
    buffer.seek(0)
    return FileResponse(buffer, as_attachment=True, filename=nombre_archivo)



# Reportes 1, 2, 3, y descargas PDF y Excel

def _datos_listado1():
    top_df = df.dropna(subset=["time"]).nlargest(10, "mag").copy()
    top_df["time"] = top_df["time"].dt.strftime("%Y-%m-%d %H:%M")
    top_df["depth"] = top_df["depth"].round(1)
    top_df["mag"] = top_df["mag"].round(2)
    explicacion = (
        "Top 10 sismos de mayor magnitud registrados históricamente a nivel global, "
        "excluyendo registros con fechas incompletas."
    )
    return top_df, "Mayores Sismos Históricos", explicacion


def listado1(request):
    top_df, titulo, explicacion = _datos_listado1()
    return render(
        request,
        "sismos/listado.html",
        {
            "datos": top_df.to_dict("records"),
            "explicacion": explicacion,
            "titulo": titulo,
            "url_pdf": "listado1_pdf",
            "url_excel": "listado1_excel",
        },
    )


def listado1_pdf(request):
    top_df, titulo, explicacion = _datos_listado1()
    encabezados = ["Fecha/Hora", "Ubicación", "Magnitud", "Profundidad (km)"]
    filas = [
        [r["time"], r["place"], r["mag"], f"{r['depth']} km"]
        for r in top_df.to_dict("records")
    ]
    return _generar_pdf_reporte(titulo, explicacion, encabezados, filas, "reporte_sismos_globales.pdf")


def listado1_excel(request):
    top_df, _, _ = _datos_listado1()
    encabezados = ["Fecha/Hora", "Ubicación", "Magnitud", "Profundidad (km)"]
    filas = [
        [str(r["time"]), str(r["place"]), float(r["mag"]), float(r["depth"])]
        for r in top_df.to_dict("records")
    ]
    return _generar_excel_reporte("Sismos Globales", encabezados, filas, "reporte_sismos_globales.xlsx")


def _datos_listado2():
    chile_df = df[df["place"].str.contains("Chile", case=False, na=False)].dropna(subset=["time"])
    top_chile = chile_df.nlargest(10, "mag").copy()
    top_chile["time"] = top_chile["time"].dt.strftime("%Y-%m-%d %H:%M")
    top_chile["depth"] = top_chile["depth"].round(1)
    top_chile["mag"] = top_chile["mag"].round(2)
    explicacion = (
        "Los 10 eventos sísmicos de mayor magnitud registrados en territorio chileno, "
        "evidenciando la alta liberación de energía en la zona de subducción de la Placa de Nazca."
    )
    return top_chile, "Sismos de Mayor Magnitud en Chile", explicacion


def listado2(request):
    top_chile, titulo, explicacion = _datos_listado2()
    return render(
        request,
        "sismos/listado.html",
        {
            "datos": top_chile.to_dict("records"),
            "explicacion": explicacion,
            "titulo": titulo,
            "url_pdf": "listado2_pdf",
            "url_excel": "listado2_excel",
        },
    )


def listado2_pdf(request):
    top_chile, titulo, explicacion = _datos_listado2()
    encabezados = ["Fecha/Hora", "Ubicación", "Magnitud", "Profundidad (km)"]
    filas = [
        [r["time"], r["place"], r["mag"], f"{r['depth']} km"]
        for r in top_chile.to_dict("records")
    ]
    return _generar_pdf_reporte(titulo, explicacion, encabezados, filas, "reporte_sismos_chile.pdf")


def listado2_excel(request):
    top_chile, _, _ = _datos_listado2()
    encabezados = ["Fecha/Hora", "Ubicación", "Magnitud", "Profundidad (km)"]
    filas = [
        [str(r["time"]), str(r["place"]), float(r["mag"]), float(r["depth"])]
        for r in top_chile.to_dict("records")
    ]
    return _generar_excel_reporte("Sismos Chile", encabezados, filas, "reporte_sismos_chile.xlsx")


def _datos_listado3():
    eventos_naturales = ["earthquake", "Ice Quake", "volcanic eruption", "landSlide", "Rock Slide"]
    antropicos_df = df[~df["type"].isin(eventos_naturales)].dropna(subset=["time"])
    top_antropicos = antropicos_df.nlargest(10, "mag").copy()
    top_antropicos["time"] = top_antropicos["time"].dt.strftime("%Y-%m-%d %H:%M")
    top_antropicos["depth"] = top_antropicos["depth"].round(1)
    top_antropicos["mag"] = top_antropicos["mag"].round(2)
    explicacion = (
        "Detalle de los 10 principales eventos sísmicos de origen artificial "
        "(pruebas nucleares y explosiones mineras), caracterizados por ocurrir a niveles muy superficiales."
    )
    return top_antropicos, "Mayores Eventos Antrópicos", explicacion


def listado3(request):
    top_antropicos, titulo, explicacion = _datos_listado3()
    return render(
        request,
        "sismos/listado.html",
        {
            "datos": top_antropicos.to_dict("records"),
            "explicacion": explicacion,
            "titulo": titulo,
            "url_pdf": "listado3_pdf",
            "url_excel": "listado3_excel",
        },
    )


def listado3_pdf(request):
    top_antropicos, titulo, explicacion = _datos_listado3()
    encabezados = ["Fecha/Hora", "Ubicación", "Magnitud", "Profundidad (km)"]
    filas = [
        [r["time"], r["place"], r["mag"], f"{r['depth']} km"]
        for r in top_antropicos.to_dict("records")
    ]
    return _generar_pdf_reporte(titulo, explicacion, encabezados, filas, "reporte_eventos_antropicos.pdf")


def listado3_excel(request):
    top_antropicos, _, _ = _datos_listado3()
    encabezados = ["Fecha/Hora", "Ubicación", "Magnitud", "Profundidad (km)"]
    filas = [
        [str(r["time"]), str(r["place"]), float(r["mag"]), float(r["depth"])]
        for r in top_antropicos.to_dict("records")
    ]
    return _generar_excel_reporte("Eventos Antropicos", encabezados, filas, "reporte_eventos_antropicos.xlsx")



# Logica para grafico 4 (dinamico) y reporte 4 (dinamico)

REGIONES_DISPONIBLES = [
    ("Chile", "Chile (Subducción de Nazca)"),
    ("Global", "Global (Catálogo Mundial)"),
    ("Japan", "Japón (Fosa del Japón)"),
    ("Indonesia", "Indonesia (Arco de Sunda)"),
    ("Alaska", "Alaska (Fosa Aleutiana)"),
    ("California", "California (Falla de San Andrés)"),
    ("Mexico", "México (Trinchera Mesoamericana)"),
    ("Peru", "Perú (Placa de Nazca)"),
]

PROFUNDIDADES_DISPONIBLES = [
    ("todas", "Todas las profundidades (0 - 700 km)"),
    ("superficial", "Cortical / Superficial (< 70 km)"),
    ("intermedia", "Subducción Intermedia (70 - 300 km)"),
    ("profunda", "Foco Profundo (> 300 km)"),
]


def _obtener_datos_dinamicos(request):
    region = request.GET.get("region", "Chile")
    if region not in dict(REGIONES_DISPONIBLES):
        region = "Chile"

    profundidad = request.GET.get("profundidad", "todas")
    if profundidad not in dict(PROFUNDIDADES_DISPONIBLES):
        profundidad = "todas"

    try:
        mag_min = round(float(request.GET.get("mag_min", 5.5)), 1)
    except ValueError:
        mag_min = 5.5

    try:
        anio_inicio = int(request.GET.get("anio_inicio", 1960))
    except ValueError:
        anio_inicio = 1960

    try:
        anio_fin = int(request.GET.get("anio_fin", 2026))
    except ValueError:
        anio_fin = 2026

    # Detectar rango invertido para invertir eje x en grafico 4
    orden_invertido = anio_inicio > anio_fin
    anio_min = min(anio_inicio, anio_fin)
    anio_max = max(anio_inicio, anio_fin)

    # Filtrado por rango de año
    mascara = (
        (df_dinamico["mag"] >= mag_min)
        & (df_dinamico["year"] >= anio_min)
        & (df_dinamico["year"] <= anio_max)
    )
    if profundidad == "superficial":
        mascara &= df_dinamico["depth"] < 70
    elif profundidad == "intermedia":
        mascara &= (df_dinamico["depth"] >= 70) & (df_dinamico["depth"] <= 300)
    elif profundidad == "profunda":
        mascara &= df_dinamico["depth"] > 300

    sub = df_dinamico[mascara]
    if region != "Global" and not sub.empty:
        sub = sub[sub["place_str"].str.contains(region, case=False, na=False)]

    etiqueta_region = dict(REGIONES_DISPONIBLES)[region]
    etiqueta_prof = dict(PROFUNDIDADES_DISPONIBLES)[profundidad]

    if sub.empty:
        resumen_registros = []
        kpis = {
            "total_sismos": 0,
            "anios_activos": 0,
            "mag_max_historica": "—",
            "anio_max_mag": "—",
            "lugar_max_historico": "Sin registros",
            "prof_media_global": "—",
            "promedio_anual": 0,
        }
        interpretacion = (
            f"No existen eventos sísmicos en la base de datos para {etiqueta_region} en el período "
            f"{anio_inicio}–{anio_fin} con magnitud ≥ {mag_min} Mw y régimen '{etiqueta_prof}'."
        )
    else:
        idx_max = sub.groupby("year")["mag"].idxmax()
        lugares_max = sub.loc[idx_max, ["year", "place_str"]].set_index("year")["place_str"].to_dict()

        agrupado = (
            sub.groupby("year")
            .agg(
                cantidad=("mag", "count"),
                mag_promedio=("mag", "mean"),
                mag_max=("mag", "max"),
                depth_promedio=("depth", "mean"),
            )
            .reset_index()
            .sort_values("year", ascending=not orden_invertido)
        )

        resumen_registros = [
            {
                "year": int(fila["year"]),
                "cantidad": int(fila["cantidad"]),
                "mag_promedio": round(float(fila["mag_promedio"]), 2),
                "mag_max": round(float(fila["mag_max"]), 2),
                "depth_promedio": round(float(fila["depth_promedio"]), 1),
                "lugar_max": str(lugares_max.get(int(fila["year"]), "Sin especificar")),
            }
            for _, fila in agrupado.iterrows()
        ]

        total_eventos = int(agrupado["cantidad"].sum())
        fila_pico = agrupado.loc[agrupado["cantidad"].idxmax()]
        fila_mayor_mag = agrupado.loc[agrupado["mag_max"].idxmax()]
        anio_mayor = int(fila_mayor_mag["year"])
        lugar_mayor = str(lugares_max.get(anio_mayor, "Sin especificar"))
        prof_global = round(float(sub["depth"].mean()), 1)
        prom_anual = round(total_eventos / len(resumen_registros), 1)

        kpis = {
            "total_sismos": total_eventos,
            "anios_activos": len(resumen_registros),
            "mag_max_historica": round(float(fila_mayor_mag["mag_max"]), 2),
            "anio_max_mag": anio_mayor,
            "lugar_max_historico": lugar_mayor,
            "prof_media_global": prof_global,
            "promedio_anual": prom_anual,
        }

        sentido_txt = "orden cronológico inverso" if orden_invertido else "orden cronológico ascendente"
        interpretacion = (
            f"Interpretación sísmica ({sentido_txt} {anio_inicio}→{anio_fin}): Se registran {total_eventos} eventos "
            f"en {len(resumen_registros)} años activos (promedio de {prom_anual} sismos/año y profundidad media de {prof_global} km). "
            f"El pico de frecuencia ocurrió en {int(fila_pico['year'])} con {int(fila_pico['cantidad'])} sismos, "
            f"mientras que la mayor liberación de energía alcanzó {kpis['mag_max_historica']} Mw en {anio_mayor} ({lugar_mayor})."
        )

    explicacion = (
        f"Origen de datos: Base de datos SQLite (tabla sismos_sismo, catálogo USGS 1900-2026). "
        f"Parámetros activos: Zona = {etiqueta_region} | Régimen = {etiqueta_prof} | "
        f"Magnitud mínima ≥ {mag_min} Mw | Período = {anio_inicio} → {anio_fin}. "
        f"Cálculos por año: Frecuencia anual, magnitud media, magnitud máxima, profundidad hipocentral media y epicentro del evento principal."
    )

    filtros = {
        "region": region,
        "profundidad": profundidad,
        "mag_min": mag_min,
        "anio_inicio": anio_inicio,
        "anio_fin": anio_fin,
        "invertido": orden_invertido,
        "etiqueta_region": etiqueta_region,
        "etiqueta_prof": etiqueta_prof,
    }
    return resumen_registros, filtros, explicacion, interpretacion, kpis


# Reporte 4 sincronizado con visualizacion de grafico 4

def listado4(request):
    registros, filtros, explicacion, interpretacion, kpis = _obtener_datos_dinamicos(request)
    if request.GET.get("ajax") == "1":
        return JsonResponse(
            {
                "datos": registros,
                "filtros": filtros,
                "explicacion": explicacion,
                "interpretacion": interpretacion,
                "kpis": kpis,
            }
        )

    return render(
        request,
        "sismos/listado.html",
        {
            "titulo": "Reporte Dinámico: Perfil Anual de Severidad y Profundidad Sísmica",
            "explicacion": explicacion,
            "interpretacion": interpretacion,
            "datos": registros,
            "kpis": kpis,
            "es_reporte_dinamico": True,
            "filtros": filtros,
            "regiones": REGIONES_DISPONIBLES,
            "profundidades": PROFUNDIDADES_DISPONIBLES,
            "url_pdf": "listado4_pdf",
            "url_excel": "listado4_excel",
        },
    )


def listado4_pdf(request):
    registros, _, explicacion, interpretacion, _ = _obtener_datos_dinamicos(request)
    titulo = "Reporte Dinámico: Perfil Anual de Severidad y Profundidad Sísmica"
    texto_completo = f"{explicacion} {interpretacion}"
    encabezados = [
        "Año",
        "Cantidad de Sismos",
        "Magnitud Promedio",
        "Magnitud Máxima",
        "Profundidad Media (km)",
        "Epicentro del Mayor Sismo del Año",
    ]
    filas = [
        [
            r["year"],
            r["cantidad"],
            r["mag_promedio"],
            r["mag_max"],
            f"{r['depth_promedio']} km",
            r["lugar_max"],
        ]
        for r in registros
    ]
    return _generar_pdf_reporte(titulo, texto_completo, encabezados, filas, "reporte_dinamico_sismos.pdf")


def listado4_excel(request):
    registros, _, _, _, _ = _obtener_datos_dinamicos(request)
    encabezados = [
        "Año",
        "Cantidad de Sismos",
        "Magnitud Promedio",
        "Magnitud Máxima",
        "Profundidad Media (km)",
        "Epicentro del Mayor Sismo del Año",
    ]
    filas = [
        [
            int(r["year"]),
            int(r["cantidad"]),
            float(r["mag_promedio"]),
            float(r["mag_max"]),
            float(r["depth_promedio"]),
            str(r["lugar_max"]),
        ]
        for r in registros
    ]
    return _generar_excel_reporte("Reporte Dinamico", encabezados, filas, "reporte_dinamico_sismos.xlsx")



# Visualizaciones graficos 1, 2, 3, 4 (dinamico)

def grafico1(request):
    eventos_naturales = ["earthquake", "Ice Quake", "volcanic eruption", "landSlide", "Rock Slide"]
    df_antropico = df[~df["type"].isin(eventos_naturales)]
    top_tipos = df_antropico["type"].value_counts().head(5)

    fig, eje = plt.subplots(figsize=(10, 5.5), dpi=120)
    colores = ["#0072B2", "#E69F00", "#009E73", "#CC79A7", "#D55E00"][: len(top_tipos)]
    barras = eje.barh(top_tipos.index.astype(str), top_tipos.values, color=colores, height=0.68)
    eje.set_title("Principales Eventos Sísmicos Artificiales (1900-2026)", fontsize=15, fontweight="bold", loc="left", pad=16)
    eje.set_xlabel("Cantidad Registrada", fontsize=11, labelpad=8)
    eje.set_ylabel("Tipo de Evento", fontsize=11, labelpad=8)
    eje.bar_label(barras, padding=5, fontsize=10, fmt="%.0f")
    eje.set_xlim(0, top_tipos.max() * 1.15 if len(top_tipos) and top_tipos.max() > 0 else 1)
    eje.invert_yaxis()
    eje.grid(axis="x", linestyle="--", linewidth=0.7, alpha=0.5)
    eje.set_axisbelow(True)
    eje.spines[["top", "right", "left"]].set_visible(False)
    eje.tick_params(axis="y", length=0, labelsize=10)
    eje.tick_params(axis="x", labelsize=9)
    fig.tight_layout()

    grafico = _guardar_grafico(fig)
    explicacion = "Gráfico de barras horizontales que muestra los 5 principales tipos de eventos sísmicos de origen antrópico o humano (no naturales)."
    return render(request, "sismos/grafico.html", {"grafico": grafico, "explicacion": explicacion, "titulo": "Eventos Antrópicos"})


def grafico2(request):
    df_g2 = df.dropna(subset=["year", "mag"]).copy()
    df_g2["rango_mag"] = np.where(df_g2["mag"] < 5.0, "Menor a 5.0", "Mayor o igual a 5.0")
    evolucion = df_g2.groupby(["year", "rango_mag"]).size().unstack(fill_value=0)

    fig, eje = plt.subplots(figsize=(10, 5.5), dpi=120)
    if "Menor a 5.0" in evolucion.columns:
        eje.plot(evolucion.index, evolucion["Menor a 5.0"], label="Magnitud < 5.0", color="#0072B2", linewidth=2.5, linestyle="-", zorder=2)
    if "Mayor o igual a 5.0" in evolucion.columns:
        eje.plot(evolucion.index, evolucion["Mayor o igual a 5.0"], label="Magnitud ≥ 5.0", color="#D62728", linewidth=2.5, linestyle="--", zorder=3)

    eje.set_title("Sismos Fuertes vs Leves (1900-2026)", fontsize=15, fontweight="bold", loc="left", pad=16)
    eje.set_xlabel("Año", fontsize=11, labelpad=8)
    eje.set_ylabel("Cantidad Registrada", fontsize=11, labelpad=8)
    eje.legend(frameon=False, fontsize=10, loc="upper left")
    eje.grid(axis="y", linestyle="--", linewidth=0.7, alpha=0.5)
    eje.set_axisbelow(True)
    eje.spines[["top", "right", "left"]].set_visible(False)
    eje.tick_params(axis="y", length=0, labelsize=9)
    eje.tick_params(axis="x", labelsize=9)
    eje.xaxis.set_major_locator(plt.MaxNLocator(nbins=8, integer=True))
    fig.tight_layout()

    grafico = _guardar_grafico(fig)
    explicacion = "Gráfico de líneas que compara la evolución anual de los sismos con magnitud inferior a 5.0 y los de magnitud igual o superior a 5.0. El aumento de las detecciones puede relacionarse con mejoras en las redes de monitoreo sísmico."
    return render(request, "sismos/grafico.html", {"grafico": grafico, "explicacion": explicacion, "titulo": "Sismos Fuertes vs Leves"})


def grafico3(request):
    datos = df.dropna(subset=["latitude", "longitude", "mag"])
    muestra = datos.sample(n=min(125000, len(datos)), random_state=42)

    fig1, eje1 = plt.subplots(figsize=(10, 5), dpi=120)
    scatter1 = eje1.scatter(
        muestra["longitude"],
        muestra["latitude"],
        c=muestra["mag"],
        cmap="turbo",
        norm=plt.Normalize(vmin=muestra["mag"].quantile(0.01), vmax=muestra["mag"].quantile(0.97)),
        s=3,
        alpha=0.7,
        edgecolors="none",
        rasterized=True,
    )
    eje1.set_title("Distribución Global de Sismos (1900-2026)", fontsize=15, fontweight="bold", loc="left", pad=16)
    eje1.set_xlabel("Longitud (°)", fontsize=11, labelpad=8)
    eje1.set_ylabel("Latitud (°)", fontsize=11, labelpad=8)
    eje1.set_xlim(-180, 180)
    eje1.set_ylim(-90, 90)
    eje1.set_xticks(np.arange(-180, 181, 60))
    eje1.set_yticks(np.arange(-90, 91, 30))
    eje1.set_aspect("equal")
    eje1.grid(axis="both", linestyle="--", linewidth=0.7, alpha=0.5)
    eje1.set_axisbelow(True)
    eje1.spines[["top", "right", "bottom", "left"]].set_visible(True)
    eje1.tick_params(axis="both", labelsize=9)
    colorbar1 = fig1.colorbar(scatter1, ax=eje1, pad=0.03, fraction=0.035)
    colorbar1.set_label("Magnitud", fontsize=11, labelpad=8)
    colorbar1.ax.tick_params(labelsize=9)
    fig1.tight_layout(rect=[0, 0, 0.96, 1])

    grafico_sin_mapa = _guardar_grafico(fig1)

    fig2, eje2 = plt.subplots(figsize=(10, 5), dpi=120)
    url_paises = "https://raw.githubusercontent.com/nvkelso/natural-earth-vector/master/geojson/ne_110m_admin_0_countries.geojson"
    mapa_mundo = gpd.read_file(url_paises)
    mapa_mundo.plot(ax=eje2, color="dimgray", edgecolor="white", linewidth=0.5, zorder=0)

    scatter2 = eje2.scatter(
        muestra["longitude"],
        muestra["latitude"],
        c=muestra["mag"],
        cmap="turbo",
        norm=plt.Normalize(vmin=muestra["mag"].quantile(0.01), vmax=muestra["mag"].quantile(0.97)),
        s=3,
        alpha=0.7,
        edgecolors="none",
        rasterized=True,
        zorder=1,
    )
    eje2.set_title("Distribución Global de Sismos (1900-2026)", fontsize=15, fontweight="bold", loc="left", pad=16)
    eje2.set_xlabel("Longitud (°)", fontsize=11, labelpad=8)
    eje2.set_ylabel("Latitud (°)", fontsize=11, labelpad=8)
    eje2.set_xlim(-180, 180)
    eje2.set_ylim(-90, 90)
    eje2.set_xticks(np.arange(-180, 181, 60))
    eje2.set_yticks(np.arange(-90, 91, 30))
    eje2.set_aspect("equal")
    eje2.grid(axis="both", linestyle="--", linewidth=0.7, alpha=0.5)
    eje2.set_axisbelow(True)
    eje2.tick_params(axis="both", labelsize=9)
    colorbar2 = fig2.colorbar(scatter2, ax=eje2, pad=0.03, fraction=0.035)
    colorbar2.set_label("Magnitud", fontsize=11, labelpad=8)
    colorbar2.ax.tick_params(labelsize=9)
    fig2.tight_layout(rect=[0, 0, 0.96, 1])

    grafico_con_mapa = _guardar_grafico(fig2)

    explicacion = "Gráfico de dispersión que muestra una muestra de hasta 125.000 eventos sísmicos. La posición representa las coordenadas geográficas y el color indica la magnitud registrada. Se puede alternar la vista del mapa para observar los límites de las placas."

    return render(
        request,
        "sismos/grafico.html",
        {
            "grafico": grafico_sin_mapa,
            "grafico_mapa": grafico_con_mapa,
            "es_grafico_interactivo": True,
            "explicacion": explicacion,
            "titulo": "Distribución Global de Sismos",
        },
    )


def grafico4(request):
    registros, filtros, explicacion, interpretacion, kpis = _obtener_datos_dinamicos(request)

    # Si el navegador solicita actualización en tiempo real al mover un slider o selector
    if request.GET.get("ajax") == "1":
        return JsonResponse(
            {
                "datos": registros,
                "filtros": filtros,
                "explicacion": explicacion,
                "interpretacion": interpretacion,
                "kpis": kpis,
            }
        )

    explicacion_grafico = (
        "Visualización interactiva multidimensional en tiempo real: relaciona el año del evento, "
        "la magnitud máxima y media alcanzada, el volumen anual de sismos (tamaño de burbuja) y la "
        "profundidad hipocentral media (escala cromática), actualizándose automáticamente al modificar cualquier parámetro."
    )

    return render(
        request,
        "sismos/grafico.html",
        {
            "titulo": "Perfil Dinámico de Severidad, Frecuencia y Profundidad Sísmica",
            "explicacion": explicacion_grafico,
            "interpretacion": interpretacion,
            "datos_json": json.dumps(registros),
            "kpis": kpis,
            "es_grafico_dinamico": True,
            "filtros": filtros,
            "regiones": REGIONES_DISPONIBLES,
            "profundidades": PROFUNDIDADES_DISPONIBLES,
        },
    )


def _guardar_grafico(figura):
    buffer = BytesIO()
    figura.savefig(buffer, format="png", dpi=120)
    plt.close(figura)
    buffer.seek(0)
    grafico = base64.b64encode(buffer.getvalue()).decode("utf-8")
    buffer.close()
    return grafico