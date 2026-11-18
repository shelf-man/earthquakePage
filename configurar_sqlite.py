import os
import sqlite3
import pandas as pd

ruta_db = "db.sqlite3"
ruta_csv = os.path.join("data", "earthquakes_1900_2026.csv")
columnas = ["time", "latitude", "longitude", "depth", "mag", "place", "mag_category", "type"]

with sqlite3.connect(ruta_db) as conn:
    cursor = conn.cursor()
    cursor.execute("PRAGMA journal_mode=WAL;")
    cursor.execute("PRAGMA synchronous=NORMAL;")
    cursor.execute("PRAGMA temp_store=MEMORY;")

    primero = True
    for bloque in pd.read_csv(ruta_csv, usecols=columnas, chunksize=100000):
        fechas = pd.to_datetime(bloque["time"], format="ISO8601", errors="coerce")
        bloque["year"] = fechas.dt.year
        bloque.to_sql(
            "sismos_sismo",
            conn,
            if_exists="replace" if primero else "append",
            index=False,
        )
        primero = False

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sismos_year ON sismos_sismo(year);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sismos_mag ON sismos_sismo(mag);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sismos_depth ON sismos_sismo(depth);")
    conn.commit()

    cursor.execute("SELECT COUNT(*) FROM sismos_sismo;")
    total = cursor.fetchone()[0]
    print(f"Base de datos SQLite configurada con {total} registros.")