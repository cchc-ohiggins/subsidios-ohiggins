#!/usr/bin/env python3
"""
Extrae el Excel "SU Otorgados 1990 a <Mes> <Año>" del MINVU y genera data/subsidios.json
para el dashboard.

Uso:
    python3 etl/build_data.py ruta/al/archivo.xlsx data/subsidios.json

Estructura del Excel (una hoja por región con código de 2 letras, "total país" y "año y región"):
  - Bloque 1: SUBSIDIOS TOTAL OTORGADOS POR PROGRAMA
  - Bloque 2: SUBSIDIOS REGULAR OTORGADOS POR PROGRAMA
  - Bloque 3: SUBSIDIOS RECONSTRUCCIÓN OTORGADOS POR PROGRAMA
  Cada programa ocupa dos filas: Unid. y U.F., con una columna por año.

El script identifica a qué región corresponde cada hoja comparando la serie anual de la hoja
con la de la hoja resumen "año y región" (no depende del código de la hoja), y valida:
  1) suma de programas = total de la hoja
  2) suma de regiones = total país
  3) total de cada hoja = hoja resumen "año y región"
"""
import json
import re
import sys
import datetime as dt
import openpyxl

REGIONES = [
    "Arica y Parinacota", "Tarapacá", "Antofagasta", "Atacama", "Coquimbo",
    "Valparaíso", "Metropolitana", "O'Higgins", "Maule", "Ñuble", "Biobío",
    "La Araucanía", "Los Ríos", "Los Lagos", "Aysén", "Magallanes",
]
SECCIONES = {"T": "SUBSIDIOS TOTAL", "R": "SUBSIDIOS REGULAR", "X": "SUBSIDIOS RECONSTRUCCI"}
HOJAS_NO_REGION = {"total país", "año y región"}  # además de cualquier hoja "Ingreso de Datos <año>"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def es_hoja_no_region(nombre):
    return nombre in HOJAS_NO_REGION or nombre.lower().startswith("ingreso de datos")


def mes_de_corte(nombre_archivo, wb=None):
    """Mes hasta el que llegan los datos del último año, leído del nombre del archivo ('... a Octubre 2026')."""
    base = nombre_archivo.lower().replace("-", " ").replace("_", " ")
    for i, m in enumerate(MESES, 1):
        if re.search(r"\b" + m + r"\b", base) or (m == "septiembre" and "setiembre" in base):
            return i
    raise RuntimeError("No pude leer el mes de corte en el nombre del archivo: " + nombre_archivo)


def limpia(s):
    if s is None:
        return None
    return re.sub(r"\s+", " ", str(s)).strip()


def num(v):
    if v is None or v == "":
        return 0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0


def parse_hoja(ws):
    """Devuelve {seccion: {'years': [...], 'total': {'u': [...], 'uf': [...]}, 'prog': {nombre: {...}}}}"""
    rows = list(ws.iter_rows(values_only=True))
    out = {}
    # ubicar encabezados de sección
    starts = []
    for i, r in enumerate(rows):
        a = limpia(r[0]) if r and r[0] else ""
        for k, pref in SECCIONES.items():
            if a.upper().startswith(pref):
                starts.append((i, k))
    for idx, (i0, k) in enumerate(starts):
        i_end = starts[idx + 1][0] if idx + 1 < len(starts) else len(rows)
        # fila de años: primera fila después del encabezado con 1990 en col D
        hdr = None
        for j in range(i0, min(i0 + 6, i_end)):
            if rows[j][3] == 1990:
                hdr = j
                break
        if hdr is None:
            continue
        cols = [(c, int(v)) for c, v in enumerate(rows[hdr]) if isinstance(v, (int, float)) and 1990 <= v <= 2100]
        years = [y for _, y in cols]
        ci = [c for c, _ in cols]
        total = None
        prog = {}
        categoria = None
        j = hdr + 1
        while j < i_end:
            r = rows[j]
            a = limpia(r[0]) if r[0] else None
            b = limpia(r[1]) if r[1] else None
            c = limpia(r[2]) if r[2] else None
            if a and a.upper().startswith("FUENTE"):
                break
            if a and not a.upper().startswith(("NOTA", "PUBLICADO")):
                categoria = a
            if c and c.lower().startswith("unid") and b:
                nxt = rows[j + 1]
                u = [num(r[x]) for x in ci]
                uf = [num(nxt[x]) for x in ci]
                if b.lower().startswith("total"):
                    total = {"u": u, "uf": uf}
                else:
                    prog[b] = {"cat": categoria, "u": u, "uf": uf}
                j += 2
                continue
            j += 1
        out[k] = {"years": years, "total": total, "prog": prog}
    return out


def serie_resumen(wb):
    """Hoja 'año y región', primer bloque (total): {region: {'u': [...], 'uf': [...]}} y total país."""
    rows = list(wb["año y región"].iter_rows(values_only=True))
    hdr = next(i for i, r in enumerate(rows) if r[2] == 1990)
    cols = [c for c, v in enumerate(rows[hdr]) if isinstance(v, (int, float)) and 1990 <= v <= 2100]
    years = [int(rows[hdr][c]) for c in cols]
    res = {}
    j = hdr + 1
    while j < len(rows):
        r = rows[j]
        a = limpia(r[0]) if r[0] else None
        if a and a.upper().startswith("FUENTE"):
            break
        if a and r[1] and str(r[1]).lower().startswith("unid"):
            res[a] = {"u": [num(r[c]) for c in cols], "uf": [num(rows[j + 1][c]) for c in cols]}
            j += 2
            continue
        j += 1
    return years, res


def main(xlsx, salida):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    years_res, resumen = serie_resumen(wb)

    hojas = {}
    for ws in wb.worksheets:
        if es_hoja_no_region(ws.title):
            continue
        hojas[ws.title] = parse_hoja(ws)
    pais = parse_hoja(wb["total país"])

    # --- mapear hoja -> región comparando la serie de unidades 1990-2026
    mapa = {}
    for cod, d in hojas.items():
        serie = d["T"]["total"]["u"]
        mejor, dist = None, None
        for reg in REGIONES:
            if reg not in resumen:
                continue
            ref = resumen[reg]["u"]
            n = min(len(ref), len(serie))
            dd = sum(abs(ref[i] - serie[i]) for i in range(n))
            if dist is None or dd < dist:
                mejor, dist = reg, dd
        mapa[cod] = (mejor, dist)
    regiones_hoja = {}
    problemas = []
    for cod, (reg, dist) in mapa.items():
        if dist and dist > 0.5:
            problemas.append(f"Hoja {cod} ~ {reg}: diferencia vs hoja 'año y región' = {dist:.1f}")
        if reg in regiones_hoja:
            problemas.append(f"Región duplicada: {reg} (hojas {regiones_hoja[reg]} y {cod})")
        regiones_hoja[reg] = cod
    faltan = [r for r in REGIONES if r not in regiones_hoja]
    if faltan:
        problemas.append(f"Regiones sin hoja: {faltan}")

    years = hojas[next(iter(hojas))]["T"]["years"]
    nY = len(years)

    # --- validaciones
    for reg, cod in regiones_hoja.items():
        for k in ("T", "R", "X"):
            sec = hojas[cod][k]
            su = [sum(p["u"][i] for p in sec["prog"].values()) for i in range(nY)]
            suf = [sum(p["uf"][i] for p in sec["prog"].values()) for i in range(nY)]
            du = max(abs(su[i] - sec["total"]["u"][i]) for i in range(nY))
            duf = max(abs(suf[i] - sec["total"]["uf"][i]) for i in range(nY))
            if du > 0.5 or duf > 5:
                problemas.append(f"{reg} [{k}] suma de programas != total (max dif unid={du:.1f}, UF={duf:.1f})")
    # total regional = suma programas = total país (sección T)
    for i, y in enumerate(years):
        s = sum(hojas[c]["T"]["total"]["u"][i] for c in regiones_hoja.values())
        t = pais["T"]["total"]["u"][i]
        if abs(s - t) > 0.5:
            problemas.append(f"{y}: suma regiones (unid) {s:.0f} != total país {t:.0f}")
        suf = sum(hojas[c]["T"]["total"]["uf"][i] for c in regiones_hoja.values())
        tuf = pais["T"]["total"]["uf"][i]
        if abs(suf - tuf) > 5:
            problemas.append(f"{y}: suma regiones (UF) {suf:.0f} != total país {tuf:.0f}")

    # --- armar salida compacta
    # programas (union, en orden de aparición) con categoría
    programas = []
    cat_prog = {}
    for k in ("T", "R", "X"):
        for p, d in pais[k]["prog"].items():
            if p not in cat_prog:
                cat_prog[p] = d["cat"]
                programas.append(p)
    for cod in regiones_hoja.values():
        for k in ("T", "R", "X"):
            for p, d in hojas[cod][k]["prog"].items():
                if p not in cat_prog:
                    cat_prog[p] = d["cat"]
                    programas.append(p)

    def empaqueta(secc_dict):
        # {prog: {"u":[...], "uf":[...]}} (redondeos: unidades a entero, UF a 1 decimal)
        o = {}
        for p, d in secc_dict["prog"].items():
            if any(d["u"]) or any(d["uf"]):
                o[p] = {"u": [round(x) for x in d["u"]], "uf": [round(x, 1) for x in d["uf"]]}
        return {
            "total": {"u": [round(x) for x in secc_dict["total"]["u"]],
                      "uf": [round(x, 1) for x in secc_dict["total"]["uf"]]},
            "prog": o,
        }

    datos = {
        "meta": {
            "fuente": "MINVU – Estadísticas: Viviendas contratadas y subsidios otorgados (reporte mensual Metas Subsidios Asignados DPH a DIFIN)",
            "url_fuente": "https://www.minvu.gob.cl/elementos-tecnicos/estadisticas/viviendas-contratadas-y-subsidios-otorgados/",
            "archivo": xlsx.split("/")[-1],
            "generado": dt.date.today().isoformat(),
            "anios": years,
            "anio_parcial": years[-1],
            "corte_mes": mes_de_corte(xlsx.split("/")[-1]),
            "corte_nombre": MESES[mes_de_corte(xlsx.split("/")[-1]) - 1],
            "nota_uf": "Las Unidades de Fomento pueden incorporar incrementos correspondientes a subsidios otorgados en años anteriores; por lo tanto, podrían no reflejar las unidades asignadas específicamente al periodo actual.",
            "validacion": "OK" if not problemas else "CON OBSERVACIONES",
            "observaciones": problemas,
        },
        "programas": [{"nombre": p, "categoria": cat_prog[p]} for p in programas],
        "regiones": REGIONES,
        "pais": {k: empaqueta(pais[k]) for k in ("T", "R", "X")},
        "reg": {reg: {k: empaqueta(hojas[cod][k]) for k in ("T", "R", "X")} for reg, cod in regiones_hoja.items()},
    }
    # fecha de publicación si está en el archivo
    for r in wb["año y región"].iter_rows(values_only=True):
        if r[0] and str(r[0]).startswith("Publicado el"):
            datos["meta"]["publicado"] = str(r[0]).replace("Publicado el", "").strip()
            break

    with open(salida, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, separators=(",", ":"))

    print("Hoja -> región:", {c: m[0] for c, m in mapa.items()})
    print("Programas:", len(programas), "| años:", years[0], "-", years[-1])
    print("Validación:", datos["meta"]["validacion"])
    for p in problemas:
        print("  !", p)
    return 0 if not problemas else 2


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    sys.exit(main(sys.argv[1], sys.argv[2]))
