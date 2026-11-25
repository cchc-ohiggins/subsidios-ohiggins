#!/usr/bin/env python3
"""
Genera las minutas Word (una por región y una nacional) a partir de data/subsidios.json.

Uso:
    python3 etl/build_minutas.py data/subsidios.json minutas/

Formato: Arial 12, interlineado 1,5. Cifras de montos en UF (unidad de la fuente).
Los textos se arman con los datos de cada región: si cambian las cifras, cambia la redacción.
"""
import json
import os
import re
import sys
import tempfile
import unicodedata
import datetime as dt

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ORANGE, BLUE, NAVY, GRAY = "#E74100", "#006EDC", "#003D6B", "#D9D9D9"
CAT_COL = ["#006EDC", "#2A9D8F", "#7B5CD6", "#B8860B"]
CAT_KEYS = ["SECTORES VULNERABLES", "SECTORES MEDIOS Y EMERGENTES", "MEJORAMIENTO Y MANTENCIÓN DE VIVIENDA", "ARRIENDO"]
CAT_NAMES = ["Sectores vulnerables", "Sectores medios y emergentes", "Mejoramiento y mantención", "Arriendo"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

REG_DE = {"País": "el país", "Metropolitana": "la Región Metropolitana", "Maule": "la Región del Maule", "Biobío": "la Región del Biobío"}
NOMBRE_PROG = {
    "Programa de Mejoramiento Viviendas y Barrios (DS 27)": "Mejoramiento de Viviendas y Barrios (DS 27)",
    "Fondo Solidario de Elección de Vivienda (FSEV) DS 49": "Fondo Solidario de Elección de Vivienda (DS 49)",
    "Programa Integración Social y Territorial DS19": "Integración Social y Territorial (DS 19)",
    "Sistema Integrado de Subsidios DS 01 Tramo 1, 2 y 3": "Sistema Integrado de Subsidios (DS 01, tramos 1, 2 y 3)",
    "Sistema Integrado de Subsidios DS 01 Llamado Arriendo": "Sistema Integrado, llamado de arriendo (DS 01)",
    "Sistema Integrado de Subsidios DS 01 Organizaciones de trabajadores": "Sistema Integrado, organizaciones de trabajadores (DS 01)",
    "Subsidio al Arriendo DS 52": "Subsidio de Arriendo (DS 52)",
    "Protección Patrimonio Familiar (PPPF) DS 255": "Protección del Patrimonio Familiar (DS 255)",
    "Leasing Habitacional DS 120": "Leasing Habitacional (DS 120)",
    "Habitabilidad Rural DS 10 Vivienda Nueva": "Habitabilidad Rural (DS 10), vivienda nueva",
    "Habitabilidad Rural DS 10 Mejoramientos": "Habitabilidad Rural (DS 10), mejoramientos",
    "Progresiva etapa I y II DS 140": "Vivienda Progresiva (DS 140)",
}


# ---------------------------------------------------------------- formato
def n(x):
    return f"{round(x):,}".replace(",", ".")


def dec(x, d=1):
    s = f"{abs(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if x < 0 else "") + s


def pc(x, d=1):
    return dec(x, d) + "%"


def mmuf(x, d=2):
    return dec(x / 1e6, d)


def slug(s):
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"['’]", "", s.lower())
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def prog_name(p):
    return NOMBRE_PROG.get(p, p)


# ---------------------------------------------------------------- datos
class Datos:
    def __init__(self, path):
        with open(path, encoding="utf-8") as f:
            self.d = json.load(f)
        m = self.d["meta"]
        self.Y = m["anios"]
        self.L = self.Y[-1]
        self.F = self.L - 1
        self.mes = m.get("corte_nombre", "octubre")
        self.abr = self.mes[:3]
        self.cat = {p["nombre"]: p["categoria"] for p in self.d["programas"]}
        self.regs = self.d["regiones"]
        self.meta = m

    def sec(self, r):
        return self.d["pais"] if r == "País" else self.d["reg"][r]

    def ix(self, y):
        return y - self.Y[0]

    def tot(self, r, sc="T"):
        return self.sec(r)[sc]["total"]

    def u(self, r, y, sc="T"):
        return self.tot(r, sc)["u"][self.ix(y)]

    def uf(self, r, y, sc="T"):
        return self.tot(r, sc)["uf"][self.ix(y)]

    def suma(self, r, a, b, k="u", sc="T"):
        return sum(self.tot(r, sc)[k][self.ix(y)] for y in range(a, b + 1))

    def prog(self, r, sc="T"):
        return self.sec(r)[sc]["prog"]

    def grupos(self, r, y):
        out = [0.0] * 4
        for name, p in self.prog(r).items():
            try:
                k = CAT_KEYS.index(self.cat.get(name, ""))
            except ValueError:
                continue
            out[k] += p["u"][self.ix(y)]
        return out

    def ranking(self, y=None, a=None, b=None):
        if y is not None:
            a = b = y
        arr = sorted(((r, self.suma(r, a, b)) for r in self.regs), key=lambda t: -t[1])
        return arr


# ---------------------------------------------------------------- redacción
def etiqueta(D, r):
    return REG_DE.get(r, f"la Región de {r}")


def variacion(cur, prev):
    if prev <= 0:
        return None
    return (cur / prev - 1) * 100


def frase_var(d, ref):
    if d is None:
        return ""
    if abs(d) < 0.5:
        return f", prácticamente lo mismo que en {ref}"
    return f", {pc(abs(d))} {'más' if d > 0 else 'menos'} que en {ref}"


def redactar(D, r):
    F, L, p = D.F, D.L, D.F - 1
    en = etiqueta(D, r)
    es_pais = r == "País"
    uF, uP, uL = D.u(r, F), D.u(r, p), D.u(r, L)
    ufF, ufL = D.uf(r, F), D.uf(r, L)
    d = variacion(uF, uP)
    out = {}

    # --- resumen
    t = f"En {F} se otorgaron **{n(uF)} subsidios habitacionales** {'en' if not es_pais else 'a nivel'} {en if not es_pais else 'nacional'}{frase_var(d, p)}. "
    t += f"El monto asociado fue de {mmuf(ufF)} millones de UF, lo que da un promedio de {n(ufF / uF)} UF por subsidio. "
    if not es_pais:
        sh = uF / D.u("País", F) * 100
        rk = [x for x, _ in D.ranking(y=F)].index(r) + 1
        t += f"La región aportó el {pc(sh)} de los subsidios del país y quedó en el lugar {rk} entre las 16 regiones. "
    t += f"Entre enero y {D.mes} de {L} van {n(uL)} subsidios, el {pc(uL / uF * 100, 0)} de lo otorgado en todo {F}."
    out["resumen"] = t

    # --- tendencia
    fy = next(y for y in D.Y if D.u(r, y) > 0)          # primer año con datos propios
    t = ""
    if fy > D.Y[0]:
        t += f"{r} figura con datos propios desde {fy}; antes de esa fecha sus subsidios están contados en otra región. "
    if fy <= F - 9:
        a1 = D.suma(r, F - 4, F) / 5
        a0 = D.suma(r, F - 9, F - 5) / 5
        t += f"En los últimos cinco años completos ({F - 4}–{F}) el promedio fue de {n(a1)} subsidios al año; en los cinco anteriores ({F - 9}–{F - 5}) había sido de {n(a0)}"
        dd = variacion(a1, a0)
        t += (f", una {'baja' if dd < 0 else 'alza'} de {pc(abs(dd), 0)}. " if dd is not None and abs(dd) >= 0.5 else ". ")
    else:
        n_anios = F - fy + 1
        t += f"Desde {fy} el promedio fue de {n(D.suma(r, fy, F) / n_anios)} subsidios al año. "
    serie = [(y, D.u(r, y)) for y in D.Y if fy <= y < L]
    ymax, vmax = max(serie, key=lambda t: t[1])
    t += f"El año con más subsidios de toda la serie ({fy}–{F}) fue {ymax}, con {n(vmax)}"
    if ymax in (2010, 2011, 2012):
        t += ", un año marcado por los subsidios de reconstrucción"
    t += ". "
    if fy <= 2013 and ymax < 2013:
        ymax2, vmax2 = max([(y, D.u(r, y)) for y in range(2013, L)], key=lambda t: t[1])
        t += f"Desde 2013, cuando dejaron de contratarse viviendas SERVIU y solo se otorgan subsidios, el máximo fue {ymax2} con {n(vmax2)}."
    out["tendencia"] = t.strip()

    # --- programas
    pr = sorted(((nm, v["u"][D.ix(F)], v["u"][D.ix(p)]) for nm, v in D.prog(r).items()), key=lambda t: -t[1])
    top = [x for x in pr if x[1] > 0][:3]
    t = ""
    if top:
        items = [f"{prog_name(a)}, con {n(b)}" for a, b, _ in top]
        lista = items[0] if len(items) == 1 else "; ".join(items[:-1]) + "; y " + items[-1]
        t += f"Por programa, en {F} los de mayor volumen fueron: {lista}. "
    g = D.grupos(r, F)
    tg = sum(g) or 1
    orden = sorted(range(4), key=lambda k: -g[k])
    t += (f"Por grupo, los {CAT_NAMES[orden[0]].lower()} reunieron el {pc(g[orden[0]] / tg * 100, 0)} de los subsidios, "
          f"seguidos por {CAT_NAMES[orden[1]].lower()} ({pc(g[orden[1]] / tg * 100, 0)}) y {CAT_NAMES[orden[2]].lower()} ({pc(g[orden[2]] / tg * 100, 0)}). ")
    cae = [x for x in sorted(pr, key=lambda t: -t[2]) if x[2] > 0 and x[1] == 0][:2]
    if cae:
        partes = [f"{prog_name(a)} no registra subsidios en {F} (en {p} había otorgado {n(c)})" for a, _, c in cae]
        t += "Entre un año y otro cambió la oferta: " + "; ".join(partes) + ". "
    out["programas"] = t.strip()

    # --- montos
    prom = ufF / uF if uF else 0
    prom_pais = D.uf("País", F) / D.u("País", F)
    prom_ref = D.uf(r, F - 6) / D.u(r, F - 6) if D.u(r, F - 6) else 0
    t = f"El monto otorgado en {F} fue de {mmuf(ufF)} millones de UF. El promedio por subsidio fue de {n(prom)} UF"
    t += f" (en {F - 6} había sido de {n(prom_ref)} UF)" if prom_ref else ""
    t += ". " if es_pais else f", frente a {n(prom_pais)} UF del promedio nacional. "
    t += "Hay que leer este promedio con cuidado: depende de qué programas se llamaron en el año, porque los de vivienda nueva tienen montos mucho mayores que los de mejoramiento."
    out["montos"] = t

    # --- comparación
    rkF = D.ranking(y=F)
    if es_pais:
        tt = sum(v for _, v in rkF)
        top3 = rkF[:3]
        t = f"Por regiones, en {F} concentraron más subsidios " + ", ".join(
            f"{a} ({n(v)}; {pc(v / tt * 100, 0)})" for a, v in top3[:-1]) + f" y {top3[-1][0]} ({n(top3[-1][1])}; {pc(top3[-1][1] / tt * 100, 0)}). "
        oh = "O'Higgins"
        t += f"{oh} quedó en el lugar {[a for a, _ in rkF].index(oh) + 1} con {n(dict(rkF)[oh])}."
    else:
        a, b = F - 4, F
        acc = D.ranking(a=a, b=b)
        rk_acc = [x for x, _ in acc].index(r) + 1
        sh_acc = dict(acc)[r] / sum(v for _, v in acc) * 100
        nombre = en
        t = (f"Entre {a} y {b} {nombre} acumuló {n(dict(acc)[r])} subsidios, el {pc(sh_acc)} del país, lo que la ubica en el lugar {rk_acc} del período. ")
        top3 = [x for x in rkF if x[0] != r][:3]
        t += f"Las regiones con más subsidios en {F} fueron " + ", ".join(f"{a} ({n(v)})" for a, v in top3[:-1]) + f" y {top3[-1][0]} ({n(top3[-1][1])}), frente a los {n(uF)} de {r}."
    out["comparacion"] = t

    # --- año en curso
    sin = [(nm, v["u"][D.ix(F)]) for nm, v in D.prog("País").items() if v["u"][D.ix(L)] == 0 and v["u"][D.ix(F)] > 0]
    sin = [x for x in sorted(sin, key=lambda t: -t[1])][:3]
    t = f"Los datos de {L} llegan hasta {D.mes} y no se pueden comparar con un año completo. Los subsidios dependen de los llamados de cada programa"
    if sin:
        items = [f"{prog_name(a)}, que otorgó {n(b)} en {F}" for a, b in sin]
        lista = items[0] if len(items) == 1 else "; ".join(items[:-1]) + "; y " + items[-1]
        t += f". A nivel nacional, a la fecha no registran subsidios este año: {lista}"
    t += "."
    out["en_curso"] = t
    return out


# ---------------------------------------------------------------- gráficos
def estilo():
    fam = "Liberation Sans"
    if not any(f.name == fam for f in font_manager.fontManager.ttflist):
        fam = "DejaVu Sans"
    plt.rcParams.update({
        "font.family": fam, "font.size": 10, "axes.edgecolor": "#B9C4D0", "axes.linewidth": .8,
        "axes.spines.top": False, "axes.spines.right": False, "xtick.color": "#4A5C70", "ytick.color": "#4A5C70",
        "axes.labelcolor": "#4A5C70", "text.color": "#14273A",
    })


def fig_anual(D, r, ruta):
    años = list(range(2015, D.L + 1))
    vals = [D.u(r, y) for y in años]
    col = BLUE if r == "País" else ORANGE
    fig, axs = plt.subplots(1, 2, figsize=(10.2, 3.5), gridspec_kw={"width_ratios": [1.5, 1], "wspace": .28})
    ax = axs[0]
    for y, v in zip(años, vals):
        ax.bar(y, v, width=.62, color=col if y < D.L else "white", edgecolor=col, hatch="////" if y == D.L else None, linewidth=0 if y < D.L else 1)
    ax.set_xticks(años)
    ax.set_xticklabels([str(y)[2:] if y < D.L else f"{str(y)[2:]}*" for y in años])
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: n(v)))
    ax.grid(axis="y", color="#E9EDF2", lw=.8); ax.set_axisbelow(True)
    ax.set_title("Subsidios otorgados por año", loc="left", fontsize=11, fontweight="bold", color=NAVY)
    full = [(y, v) for y, v in zip(años, vals) if y < D.L]
    ym, vm = max(full, key=lambda t: t[1])
    ax.annotate(n(vm), (ym, vm), ha="center", va="bottom", fontsize=9, fontweight="bold", xytext=(0, 3), textcoords="offset points")
    yl, vl = full[-1]
    if yl != ym:
        ax.annotate(n(vl), (yl, vl), ha="center", va="bottom", fontsize=9, fontweight="bold", xytext=(0, 3), textcoords="offset points")
    ax.text(0, -.2, f"* {D.L}: enero–{D.mes}", transform=ax.transAxes, fontsize=8, color="#6B7C8F")

    ax = axs[1]
    años2 = list(range(D.F - 4, D.F + 1))
    bott = [0] * len(años2)
    for k in range(4):
        vs = [D.grupos(r, y)[k] for y in años2]
        ax.bar(range(len(años2)), vs, bottom=bott, width=.62, color=CAT_COL[k], label=CAT_NAMES[k], edgecolor="white", linewidth=1)
        bott = [a + b for a, b in zip(bott, vs)]
    ax.set_xticks(range(len(años2))); ax.set_xticklabels([str(y) for y in años2])
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: n(v)))
    ax.grid(axis="y", color="#E9EDF2", lw=.8); ax.set_axisbelow(True)
    ax.set_title("Por grupo de programas", loc="left", fontsize=11, fontweight="bold", color=NAVY)
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.12), ncol=2, frameon=False, fontsize=8, handlelength=1, columnspacing=1)
    fig.savefig(ruta, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def fig_ranking(D, r, ruta):
    a, b = D.F - 4, D.F
    arr = D.ranking(a=a, b=b)
    fig, ax = plt.subplots(figsize=(10.2, 3.7))
    nombres = [x for x, _ in arr][::-1]
    vals = [v for _, v in arr][::-1]
    colores = [ORANGE if x == r else "#4D95E3" for x in nombres]
    ax.barh(nombres, vals, color=colores, height=.62)
    for i, v in enumerate(vals):
        ax.text(v + max(vals) * .01, i, n(v), va="center", fontsize=8.5, fontweight="bold" if nombres[i] == r else "normal")
    ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: n(v)))
    ax.set_xlim(0, max(vals) * 1.12)
    ax.grid(axis="x", color="#E9EDF2", lw=.8); ax.set_axisbelow(True)
    ax.set_title(f"Subsidios otorgados por región, {a}–{b}", loc="left", fontsize=11, fontweight="bold", color=NAVY)
    for lbl in ax.get_yticklabels():
        if lbl.get_text() == r:
            lbl.set_fontweight("bold")
    fig.savefig(ruta, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ---------------------------------------------------------------- docx
def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hexcolor.lstrip("#"))
    tcPr.append(sh)


def cell_borders(cell, **kw):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for edge, (sz, color) in kw.items():
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), "0"); e.set(qn("w:color"), color.lstrip("#"))
        b.append(e)
    tcPr.append(b)


def fuente(run, size=12, bold=False, color=None, italic=False):
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color.lstrip("#"))


def rico(par, texto, size=12, color=None):
    for i, parte in enumerate(texto.split("**")):
        if parte:
            fuente(par.add_run(parte), size=size, bold=(i % 2 == 1), color=color)


def parrafo(doc, texto, size=12, espacio=6, align=None, color=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(espacio)
    pf.space_before = Pt(0)
    if align:
        p.alignment = align
    rico(p, texto, size=size, color=color)
    return p


def titulo(doc, texto):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(10); pf.space_after = Pt(2); pf.line_spacing = 1.5; pf.keep_with_next = True
    fuente(p.add_run(texto), size=14, bold=True, color=NAVY)
    return p


def campo_pagina(par):
    def run_fld(tipo=None, instr=None, texto=None):
        r = par.add_run(); fuente(r, size=9, color="#6B7C8F")
        if tipo:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tipo); r._r.append(e)
        if instr:
            i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = instr; r._r.append(i)
        if texto:
            r.text = texto
    run_fld("begin"); run_fld(instr="PAGE"); run_fld("separate"); run_fld(texto="1"); run_fld("end")


def tabla_cifras(doc, D, r):
    años = list(range(D.F - 4, D.L + 1))
    cab = ["Año", "Subsidios", "Variación anual", "Monto (MM UF)", "UF por subsidio", "% del país"]
    t = doc.add_table(rows=1, cols=len(cab))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(cab):
        c = t.rows[0].cells[i]; shade(c, NAVY)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT
        fuente(p.add_run(h), size=10, bold=True, color="#FFFFFF")
    for k, y in enumerate(años):
        u, uf = D.u(r, y), D.uf(r, y)
        var = variacion(u, D.u(r, y - 1)) if y < D.L else None
        pp = D.u("País", y)
        fila = [f"{y}" + (f" (ene–{D.abr})" if y == D.L else ""), n(u), (pc(var) if var is not None else "—"), mmuf(uf), n(uf / u) if u else "—", (pc(u / pp * 100) if r != "País" else "100%")]
        cells = t.add_row().cells
        for i, v in enumerate(fila):
            if k % 2 == 1:
                shade(cells[i], "#F2F5F9")
            cell_borders(cells[i], bottom=(4, GRAY))
            p = cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            fuente(p.add_run(v), size=10)
    for ri, row in enumerate(t.rows):
        trPr = row._tr.get_or_add_trPr(); cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.keep_with_next = ri < len(t.rows) - 1
    return t


def construir(D, r, ruta_docx, tmp):
    txt = redactar(D, r)
    f1 = os.path.join(tmp, f"anual_{slug(r)}.png")
    f2 = os.path.join(tmp, f"rank_{slug(r)}.png")
    fig_anual(D, r, f1)
    fig_ranking(D, r, f2) if r != "País" else fig_ranking(D, "O'Higgins", f2)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)  # carta
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(2.0)
    st = doc.styles["Normal"]
    st.font.name = "Arial"; st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    # ---- banda de título
    b = doc.add_table(rows=1, cols=1)
    b.autofit = False
    c = b.rows[0].cells[0]; c.width = Cm(17.19); shade(c, NAVY); cell_borders(c, bottom=(36, ORANGE))
    p = c.paragraphs[0]; p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(0)
    fuente(p.add_run("Minuta · Subsidios habitacionales otorgados"), size=12, color="#CFE0F2")
    p = c.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    fuente(p.add_run("Total país" if r == "País" else r), size=28, bold=True, color="#FFFFFF")
    hoy = dt.date.today()
    p = c.add_paragraph(); p.paragraph_format.space_after = Pt(8)
    pub = D.meta.get("publicado", "")
    fuente(p.add_run(f"CChC Sede Región de O'Higgins · Área de Estudios · {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}" + (f" · Datos MINVU publicados el {pub}" if pub else "")), size=9, color="#CFE0F2")

    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    titulo(doc, "Resumen")
    parrafo(doc, txt["resumen"])
    parrafo(doc, txt["tendencia"])
    pic = doc.add_paragraph(); pic.alignment = WD_ALIGN_PARAGRAPH.CENTER; pic.paragraph_format.space_after = Pt(4)
    pic.add_run().add_picture(f1, width=Cm(17))

    titulo(doc, "Cifras de los últimos años")
    tabla_cifras(doc, D, r)
    parrafo(doc, f"Fuente: MINVU. Montos en millones de UF (MM UF), tal como los publica el ministerio. {D.L} corresponde a enero–{D.mes}.", size=9, espacio=4, color="#6B7C8F")

    titulo(doc, "Programas")
    parrafo(doc, txt["programas"])
    titulo(doc, "Montos")
    parrafo(doc, txt["montos"])
    titulo(doc, "Comparación con otras regiones")
    parrafo(doc, txt["comparacion"])
    pic = doc.add_paragraph(); pic.alignment = WD_ALIGN_PARAGRAPH.CENTER; pic.paragraph_format.space_after = Pt(4)
    pic.add_run().add_picture(f2, width=Cm(16))
    titulo(doc, f"Qué tener en cuenta sobre {D.L}")
    parrafo(doc, txt["en_curso"])
    parrafo(doc, "Las UF pueden incluir ajustes de subsidios otorgados en años anteriores, de modo que no siempre reflejan solo lo asignado en el período. Las regiones de Arica y Parinacota, Los Ríos y Ñuble aparecen con datos desde su creación; antes, sus subsidios están dentro de Tarapacá, Los Lagos y Biobío.", size=11)

    # ---- pie de página
    pie = sec.footer.paragraphs[0]
    fuente(pie.add_run("Fuente: MINVU, Estadísticas · Viviendas contratadas y subsidios otorgados. Elaboración: CChC O'Higgins, Estudios.    Página "), size=9, color="#6B7C8F")
    campo_pagina(pie)
    doc.save(ruta_docx)


def main(json_path, out_dir):
    estilo()
    D = Datos(json_path)
    os.makedirs(out_dir, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        for r in ["País"] + D.regs:
            nombre = "pais" if r == "País" else slug(r)
            construir(D, r, os.path.join(out_dir, f"Minuta_subsidios_{nombre}.docx"), tmp)
            print("OK", nombre)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(1)
    main(sys.argv[1], sys.argv[2])
