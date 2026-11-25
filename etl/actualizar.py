#!/usr/bin/env python3
"""
Actualiza el dashboard cuando el MINVU publica un Excel nuevo.

1) Entra a la página de Estadísticas del MINVU y busca el enlace "SU Otorgados 1990 a <Mes> <Año>".
2) Si el archivo cambió respecto de la última vez (se compara su huella SHA-256), lo descarga,
   regenera data/subsidios.json y las minutas Word.
3) Si no cambió, termina sin tocar nada.

Uso:
    python3 etl/actualizar.py                  # busca el archivo en la página del MINVU
    python3 etl/actualizar.py --archivo X.xlsx # usa un Excel que ya tienes (descarga manual)
    python3 etl/actualizar.py --forzar         # regenera aunque no haya cambios
"""
import argparse
import hashlib
import os
import re
import shutil
import sys
import tempfile
import urllib.parse
import urllib.request

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(RAIZ, "etl"))
import build_data      # noqa: E402
import build_minutas   # noqa: E402

PAGINA = "https://www.minvu.gob.cl/elementos-tecnicos/estadisticas/viviendas-contratadas-y-subsidios-otorgados/"
UA = {"User-Agent": "Mozilla/5.0 (compatible; CChC-OHiggins-Estudios/1.0)"}
HUELLA = os.path.join(RAIZ, "data", "fuente.sha256")
JSON_OUT = os.path.join(RAIZ, "data", "subsidios.json")
MINUTAS = os.path.join(RAIZ, "minutas")


def descargar(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=90) as r:
        return r.read()


def buscar_enlace():
    html = descargar(PAGINA).decode("utf-8", "ignore")
    enlaces = re.findall(r'href="([^"]*SU-Otorgados[^"]*\.xlsx)"', html, flags=re.I)
    if not enlaces:
        raise RuntimeError("No encontré el enlace al Excel 'SU Otorgados' en la página del MINVU. ¿Cambió la página?")
    return urllib.parse.urljoin(PAGINA, enlaces[0])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--archivo", help="ruta a un Excel ya descargado")
    ap.add_argument("--forzar", action="store_true")
    a = ap.parse_args()

    with tempfile.TemporaryDirectory() as tmp:
        if a.archivo:
            nombre = os.path.basename(a.archivo)
            ruta = os.path.join(tmp, nombre)
            shutil.copy(a.archivo, ruta)
            contenido = open(ruta, "rb").read()
        else:
            url = buscar_enlace()
            nombre = urllib.parse.unquote(os.path.basename(urllib.parse.urlparse(url).path))
            print("Enlace encontrado:", url)
            contenido = descargar(url)
            ruta = os.path.join(tmp, nombre)
            with open(ruta, "wb") as f:
                f.write(contenido)

        huella = hashlib.sha256(contenido).hexdigest()
        previa = open(HUELLA).read().strip() if os.path.exists(HUELLA) else ""
        if huella == previa and not a.forzar:
            print("Sin cambios: el Excel del MINVU es el mismo de la última actualización.")
            return 0

        print("Archivo nuevo o modificado:", nombre)
        codigo = build_data.main(ruta, JSON_OUT)
        if codigo != 0:
            print("La validación de datos reportó observaciones. No se actualizan las minutas ni la huella.")
            return codigo
        build_minutas.main(JSON_OUT, MINUTAS)
        with open(HUELLA, "w") as f:
            f.write(huella + "\n")
        print("Listo: datos y minutas actualizados.")
        return 0


if __name__ == "__main__":
    sys.exit(main())
