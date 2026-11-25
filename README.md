# Subsidios otorgados MINVU · Región de O'Higgins

Dashboard interactivo y minutas descargables con los subsidios otorgados por el MINVU entre 1990 y 2026, a nivel nacional y por región.

Elaborado por el Área de Estudios de la CChC, Sede Región de O'Higgins.

**Fuente:** MINVU, Estadísticas → Viviendas contratadas y subsidios otorgados (archivo «SU Otorgados 1990 a <mes> <año>»).
https://www.minvu.gob.cl/elementos-tecnicos/estadisticas/viviendas-contratadas-y-subsidios-otorgados/

## Qué contiene

| Carpeta / archivo | Para qué sirve |
|---|---|
| `index.html`, `assets/` | La página web (dashboard). |
| `data/subsidios.json` | Datos ya procesados que lee la página. |
| `minutas/` | 17 minutas Word (país + 16 regiones), que se descargan desde la página. |
| `etl/` | Programas que leen el Excel del MINVU y generan los datos y las minutas. |
| `.github/workflows/actualizar.yml` | Automatización: revisa cada día si el MINVU publicó un Excel nuevo. |
| `fuente/` | Carpeta opcional para subir a mano un Excel nuevo. |

## Cómo publicarlo en GitHub (paso a paso)

1. Entra a https://github.com con la cuenta **cchc-ohiggins** y presiona **New repository**.
2. Nombre: `subsidios-ohiggins`. Déjalo **Public** (GitHub Pages gratuito lo requiere). No marques "Add a README". Presiona **Create repository**.
3. En la página del repositorio vacío, presiona **uploading an existing file**.
4. Descomprime este paquete y arrastra **todo su contenido** (las carpetas `assets`, `data`, `etl`, `minutas`, `fuente`, `.github` y los archivos sueltos) a la ventana. Presiona **Commit changes**.
   - Las carpetas que empiezan con punto (`.github`) a veces no se ven en el explorador: en Windows activa "Elementos ocultos"; en Mac usa Cmd+Shift+punto.
5. Ve a **Settings → Pages**. En *Source* elige **Deploy from a branch**, rama **main**, carpeta **/ (root)** y presiona **Save**.
6. A los 1–2 minutos la página queda en:
   **https://cchc-ohiggins.github.io/subsidios-ohiggins/**
7. Ve a **Settings → Actions → General → Workflow permissions** y marca **Read and write permissions**. Guarda. Esto permite que la actualización automática publique los datos nuevos.

## Cómo se actualiza

- **Automático:** el workflow se ejecuta todos los días. Si el Excel del MINVU cambió, regenera los datos y las 17 minutas y las publica solo. Si no cambió, no hace nada. El mes de corte («enero–octubre», «enero–noviembre»…) se lee del nombre del archivo, así que los textos se ajustan solos. Una vez al mes deja anotada la fecha de revisión, que aparece al pie de la página y evita que GitHub pause la automatización por inactividad.
- **Si algo falla:** GitHub envía un correo al dueño del repositorio cuando el workflow termina con error (por ejemplo, si el MINVU cambia su página o bloquea la descarga). En ese caso usa la opción de subir el Excel a `fuente/`.
- **Manual desde GitHub:** pestaña **Actions → Actualizar subsidios MINVU → Run workflow**.
- **Si el MINVU bloquea la descarga automática:** descarga el Excel desde su sitio y súbelo a la carpeta `fuente/` del repositorio (**Add file → Upload files**). Se actualiza solo.
- **Desde tu computador:** `pip install -r requirements.txt` y luego `python3 etl/actualizar.py --archivo ruta/al/excel.xlsx`.

## Probar en tu computador

La página necesita un servidor web, no se abre con doble clic. Dentro de la carpeta:

```
python3 -m http.server 8000
```

y abre http://localhost:8000 en el navegador.

## Notas metodológicas

- El archivo del MINVU trae subsidios (unidades) y montos en UF por programa, región y año. El último año es parcial (enero a la fecha de corte); el dashboard lo marca con trama y no lo usa en comparaciones de variación.
- Los montos se mantienen en UF. El MINVU advierte que la UF puede incluir incrementos de años anteriores. No se convierten a pesos porque el archivo no informa la fecha de otorgamiento de cada subsidio.
- El proceso valida que los programas sumen el total de cada región, que las regiones sumen el total país y que las hojas coincidan con el resumen por región. Si algo no cuadra, no se actualiza nada.
- Arica y Parinacota, Los Ríos y Ñuble figuran con datos propios desde su creación como regiones; antes están contados en la región de origen.
