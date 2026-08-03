# Motor del Visor Ejecutivo GE

Reconstruye el visor y lo publica en GitHub Pages. **Fuente de verdad = los históricos**
(no hay estado incremental que se corrompa). Cada corrida vuelve a leer todo y, encima,
proyecta el mes en curso con el Avance de Ventas.

## Cómo se ejecuta

```
python3 engine/run.py            # corrida normal (respeta el marcador y publica en GitHub)
python3 engine/run.py --dry      # arma index.html pero NO hace git push (pruebas)
python3 engine/run.py --force    # publica aunque no estén frescos los 3 formatos
```

Antes de correr hay que dejar los 6 insumos en `inputs/` (los descarga la Skill desde
Google Drive; ver `inputs_map.json`):

| Archivo en `inputs/` | Archivo en Drive | Carpeta |
|---|---|---|
| `Historico_WM.xlsx`  | Histórico WM (con estrellas).xlsx | Visor Ejecutivo |
| `Historico_BA.xlsx`  | Histórico BA (con estrellas).xlsx | Visor Ejecutivo |
| `Historico_SC.xlsx`  | Histórico SC (con estrellas).xlsx | Visor Ejecutivo |
| `Avance_Ventas.xlsx` | Avance de Ventas.xlsx | Avance Ventas Bot |
| `Parametros.xlsx`    | Parámetros Base.xlsx | Visor Ejecutivo |
| `Segmentacion.xlsx`  | Segmentación.xlsx | Avance Ventas Bot |

## Pipeline (lo que hace `run.py`)

1. **ingest.py** — lee los históricos (todas las hojas de mes que existan) → `data.json`
   con ventas/piezas/meta/elegibles por tienda y mes. Detecta el último mes **cerrado**.
2. **read_params.py** — lee `Parámetros Base` → `params.json` (crecimiento, α, pisos por
   formato, conversión objetivo, implants y las claves de acceso del gate).
3. **stars.py** — recalcula las estrellas (modelo no-incremental trimestral) sobre los
   meses cerrados.
4. **recompute_all.py** — construye el presupuesto «Histórico Suavizado» y la segmentación
   (territoriales/distritos + implant por formato) a partir de los parámetros.
5. **daily_update.py** — proyecta el **mes en curso** (el siguiente al último cerrado)
   linealmente por días transcurridos (factor = días del mes / día de corte). La meta NO
   se proyecta. Marca los datos como `projected`.
6. **build_visor.py** — inyecta datos + Chart.js + SheetJS en la plantilla → visor de un
   solo archivo.
7. **gate** (`gate_fragment.html` + `gate_script.html`) — antepone la pantalla de acceso
   con las claves de `Parámetros Base` → `index.html`.

## Marcador de publicación (`state.json`, en la raíz del repo)

Evita publicar a medias o repetido. Solo publica cuando:

- los **3 formatos** del Avance tienen datos nuevos respecto a la última publicación
  (huella MD5 por formato), **o**
- apareció una **hoja de mes nueva** en los históricos (cierre de mes), **y**
- no se publicó ya hoy.

Así la tarea se puede disparar varias veces al día sin riesgo: publica **una sola vez**,
cuando ya están los 3 formatos.

## Cierre de mes (automático)

No hay rutina especial. En cuanto el histórico gana una hoja nueva (p. ej. `Jul26`), ese
mes pasa a **cerrado** (con estrellas recalculadas) y la proyección avanza sola al mes
siguiente. Es la señal que controla el usuario: mientras no aparezca la hoja del mes
siguiente, se sigue proyectando el mes en curso.
