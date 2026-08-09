# Genera "Evaluacion Tiendas.xlsx": por tienda, su nivel de estrellas, el detalle
# de los 5 criterios (checar/tachar) y la PROXIMA estrella + como lograrla.
# NO incluye segmentacion (eso vive en Segmentacion.xlsx). Una hoja por formato.
# Usa el ULTIMO MES CERRADO (no el proyectado).
#
# Uso: python3 build_evaluacion.py data.json params.json "Evaluacion Tiendas.xlsx" [MES_LABEL]
import json, sys, math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DATA   = sys.argv[1] if len(sys.argv) > 1 else "data.json"
PARAMS = sys.argv[2] if len(sys.argv) > 2 else "params.json"
OUT    = sys.argv[3] if len(sys.argv) > 3 else "Evaluacion Tiendas.xlsx"

DB = json.load(open(DATA))
try:
    P = json.load(open(PARAMS))
except Exception:
    P = {}
CT = P.get("conv_target") or {"WM": 0.20, "BA": 0.07, "SC": 0.15}

months = DB["meta"]["months"]
proj   = DB["meta"].get("projected")
mi     = (proj["mi"] - 1) if proj else len(months) - 1      # ultimo mes CERRADO
MES = {"01":"Ene","02":"Feb","03":"Mar","04":"Abr","05":"May","06":"Jun",
       "07":"Jul","08":"Ago","09":"Sep","10":"Oct","11":"Nov","12":"Dic"}
y, m = months[mi].split("-"); LABEL = MES[m] + y[2:]

S = DB["stars"]; START = S["start_idx"]; DIC25 = S["dic25_idx"]; crit = S["crit"]
stores = DB["stores"]
byidx = {s["i"]: s for s in stores}

# historia por tienda: hist[si][mi] = [monto, ge, meta, eleg]
hist = {}
for r in DB["recs"]:
    si, rmi, monto, ge, meta, eleg = r
    hist.setdefault(si, {})[rmi] = [monto, ge, meta, eleg]

def critbits(si, mm):
    a = crit.get(str(si))
    if not a: return None
    k = mm - START
    if k < 0 or k >= len(a) or a[k] < 0: return None
    return [(a[k] >> i) & 1 for i in range(5)]

PROXNAME = ["Venta mínima", "Crecimiento", "Meta", "Regularidad", "Penetración"]

def proxima(si, mm):
    f = byidx[si]["f"]; H = hist.get(si, {})
    def r(i): return H.get(i)
    def S3(a, b):
        s = 0
        for i in range(a, b + 1):
            rr = r(i)
            if rr: s += max(rr[1], 0)
        return s
    ca = me = el = 0; hits = 0
    for i in range(mm - 11, mm + 1):
        rr = r(i)
        if rr:
            ca += max(rr[1], 0); me += (rr[2] or 0); el += max(rr[3] or 0, 0)
            if (rr[2] or 0) > 0 and rr[1] >= rr[2]: hits += 1
    if ca == 0: return None
    conv = (ca / el) if el else 0; thr = 0.8 * CT[f]
    qNow = S3(mm - 2, mm); qPrev = S3(mm - 14, mm - 12)
    prevpres = any(r(i) for i in range(mm - 14, mm - 11))
    isnew = not prevpres
    cands = []
    if ca < 12: cands.append((ca/12, 0, f"Vender {math.ceil(12-ca)} GE más en 12 meses"))
    if mm >= DIC25 and not isnew and qPrev > 0 and qNow < qPrev:
        cands.append((qNow/qPrev, 1, f"Vender {math.ceil(qPrev-qNow)} GE más en el últ. trimestre vs. año previo"))
    if not (me > 0 and ca >= me):
        cands.append(((ca/me) if me else 0, 2, f"Vender {math.ceil(max(0, me-ca))} GE más para cumplir su meta acumulada"))
    if hits < 10: cands.append((hits/10, 3, f"Cumplir meta en {10-hits} mes(es) más"))
    if not (el > 0 and conv >= thr):
        cands.append(((conv/thr) if thr > 0 else 0, 4, f"Subir conversión a {thr*100:.1f}% (hoy {conv*100:.1f}%)"))
    if not cands: return {"max": True}
    cands.sort(key=lambda x: -x[0])
    return {"name": PROXNAME[cands[0][1]], "how": cands[0][2]}

# ---------- construir el libro ----------
hf = Font(bold=True, color="FFFFFF"); hfill = PatternFill("solid", fgColor="1F4E79")
ctr = Alignment(horizontal="center", vertical="center"); lft = Alignment(horizontal="left", vertical="center")
OKF = Font(color="1E7B34", bold=True); NOF = Font(color="C0392B", bold=True)
FNAME = {"WM": "Walmart", "BA": "Bodega Aurrerá", "SC": "Sam's Club"}
SHORT = ["≥12 GE", "Sin Decremento", "≥Meta Anual", "10 meses", "Conversión"]
HDR = ["Determinante", "Tienda", "Nivel"] + SHORT + ["Próxima estrella", "Cómo lograrla"]

wb = openpyxl.Workbook(); wb.remove(wb.active)
for f in ["WM", "BA", "SC"]:
    ws = wb.create_sheet(FNAME[f][:31])
    for c, h in enumerate(HDR, 1):
        cc = ws.cell(1, c, h); cc.font = hf; cc.fill = hfill; cc.alignment = ctr
    n = 0
    for s in stores:
        if s["f"] != f: continue
        b = critbits(s["i"], mi)
        if b is None: continue
        nivel = sum(b)
        pr = proxima(s["i"], mi)
        if pr is None:
            pname, phow = "—", "—"
        elif pr.get("max"):
            pname, phow = "Nivel máximo (5★)", "Mantener los 5 logros"
        else:
            pname, phow = pr["name"], pr["how"]
        ws.append([s["clave"], s["n"], nivel,
                   *["✓" if x else "✗" for x in b], pname, phow])
        rr = ws.max_row
        ws.cell(rr, 2).alignment = lft
        ws.cell(rr, 3).alignment = ctr
        for i in range(5):
            cell = ws.cell(rr, 4 + i); cell.alignment = ctr
            cell.font = OKF if b[i] else NOF
        ws.cell(rr, 9).alignment = lft; ws.cell(rr, 10).alignment = lft
        n += 1
    widths = [13, 28, 7, 10, 15, 12, 10, 12, 15, 52]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    print(f"{f}: {n} tiendas · hoja {FNAME[f]}")

# Leyenda
lg = wb.create_sheet("Leyenda")
rows = [
    [f"Evaluación de tiendas · nivel de estrellas y acciones · al cierre de {LABEL}"], [],
    ["Criterio (columna)", "Qué significa"],
    ["≥12 GE", "Venta mínima: vendió ≥12 GE en los últimos 12 meses"],
    ["Sin Decremento", "Crecimiento: no decreció vs. el mismo periodo del año previo"],
    ["≥Meta Anual", "Meta: cumplió la suma de sus metas de 12 meses"],
    ["10 meses", "Regularidad: cumplió su meta en al menos 10 meses"],
    ["Conversión", "Penetración: ≥80% de la conversión objetivo (WM 20% · BA 7% · SC 15%)"],
    [],
    ["✓", "Cumple el criterio"],
    ["✗", "No lo cumple"],
    ["Nivel", "Cantidad de criterios cumplidos (0 a 5)"],
    [],
    ["Próxima estrella", "El criterio prioritario más cercano de alcanzar"],
    ["Cómo lograrla", "La acción concreta para ganar esa estrella"],
    [],
    ["Nota", "Este archivo NO contiene segmentación (distritos/territoriales); eso vive en Segmentación.xlsx."],
]
for rr in rows: lg.append(rr)
lg.cell(1, 1).font = Font(bold=True, size=12)
for rn in (3,): lg.cell(rn, 1).font = Font(bold=True); lg.cell(rn, 2).font = Font(bold=True)
lg.column_dimensions["A"].width = 18; lg.column_dimensions["B"].width = 80

wb.save(OUT)
print(f"Listo · {OUT} · mes {LABEL}")
