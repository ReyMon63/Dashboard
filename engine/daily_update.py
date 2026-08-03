# Motor DIARIO: proyecta Avance Ventas sobre el mes en curso y lo agrega/actualiza en data.json.
# NO recalcula estrellas/presupuesto/segmentación (eso es mensual/al cierre).
import json, openpyxl, calendar, sys
DB=json.load(open("data.json"))
months=DB['meta']['months']
stores=DB['stores']
look={}  # (f,clave)->si
for s in stores: look[(s['f'],int(s['clave']) if str(s['clave']).lstrip('-').isdigit() else s['clave'])]=s['i']

# --- parámetros de la corrida (en producción se derivan de la fecha/histórico) ---
CUR_MONTH = sys.argv[1] if len(sys.argv)>1 else "2026-07"   # mes en curso (última hoja del histórico)
ASOF_DAY  = int(sys.argv[2]) if len(sys.argv)>2 else 28      # día de corte con datos
y,m=map(int,CUR_MONTH.split('-')); dim=calendar.monthrange(y,m)[1]
factor=dim/ASOF_DAY
SHEET={'BA':'BA','WM':'WM','SC':'SMC'}

wb=openpyxl.load_workbook("Avance_Ventas.xlsx", data_only=True)
def to_int(v):
    try: return int(v)
    except: return None

# proyección por tienda
proj_recs=[]; agg={f:[0.0,0.0,0.0,0.0,0] for f in ['WM','BA','SC']}
matched=unmatched=0
for f,sh in SHEET.items():
    ws=wb[sh]
    for r in ws.iter_rows(min_row=2, values_only=True):
        clave=to_int(r[0])
        if clave is None: continue
        venta=(r[1] or 0); eleg=(r[2] or 0); plan=(r[3] or 0); ge=(r[4] or 0)
        si=look.get((f,clave))
        if si is None: unmatched+=1; continue
        matched+=1
        monto=venta*factor; cant=ge*factor; el=eleg*factor; meta=plan   # meta NO se proyecta
        proj_recs.append([si, monto, cant, meta, el])
        agg[f][0]+=monto; agg[f][1]+=cant; agg[f][2]+=meta; agg[f][3]+=el
        if ge>0: agg[f][4]+=1

# ¿el mes ya existe en data.json? -> actualizar; si no -> agregar
if CUR_MONTH in months:
    mi=months.index(CUR_MONTH)
    # limpiar recs viejos de ese mes
    DB['recs']=[rc for rc in DB['recs'] if rc[1]!=mi]
    for f in ['WM','BA','SC']: DB['agg'][f][mi]=[round(agg[f][i]) if i<4 else agg[f][4] for i in range(5)]
else:
    months.append(CUR_MONTH); mi=len(months)-1
    for f in ['WM','BA','SC']: DB['agg'][f].append([round(agg[f][i]) if i<4 else agg[f][4] for i in range(5)])
# agregar recs proyectados
for si,monto,cant,meta,el in proj_recs:
    DB['recs'].append([si, mi, round(monto), round(cant), round(meta), round(el)])

DB['meta']['months']=months
DB['meta']['projected']={'mi':mi,'month':CUR_MONTH,'asof_day':ASOF_DAY,'days_in_month':dim,'factor':round(factor,4)}
json.dump(DB, open("data.json","w"), ensure_ascii=False, separators=(',',':'))
tot=sum(agg[f][0] for f in agg)
print(f"Mes en curso: {CUR_MONTH} (índice {mi}) · corte día {ASOF_DAY}/{dim} · factor ×{factor:.3f}")
print(f"Tiendas cruzadas: {matched} (sin cruce {unmatched})")
for f in ['WM','BA','SC']:
    print(f"  {f}: monto proy {agg[f][0]:,.0f} · GE proy {agg[f][1]:,.0f} · tiendas {agg[f][4]}")
print(f"  GRUPO monto proyectado: {tot:,.0f}")
print("data.json actualizado. NM =", len(months))
