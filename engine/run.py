#!/usr/bin/env python3
# Orquestador del Visor GE. Reconstruye TODO desde los históricos (fuente de verdad),
# proyecta el mes en curso con el Avance de Ventas y publica en GitHub Pages.
#
# Uso: python3 run.py [--dry] [--force]
#   --dry   : construye index.html pero NO hace git push (para pruebas)
#   --force : ignora el marcador de las 3 huellas (publica aunque no todo esté fresco)
#
# Diseño clave: NO se guarda un "data_closed.json" incremental. Cada corrida vuelve a leer
# los históricos completos (ingest -> estrellas -> presupuesto/segmentación) y encima proyecta
# el mes en curso. Así el cierre de mes es AUTOMÁTICO: en cuanto aparece una hoja nueva en el
# histórico, ese mes pasa a "cerrado" (con estrellas recalculadas) y la proyección avanza al
# siguiente mes. No hay estado que se corrompa.
import os, sys, json, re, shutil, subprocess, calendar, hashlib, datetime
import openpyxl

ENGINE = os.path.dirname(os.path.abspath(__file__))
REPO   = os.path.dirname(ENGINE)
INPUTS = os.path.join(REPO, "inputs")      # xlsx descargados en runtime
WORK   = os.path.join(REPO, ".work")
DRY    = "--dry"   in sys.argv
FORCE  = "--force" in sys.argv

MES={'Ene':'01','Feb':'02','Mar':'03','Abr':'04','May':'05','Jun':'06',
     'Jul':'07','Ago':'08','Sep':'09','Oct':'10','Nov':'11','Dic':'12'}
def sheet_to_ym(name):
    m=re.match(r'([A-Za-z]{3})\s?(\d{2})$', str(name).strip())
    if not m or m.group(1).capitalize() not in MES: return None
    return f"20{m.group(2)}-{MES[m.group(1).capitalize()]}"

def next_month(ym):
    y,m=map(int,ym.split('-')); m+=1
    if m>12: y+=1; m=1
    return f"{y:04d}-{m:02d}"

def latest_hist_month():
    wb=openpyxl.load_workbook(os.path.join(INPUTS,"Historico_WM.xlsx"), read_only=True)
    yms=[sheet_to_ym(s) for s in wb.sheetnames]; yms=[y for y in yms if y]; wb.close()
    return max(yms)

def fingerprint():
    wb=openpyxl.load_workbook(os.path.join(INPUTS,"Avance_Ventas.xlsx"), data_only=True, read_only=True)
    fp={}
    for f,sh in {'BA':'BA','WM':'WM','SC':'SMC'}.items():
        ws=wb[sh]; s=0.0; n=0
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] is None: continue
            s+=(r[1] or 0)+(r[4] or 0); n+=1
        fp[f]=hashlib.md5(f"{n}:{round(s,2)}".encode()).hexdigest()[:12]
    wb.close(); return fp

def load_state():
    p=os.path.join(REPO,"state.json")
    return json.load(open(p)) if os.path.exists(p) else {"published_date":"","fp":{},"hist":""}

def run(cmd, cwd):
    r=subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    tag=" ".join(cmd)
    if r.returncode!=0:
        print("ERROR en", tag); print(r.stdout); print(r.stderr); sys.exit(1)
    last=[l for l in r.stdout.strip().splitlines() if l.strip()]
    if last: print("  ·", last[-1])
    return r.stdout

def build_index(work):
    visor=open(os.path.join(work,"Visor_GE_Walmart.html"),encoding="utf-8").read()
    frag=open(os.path.join(work,"gate_fragment.html"),encoding="utf-8").read()
    scr=open(os.path.join(work,"gate_script.html"),encoding="utf-8").read()
    hashes=json.load(open(os.path.join(work,"params.json")))["users"]
    scr=scr.replace("__HASHES__", json.dumps(hashes, ensure_ascii=False))
    bi=visor.rfind("<body"); m=re.match(r"<body[^>]*>", visor[bi:]); pos=bi+m.end()
    out=visor[:pos]+"\n"+frag+"\n"+visor[pos:]
    idx=out.rfind("</body>"); out=out[:idx]+scr+"\n"+out[idx:]
    open(os.path.join(work,"index.html"),"w",encoding="utf-8").write(out)
    return len(hashes)

def main():
    today=datetime.date.today()
    last_closed=latest_hist_month()            # última hoja del histórico = último mes CERRADO
    cur_month=next_month(last_closed)           # el mes en curso es el SIGUIENTE
    y,mn=map(int,cur_month.split('-')); dim=calendar.monthrange(y,mn)[1]
    this_cal=f"{today.year:04d}-{today.month:02d}"
    if cur_month==this_cal: asof=min(today.day,dim)   # mes calendario en curso -> proyecta
    else:                   asof=dim                  # mes ya transcurrido, aún sin hoja de cierre
    print(f"Último cierre: {last_closed} · mes en curso: {cur_month} · corte día {asof}/{dim}")

    st=load_state(); fp=fingerprint()
    fresh={f: fp[f]!=st["fp"].get(f) for f in fp}
    hist_changed=(last_closed!=st.get("hist",""))
    already=(st["published_date"]==today.isoformat())
    if not FORCE:
        if already and not hist_changed:
            print("Ya publiqué hoy y no hay cierre nuevo. Nada que hacer."); return
        if not all(fresh.values()) and not hist_changed:
            faltan=[f for f in fresh if not fresh[f]]
            print(f"Aún no están frescos los 3 formatos (faltan: {faltan}). No publico todavía."); return
    print("Cierre nuevo detectado." if hist_changed else "Marcador OK: los 3 formatos frescos.", "→ construyo y publico.")

    # preparar working dir
    if os.path.exists(WORK): shutil.rmtree(WORK)
    os.makedirs(WORK)
    for fn in os.listdir(ENGINE):
        src=os.path.join(ENGINE,fn)
        if fn!="run.py" and os.path.isfile(src): shutil.copy(src, os.path.join(WORK,fn))
    # inputs con los nombres que esperan los scripts
    for src in ["Historico_WM.xlsx","Historico_BA.xlsx","Historico_SC.xlsx"]:
        shutil.copy(os.path.join(INPUTS,src), os.path.join(WORK,src))
    shutil.copy(os.path.join(INPUTS,"Avance_Ventas.xlsx"), os.path.join(WORK,"Avance_Ventas.xlsx"))
    shutil.copy(os.path.join(INPUTS,"Parametros.xlsx"),    os.path.join(WORK,"Parametros.xlsx"))
    shutil.copy(os.path.join(INPUTS,"Segmentacion.xlsx"),  os.path.join(WORK,"Segmentacion_usuario.xlsx"))
    # librerías donde build_visor las busca
    for pkg,fn in [("chart.js","chart.umd.min.js"),("xlsx","xlsx.full.min.js")]:
        d=os.path.join(WORK,"node_modules",pkg,"dist"); os.makedirs(d, exist_ok=True)
        shutil.copy(os.path.join(WORK,fn), os.path.join(d,fn))

    print("Pipeline:")
    run(["python3","ingest.py",".","data.json"], WORK)              # 1) históricos cerrados -> data.json
    run(["python3","read_params.py","Parametros.xlsx"], WORK)       # 2) parámetros (fuente de verdad)
    run(["python3","stars.py","data.json","params.json"], WORK)     # 3) estrellas (sobre meses cerrados)
    run(["python3","recompute_all.py"], WORK)                       # 4) presupuesto + segmentación
    run(["python3","daily_update.py",cur_month,str(asof)], WORK)    # 5) proyección del mes en curso
    run(["python3","build_visor.py"], WORK)                         # 6) visor autocontenido
    nusers=build_index(WORK)                                        # 7) + gate de acceso
    print(f"  · index.html armado · {nusers} usuarios en el gate")

    # publicar
    shutil.copy(os.path.join(WORK,"index.html"), os.path.join(REPO,"index.html"))
    st={"published_date":today.isoformat(),"fp":fp,"hist":last_closed,"cur_month":cur_month,"asof":asof}
    json.dump(st, open(os.path.join(REPO,"state.json"),"w"), ensure_ascii=False, indent=1)
    if DRY:
        print("[--dry] index.html actualizado. No hago git push."); return
    run(["git","add","index.html","state.json"], REPO)
    run(["git","-c","user.name=Visor GE Bot","-c","user.email=ReyMon63@users.noreply.github.com",
         "commit","-q","-m",f"Actualización · {cur_month} corte {asof}/{dim}"], REPO)
    tok=os.environ.get("GH_TOKEN","")
    url=f"https://x-access-token:{tok}@github.com/ReyMon63/Dashboard.git" if tok else "origin"
    run(["git","push",url,"main"], REPO)
    print("Publicado en GitHub Pages.")

if __name__=="__main__":
    main()
