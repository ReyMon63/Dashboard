#!/usr/bin/env python3
# Convierte la descarga de Google Drive (JSON con 'content' en base64) a un .xlsx en disco.
# Uso: python3 decode_drive.py <archivo_json_de_la_descarga> <salida.xlsx>
# El JSON puede venir del resultado de la herramienta download_file_content
# (cuando es grande, la plataforma lo guarda en un archivo; pásale esa ruta).
import json, base64, sys
src=sys.argv[1]; out=sys.argv[2]
d=json.load(open(src, encoding="utf-8"))
data=d["content"] if isinstance(d, dict) else d
open(out,"wb").write(base64.b64decode(data))
print(f"OK · {out} ({len(base64.b64decode(data))} bytes) desde «{d.get('title','?') if isinstance(d,dict) else src}»")
