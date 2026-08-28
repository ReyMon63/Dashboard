# Visor GE — Conocimiento base del Proyecto

> Pega este documento en el Proyecto de Claude (sección de instrucciones/conocimiento).
> Sirve para que cualquier chat nuevo entienda el proyecto desde el primer mensaje.
> Última actualización: 27 de agosto de 2026.

## Qué es
El **Visor Ejecutivo de Garantías Extendidas (GE) — Grupo Walmart** es un tablero web (una sola página HTML) que muestra, por formato **WM (Walmart) / BA (Bodega Aurrerá) / SC (Sam's Club)**: ventas de GE, metas, presupuesto sugerido, clasificación de tiendas por **estrellas** (0–5) y ranking. Se arma con un pipeline de Python y se publica en GitHub Pages detrás de una pantalla de acceso.

Reymon es el dueño del proyecto y **no es programador**: todas las explicaciones deben ser en **español**, claras y sin tecnicismos innecesarios.

## Cómo se publica (importante)
- Un solo comando republica todo: **`./run_local.sh --force`** desde `/Users/reymon/Desktop/Proyectos/visor_bot`.
- **Debe correr "En tu computadora"** (Cowork local), NO en la nube: el build lee muchos archivos locales y hace `git push`; en la nube el puente falla con bloqueos de archivos.
- URL en línea (secreta): `reymon63.github.io/Dashboard/ge-d3792841/` · Login definido en la hoja «Usuarios» de Parámetros Base.

## Reglas de oro (para no romper nada)
1. **Antes de cambiar algo, consulta la Matriz Maestra** (el mapa de dependencias) y toca TODOS los lugares marcados «a mano».
2. **Nombres cortos de estrellas** (Vende · Crece · Cumple · Constancia · Eficiencia): viven en 6 lugares (dashboard_template.html, dashboard_ejecutivo_template.html, build_evaluacion.py, Instructivo en Drive + 2 copias en Descargas, y la app visitas-diarias). Cambiar uno = cambiar los 6.
3. **Conversión objetivo (WM 20% · BA 7% · SC 15%)**, pisos, crecimiento, alpha, implants y usuarios: se controlan desde **Parámetros Base.xlsx** (Drive). El cálculo se actualiza solo, pero los **textos** que mencionan esos números (ayuda del visor, Leyenda del Excel de Evaluación, Instructivo) se editan **a mano**.
4. **Excepción:** los **meses de evento (mayo y noviembre)** NO están en el Excel, están escritos en `engine/read_params.py`.
5. **Accesos (quién entra):** editar la hoja «Usuarios» de Parámetros Base y republicar. No se toca código.
6. **URL secreta:** si se cambia la palabra `ge-d3792841`, editar `PUBREL` en `run.py` y renombrar la carpeta; el enlace viejo deja de servir.

## Insumos y salidas (Google Drive)
- **Parámetros Base.xlsx** (Drive «Visor Ejecutivo») — panel de control.
- **Avance de Ventas.xlsx** (Descargas / Drive «Avance Ventas Bot») — mes en curso.
- **Segmentación.xlsx** (Drive «Avance Ventas Bot») — distritos/territoriales.
- **Históricos WM/BA/SC (con estrellas).xlsx** (Drive «Visor Ejecutivo») — se actualizan al cerrar mes.
- **Evaluación de Tiendas.xlsx** (Drive) — se regenera cada corrida.
- **data_closed.json** (repo) — memoria de meses cerrados (ene-2024 → jul-2026, 31 meses, 1 675 tiendas).

## Infraestructura y llaves
- Repo: `github.com/ReyMon63/Dashboard` (rama `main`). Token en `visor_bot/.gh_token` (respaldo en Drive: `visor_gh_token.txt`).
- Credenciales de Drive para subir la Evaluación: `Desktop/Proyectos/avance_ventas_bot/token.json` + `credentials.json`.

## Pendiente reservado (siguiente nivel)
- **Dominio propio + login real con Netlify** (Netlify Identity, invitación por correo). Guardado en la nota "PENDIENTE — Visor GE: dominio + login Netlify" y en la guía. Frase para reactivar: *"montemos el login de Netlify"*.

## Cómo trabajar con Claude en este proyecto
- Para **publicar o tocar archivos**, usar un chat **"En tu computadora"**.
- Para **planear o consultar**, cualquier chat sirve (incluido el iPhone).
- **Referencias vivas:** la *Matriz Maestra del Visor GE* (página interactiva + Excel) es el mapa oficial de dependencias. Consultarla antes de cualquier cambio.
- **Metodología (know-how):** `docs/Metodologia_Calculo_Visor_GE.html` guarda las **bases de cálculo y rangos** de lo que el visor decide solo (diagnóstico ejecutivo, padrón, etc.). Documento **interno**, no se comparte con el Cliente.
- **Los tres registros maestros están vinculados:** *Conocimiento Base* (el qué) · *Matriz Maestra* (el dónde) · *Metodología* (el cómo). Cada concepto lleva una **clave** igual en la Matriz y en la Metodología; la Matriz es el centro que enumera TODOS los lugares a tocar. Antes de cambiar algo, se consulta esa clave para ajustar código y documentos vinculados.
