import json
tpl=open('dashboard_template.html',encoding='utf-8').read()
chartjs=open('node_modules/chart.js/dist/chart.umd.min.js',encoding='utf-8').read()
sheetjs=open('node_modules/xlsx/dist/xlsx.full.min.js',encoding='utf-8').read()
data=open('data.json',encoding='utf-8').read()
out=tpl.replace('__CHARTJS__',chartjs).replace('__SHEETJS__',sheetjs).replace('__DATA__',data)
open('Visor_GE_Walmart.html','w',encoding='utf-8').write(out)
import os
print('Visor OK ·', round(os.path.getsize('Visor_GE_Walmart.html')/1024/1024,2),'MB')
# sanity: no leftover placeholders
for ph in ['__CHARTJS__','__SHEETJS__','__DATA__']:
    assert ph not in out, 'LEFTOVER '+ph
print('placeholders clean · budget3 refs:', out.count('BUD3'))
