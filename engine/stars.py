# Recalcula estrellas (levels/crit/newmark) desde DB['recs']. Reusa el modelo no-incremental trimestral.
from collections import defaultdict
def recompute_stars(DB, conv_target):
    stores=DB['stores']; recs=DB['recs']; months=DB['meta']['months']; NM=len(months)
    byS=defaultdict(dict)
    for si,mi,mo,ca,me,el in recs: byS[si][mi]=(mo,ca,me,el)
    START=months.index('2025-01'); FULL=months.index('2025-03')
    def fmt(si): return stores[si]['f']
    def win(si,a,b):
        ca=me=el=0.0;hits=0;mp=0
        for mi in range(a,b+1):
            if mi in byS[si]:
                _,c,m,e=byS[si][mi];ca+=max(c,0);me+=(m or 0);el+=max(e,0);mp+=1
                if (m or 0)>0 and c>=m:hits+=1
        return ca,me,el,hits,mp
    def ssum(si,a,b): return sum(byS[si].get(mi,(0,0,0,0))[1] for mi in range(a,b+1))
    def crit_at(si,m):
        f=fmt(si);ca,me,el,hits,mp=win(si,m-11,m)
        if mp==0: return None
        c1= ca>=12; c3= me>0 and ca>=me; c4= hits>=10; c5= el>0 and (ca/el)>=0.8*conv_target[f]
        prov=m<FULL; isnew=False
        if not prov:
            prevpres=any(mi in byS[si] for mi in range(m-14,m-11)); isnew=not prevpres
            c2= isnew or (ssum(si,m-2,m) >= ssum(si,m-14,m-12))
        else: c2=False
        mask=(1 if c1 else 0)|(2 if c2 else 0)|(4 if c3 else 0)|(8 if c4 else 0)|(16 if c5 else 0)
        cnt=(c1+c3+c4+c5) if prov else (c1+c2+c3+c4+c5)
        return cnt,mask,(1 if isnew else 0)
    levels={};crit={};newmark={}
    for si in byS:
        L=[];M=[];Nw=[]
        for m in range(START,NM):
            r=crit_at(si,m)
            if r is None: L.append(-1);M.append(0);Nw.append(0)
            else: L.append(r[0]);M.append(r[1]);Nw.append(r[2])
        levels[str(si)]=L;crit[str(si)]=M;newmark[str(si)]=Nw
    DB['stars']={'start_idx':START,'dic25_idx':FULL,'full_month':months[FULL],
                 'levels':levels,'crit':crit,'newmark':newmark,'model':'no-incremental','c2':'trimestral'}
    return DB

if __name__=="__main__":
    import json, sys
    data=sys.argv[1] if len(sys.argv)>1 else "data.json"
    params=sys.argv[2] if len(sys.argv)>2 else "params.json"
    DB=json.load(open(data)); P=json.load(open(params))
    DB=recompute_stars(DB, P["conv_target"])
    json.dump(DB, open(data,"w"), ensure_ascii=False, separators=(',',':'))
    lv=DB['stars']['levels']; import statistics
    last=len(DB['meta']['months'])-1-DB['stars']['start_idx']
    dist={}
    for si,L in lv.items():
        v=L[last] if 0<=last<len(L) else -1
        dist[v]=dist.get(v,0)+1
    print("Estrellas recalculadas · distribución último mes:", dict(sorted(dist.items())))
