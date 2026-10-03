import os, io, re, glob, hashlib, time, datetime as dt
import pandas as pd
import uvicorn
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, Response, RedirectResponse
from fastapi.staticfiles import StaticFiles

BASE = os.path.dirname(os.path.abspath(__file__))  # rutas fijas, sin depender de desde dónde se arranque
EXCEL_PATH = os.path.join(BASE, "Control_Inventario_Pole_Dance.xlsx")
os.makedirs(os.path.join(BASE, "static"), exist_ok=True)
# Las fotos de productos se guardan en disco (persisten al reiniciar y al hacer git pull).
FOTOS_DIR = os.environ.get("FOTOS_DIR", os.path.join(BASE, "static", "fotos"))
os.makedirs(FOTOS_DIR, exist_ok=True)


def buscar_logo():
    """Busca logo.png/jpg/webp en la carpeta de la app o en static/."""
    for d in (BASE, os.path.join(BASE, "static")):
        for f in sorted(glob.glob(os.path.join(d, "[Ll]ogo*"))):
            if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
                return f
    return None


app = FastAPI(title="Pole Dance Rojas Sport", version="10.0.0")
app.mount("/static", StaticFiles(directory=os.path.join(BASE, "static")), name="static")
_cache = {"mtime": None, "data": None}


def limpio(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return str(v)[:10]
    if hasattr(v, "item"):
        v = v.item()
    return None if (isinstance(v, float) and v != v) else v


def leer_excel():
    """Lee TODAS las hojas (valores calculados). Nunca modifica el Excel."""
    m = os.path.getmtime(EXCEL_PATH)
    if _cache["mtime"] == m:
        return _cache["data"]
    hojas = {}
    libres = ("contable", "dashboard", "dinámico", "dinamico", "configuraci")
    for nombre, df in pd.read_excel(EXCEL_PATH, sheet_name=None, header=None).items():
        df = df.dropna(how="all").dropna(axis=1, how="all")
        filas = [[limpio(x) for x in r] for r in df.values.tolist()]
        if any(k in nombre.lower() for k in libres):
            hojas[nombre] = {"libre": True, "cols": [], "filas": filas}
        elif filas:
            hojas[nombre] = {"libre": False, "cols": [str(c) for c in filas[0]], "filas": filas[1:]}
    _cache.update(mtime=m, data=hojas)
    return hojas


@app.get("/api/datos")
def datos():
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(404, "No se encuentra el Excel. Súbelo con el botón ⬆️.")
    return JSONResponse({"hojas": leer_excel()})


@app.post("/subir-excel")
async def subir_excel(archivo: UploadFile = File(...)):
    if not archivo.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Sube un archivo .xlsx")
    with open(EXCEL_PATH, "wb") as f:
        f.write(await archivo.read())
    _cache["mtime"] = None
    return {"ok": True}


@app.get("/descargar-excel")
def descargar_excel():
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(404, "El Excel no existe")
    return FileResponse(EXCEL_PATH, filename="Control_Inventario_Pole_Dance.xlsx")


# ---------- Fotos de productos (una por ID de entrada, ej. ENT-001) ----------
# Con las variables CLOUDINARY_* definidas, las fotos se guardan en Cloudinary (gratis y
# permanente). Sin ellas se guardan en disco local, que en Render gratis se borra.
CLD = (os.environ.get("CLOUDINARY_CLOUD_NAME"), os.environ.get("CLOUDINARY_API_KEY"),
       os.environ.get("CLOUDINARY_API_SECRET"))
USA_CLOUD = all(CLD)


def _pid(pid):
    pid = re.sub(r"[^A-Za-z0-9_-]", "", pid)
    if not pid:
        raise HTTPException(400, "ID inválido")
    return pid


def _jpeg(data: bytes) -> bytes:
    """Reduce la foto a máx. 1000px (JPEG) para ahorrar espacio."""
    try:
        from PIL import Image, ImageOps
        img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
        img.thumbnail((1000, 1000))
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=85)
        return buf.getvalue()
    except ImportError:
        return data
    except Exception:
        raise HTTPException(400, "Imagen no válida")


def _cld_post(accion, params, files=None):
    import requests
    params = {**params, "timestamp": int(time.time())}
    base = "&".join(f"{k}={v}" for k, v in sorted(params.items()))
    firma = hashlib.sha1((base + CLD[2]).encode()).hexdigest()
    r = requests.post(f"https://api.cloudinary.com/v1_1/{CLD[0]}/image/{accion}",
                      data={**params, "api_key": CLD[1], "signature": firma}, files=files, timeout=30)
    if not r.ok:
        raise HTTPException(502, "Error con Cloudinary: " + r.text[:200])


@app.get("/api/fotos")
def lista_fotos():
    if USA_CLOUD:
        import requests
        ids, cursor = [], None
        while True:
            p = {"prefix": "polesport/", "max_results": 500}
            if cursor:
                p["next_cursor"] = cursor
            r = requests.get(f"https://api.cloudinary.com/v1_1/{CLD[0]}/resources/image/upload",
                             params=p, auth=(CLD[1], CLD[2]), timeout=30)
            if not r.ok:
                return ids
            j = r.json()
            ids += [x["public_id"].split("/", 1)[1] for x in j.get("resources", [])]
            cursor = j.get("next_cursor")
            if not cursor:
                return ids
    return [os.path.splitext(f)[0] for f in os.listdir(FOTOS_DIR) if f.endswith(".jpg")]


@app.get("/foto/{pid}")
def ver_foto(pid: str, w: int = 0):
    pid = _pid(pid)
    if USA_CLOUD:
        t = f"c_fill,w_{w},h_{w},q_auto,f_auto/" if w else "q_auto,f_auto/"
        return RedirectResponse(f"https://res.cloudinary.com/{CLD[0]}/image/upload/{t}polesport/{pid}")
    ruta = os.path.join(FOTOS_DIR, pid + ".jpg")
    if not os.path.exists(ruta):
        raise HTTPException(404, "Sin foto")
    return FileResponse(ruta, media_type="image/jpeg")


@app.post("/fotos/{pid}")
async def subir_foto(pid: str, archivo: UploadFile = File(...)):
    pid = _pid(pid)
    data = await archivo.read()
    if len(data) > 20 * 1024 * 1024:
        raise HTTPException(413, "Foto demasiado grande")
    data = _jpeg(data)
    if USA_CLOUD:
        _cld_post("upload", {"public_id": f"polesport/{pid}", "overwrite": "true", "invalidate": "true"},
                  files={"file": (pid + ".jpg", data)})
    else:
        with open(os.path.join(FOTOS_DIR, pid + ".jpg"), "wb") as f:
            f.write(data)
    return {"ok": True}


@app.delete("/fotos/{pid}")
def borrar_foto(pid: str):
    pid = _pid(pid)
    if USA_CLOUD:
        _cld_post("destroy", {"public_id": f"polesport/{pid}", "invalidate": "true"})
    else:
        ruta = os.path.join(FOTOS_DIR, pid + ".jpg")
        if os.path.exists(ruta):
            os.remove(ruta)
    return {"ok": True}


# ---------- Logo como ícono de la app (PWA) ----------
@app.get("/icono/{n}.png")
def icono(n: int):
    logo = buscar_logo()
    if not logo:
        raise HTTPException(404, "No se encontró el logo (logo.png)")
    try:
        from PIL import Image
        img = Image.open(logo).convert("RGBA")
        img.thumbnail((int(n * 0.78), int(n * 0.78)))  # logo completo, con margen
        fondo = Image.new("RGBA", (n, n), (255, 255, 255, 255))
        fondo.paste(img, ((n - img.width) // 2, (n - img.height) // 2), img)
        buf = io.BytesIO()
        fondo.convert("RGB").save(buf, "PNG")
        return Response(buf.getvalue(), media_type="image/png")
    except ImportError:
        return FileResponse(logo)


@app.get("/manifest.webmanifest")
def manifest():
    return JSONResponse({
        "name": "Pole Dance Rojas Sport", "short_name": "Pole Sport",
        "start_url": "/", "scope": "/", "display": "standalone",
        "background_color": "#ffffff", "theme_color": "#070b24",
        "icons": [
            {"src": "/icono/192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "/icono/512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"},
        ]}, media_type="application/manifest+json")


@app.get("/sw.js")
def sw():
    return Response("self.addEventListener('install',e=>self.skipWaiting());"
                    "self.addEventListener('fetch',()=>{});", media_type="application/javascript")


@app.get("/", response_class=HTMLResponse)
def interfaz():
    return HTML


HTML = r'''<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Pole Dance Rojas Sport</title>
<link rel="manifest" href="/manifest.webmanifest"><meta name="theme-color" content="#070b24">
<link rel="icon" href="/icono/192.png"><link rel="apple-touch-icon" href="/icono/180.png">
<meta name="apple-mobile-web-app-capable" content="yes"><meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Pole Sport">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<style>
:root{--p:#22a94b;--a:#1e3a8a;--ok:#10b981;--bad:#ef4444;--warn:#f59e0b;--tx:#1e1b4b;--mu:#64748b}
*{box-sizing:border-box;margin:0;padding:0;font-family:system-ui,-apple-system,"Segoe UI",sans-serif}
body{background:linear-gradient(135deg,#070b24,#0f1a4a 55%,#123a3a) fixed;min-height:100vh;color:var(--tx);
padding:env(safe-area-inset-top) 12px 40px}
.w{max-width:1100px;margin:0 auto}
header{display:flex;align-items:center;gap:14px;padding:16px 4px;color:#fff;flex-wrap:wrap}
header img{width:56px;height:56px;border-radius:14px;object-fit:contain;background:#fff;padding:4px;box-shadow:0 0 20px #22c55e55}
header h1{font-size:1.25rem;font-weight:800;background:linear-gradient(135deg,#fff,#86efac,#22c55e);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
header small{color:#cbd5e1;display:block}.sp{flex:1}
.bt{background:#ffffff22;color:#fff;border:1px solid #ffffff44;padding:8px 14px;border-radius:30px;font-weight:600;font-size:.8rem;cursor:pointer;text-decoration:none}
nav{display:flex;gap:6px;overflow-x:auto;padding:6px;background:#00000040;border-radius:40px;margin-bottom:18px;position:sticky;top:0;z-index:5;backdrop-filter:blur(8px)}
nav button{border:0;background:none;color:#a5b4c8;padding:10px 16px;border-radius:30px;font-weight:700;font-size:.85rem;white-space:nowrap;cursor:pointer}
nav button.on{background:linear-gradient(135deg,var(--p),var(--a));color:#fff}
.g{display:grid;gap:14px}.k{grid-template-columns:repeat(auto-fit,minmax(150px,1fr));margin-bottom:16px}
.c2{grid-template-columns:repeat(auto-fit,minmax(310px,1fr))}
.card{background:#fffffff5;border-radius:18px;padding:18px;box-shadow:0 10px 30px -8px #0006}
.kpi{cursor:pointer;text-align:center;transition:.2s}.kpi:hover{transform:translateY(-3px)}
.kpi h4{font-size:.7rem;text-transform:uppercase;letter-spacing:.6px;color:var(--mu)}
.kpi b{font-size:1.45rem;display:block;margin-top:4px}
.card h3{font-size:1rem;margin-bottom:12px}
.br{display:flex;align-items:center;gap:8px;margin:7px 0;font-size:.82rem;cursor:pointer;border-radius:8px;padding:2px}
.br:hover{background:#f0fdf4}.br span:first-child{width:38%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.br i{display:block;height:14px;border-radius:7px;background:linear-gradient(90deg,var(--p),var(--a));min-width:3px}
.br em{font-style:normal;font-weight:700;margin-left:auto;white-space:nowrap}
.mes{display:flex;align-items:flex-end;gap:10px;height:170px;padding-top:10px;overflow-x:auto}
.mes div{flex:1;min-width:54px;text-align:center;font-size:.7rem}.mes .bb{display:flex;align-items:flex-end;gap:3px;height:130px;justify-content:center}
.mes .bb i{width:20px;border-radius:5px 5px 0 0}
.tb{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}
.tb input,.tb select{padding:10px 14px;border:1.5px solid #cbd5e1;border-radius:30px;font-size:.9rem;background:#f8fafc;outline:none;flex:1;min-width:150px}
.tw{overflow:auto;max-height:68vh;border:1px solid #e2e8f0;border-radius:12px}
table{border-collapse:collapse;width:100%;background:#fff}
th{position:sticky;top:0;background:#f8fafc;font-size:.7rem;text-transform:uppercase;color:#475569;padding:10px 12px;text-align:left;cursor:pointer;white-space:nowrap;border-bottom:1px solid #e2e8f0}
td{padding:8px 12px;border-bottom:1px solid #f1f5f9;font-size:.85rem;white-space:nowrap}
tr:hover td{background:#f0fdf4}
.chip{padding:3px 10px;border-radius:20px;font-size:.72rem;font-weight:700;background:#dcfce7;color:#166534}
.neg{color:var(--bad);font-weight:700}.pos{color:var(--ok);font-weight:700}
.mu{color:var(--mu);font-size:.8rem;margin:8px 2px}.more{width:100%;margin-top:10px;padding:10px;border:1.5px dashed #cbd5e1;border-radius:10px;background:#f8fafc;cursor:pointer;font-weight:700;color:var(--mu)}
</style></head><body><div class="w">
<header><img src="/icono/192.png" alt="" onerror="this.style.display='none'">
<div><h1>Pole Dance Rojas Sport</h1><small id="sub">Cargando…</small></div><div class="sp"></div>
<a class="bt" href="/descargar-excel">⬇️ Excel</a><button class="bt" onclick="cargar()">🔄</button></header>
<nav id="nav"></nav><main id="main"></main></div>
<script>
if('serviceWorker'in navigator)navigator.serviceWorker.register('/sw.js');
const T={resumen:{n:'📊 Resumen'},alumnas:{n:'🎓 Alumnas',kw:'alumnas',req:'Nombre',fc:'Estado'},
ingresos:{n:'🤸 Ingresos',kw:'ingresos',req:'Alumna',fc:'Tipo'},gastos:{n:'💸 Gastos',kw:'gastos',req:'Valor ($)',fc:'Categoría'},
stock:{n:'👗 Stock'},merc:{n:'📦 Entradas',kw:'marcancia',req:'Código Prenda',fc:'Clasificación'},
activos:{n:'🏢 Activos',kw:'activos',req:'Nombre del Activo',fc:'Estado'},ventas:{n:'🛒 Ventas',kw:'ventas',req:'Código Prenda',fc:'Método Pago'},
cont:{n:'📑 Contable',kw:'contable'}};
let D={},tab='resumen',S={},F=new Map();
const $=s=>document.querySelector(s),esc=v=>String(v??'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const fmt=n=>'$'+Math.round(n||0).toLocaleString('es-CO'),num=v=>Number(v)||0;
const hoja=kw=>{const k=Object.keys(D).find(n=>n.toLowerCase().includes(kw));return k?D[k]:null};
function tabla(id){const t=T[id];if(id==='stock')return stock();const h=hoja(t.kw);if(!h)return{cols:[],rows:[]};
const rows=h.filas.map(f=>Object.fromEntries(h.cols.map((c,i)=>[c,f[i]]))).filter(r=>r[t.req]!=null&&r[t.req]!=='');return{cols:h.cols,rows}}
function stock(){const m={};tabla('merc').rows.forEach(r=>{const k=String(r['Código Prenda']).trim().toUpperCase();
const x=m[k]??={'Código':k,Entradas:0,Vendidas:0,Stock:0,'Costo total ($)':0};x.Entradas+=num(r['Cantidad (+)']);x['Costo total ($)']+=num(r['Costo Total ($)'])});
tabla('ventas').rows.forEach(r=>{const k=String(r['Código Prenda']).trim().toUpperCase();if(m[k])m[k].Vendidas+=num(r['Cant. Vendida (-)'])});
const rows=Object.values(m).map(x=>({...x,Stock:x.Entradas-x.Vendidas}));return{cols:['Código','Entradas','Vendidas','Stock','Costo total ($)'],rows}}
const sum=(a,c)=>a.reduce((s,r)=>s+num(r[c]),0);
function group(rows,key,val){const m={};rows.forEach(r=>{const k=r[key]||'(sin dato)';m[k]=(m[k]||0)+(val?num(r[val]):1)});
return Object.entries(m).sort((a,b)=>b[1]-a[1])}
function barras(items,id,col,money){const mx=Math.max(...items.map(i=>i[1]),1);
return items.slice(0,8).map(([k,v])=>`<div class="br" onclick="ir('${id}','${col}',this.dataset.k)" data-k="${esc(k)}"><span>${esc(k)}</span>
<i style="width:${v/mx*45}%"></i><em>${money?fmt(v):v}</em></div>`).join('')||'<p class="mu">Sin datos</p>'}
function ir(id,col,val){S[id]={q:'',f:val&&col?val:'',sc:null,asc:1,n:100};tab=id;pintar()}
function resumen(){const al=tabla('alumnas').rows,ing=tabla('ingresos').rows,ga=tabla('gastos').rows,ac=tabla('activos').rows,st=tabla('stock').rows;
const I=sum(ing,'Total ($)'),G=sum(ga,'Valor ($)'),nuevas=al.filter(a=>a.Estado==='Nueva').length,rec=al.filter(a=>a.Estado==='Recurrente').length;
const meses={};ing.forEach(r=>{const m=String(r.Fecha||'').slice(0,7);if(m)(meses[m]??={i:0,g:0}).i+=num(r['Total ($)'])});
ga.forEach(r=>{const m=String(r.Fecha||'').slice(0,7);if(m)(meses[m]??={i:0,g:0}).g+=num(r['Valor ($)'])});
const ms=Object.keys(meses).sort(),mx=Math.max(...ms.map(m=>Math.max(meses[m].i,meses[m].g)),1);
const K=(t,v,id,c='')=>`<div class="card kpi" onclick="ir('${id}')"><h4>${t}</h4><b class="${c}">${v}</b></div>`;
return `<div class="g k">${K('Ingresos',fmt(I),'ingresos','pos')}${K('Gastos',fmt(G),'gastos','neg')}${K('Resultado',fmt(I-G),'cont',I-G>=0?'pos':'neg')}
${K('Alumnas',al.length,'alumnas')}${K('Nuevas / Recurrentes',nuevas+' / '+rec,'alumnas')}${K('Prendas en stock',sum(st,'Stock'),'stock')}${K('Valor activos',fmt(sum(ac,'Valor Total ($)')),'activos')}</div>
<div class="g c2"><div class="card" style="grid-column:1/-1"><h3>📅 Ingresos vs Gastos por mes <small class="mu">(🟣 ingresos · 🔴 gastos)</small></h3><div class="mes">
${ms.map(m=>`<div><div class="bb"><i title="${fmt(meses[m].i)}" style="height:${meses[m].i/mx*100}%;background:var(--p)"></i><i title="${fmt(meses[m].g)}" style="height:${meses[m].g/mx*100}%;background:var(--bad)"></i></div>${m}</div>`).join('')}</div></div>
<div class="card"><h3>💸 Gastos por categoría</h3>${barras(group(ga,'Categoría','Valor ($)'),'gastos','Categoría',1)}</div>
<div class="card"><h3>🤸 Ingresos por tipo</h3>${barras(group(ing,'Tipo','Total ($)'),'ingresos','Tipo',1)}</div>
<div class="card"><h3>💳 Ingresos por método de pago</h3>${barras(group(ing,'Método de Pago','Total ($)'),'ingresos','Método de Pago',1)}</div>
<div class="card"><h3>🎓 Alumnas por estado</h3>${barras(group(al,'Estado'),'alumnas','Estado')}</div>
<div class="card"><h3>🏆 Alumnas que más han pagado</h3>${barras(al.map(a=>[a.Nombre,num(a['Total pagado ($)'])]).sort((a,b)=>b[1]-a[1]),'alumnas','Nombre',1)}</div>
<div class="card"><h3>🏢 Activos por estado</h3>${barras(group(ac,'Estado','Valor Total ($)'),'activos','Estado',1)}</div></div>
<p class="mu">Toca cualquier tarjeta o barra para ver el detalle filtrado.</p>`}
function vista(id){const t=T[id],s=S[id]??={q:'',f:'',sc:null,asc:1,n:100};
if(id==='cont'){const h=hoja('contable');return `<div class="card"><div class="tw"><table>${(h?h.filas:[]).map(f=>`<tr>${f.map(c=>`<td>${typeof c==='number'?fmt(c):esc(c)}</td>`).join('')}</tr>`).join('')}</table></div></div>`}
let {cols,rows}=tabla(id);if(id==='merc')cols=['Foto',...cols];const fc=t.fc;
const opts=fc&&cols.includes(fc)?[...new Set(rows.map(r=>r[fc]).filter(v=>v!=null))].sort():[];
let r=rows.filter(x=>(!s.f||String(x[fc])===s.f)&&(!s.q||Object.values(x).join(' ').toLowerCase().includes(s.q.toLowerCase())));
if(s.sc)r.sort((a,b)=>(a[s.sc]>b[s.sc]?1:-1)*s.asc);
const cell=(c,v,r)=>c==='Foto'?foto(r['ID Entrada']):v==null?'':typeof v==='number'&&/\(\$\)|precio|valor/i.test(c)?fmt(v):c==='Estado'||c===fc&&c!=='Nombre'?`<span class="chip">${esc(v)}</span>`:esc(v);
return `<div class="card"><div class="tb"><input placeholder="🔍 Buscar en ${t.n}…" value="${esc(s.q)}" oninput="S['${id}'].q=this.value;S['${id}'].n=100;repintar('${id}')" id="q">
${opts.length?`<select onchange="S['${id}'].f=this.value;pintar()"><option value="">Todos (${fc})</option>${opts.map(o=>`<option ${o===s.f?'selected':''}>${esc(o)}</option>`).join('')}</select>`:''}</div>
<div class="mu">${r.length} de ${rows.length} registros${id==='ingresos'||id==='gastos'?' · Total filtrado: <b>'+fmt(sum(r,id==='gastos'?'Valor ($)':'Total ($)'))+'</b>':''}</div>
<div class="tw"><table><thead><tr>${cols.map(c=>`<th onclick="orden('${id}',this.dataset.c)" data-c="${esc(c)}">${esc(c)}${s.sc===c?(s.asc>0?' ▲':' ▼'):''}</th>`).join('')}</tr></thead>
<tbody>${r.slice(0,s.n).map(x=>`<tr>${cols.map(c=>`<td>${cell(c,x[c],x)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
${r.length>s.n?`<button class="more" onclick="S['${id}'].n+=200;pintar()">Ver más</button>`:''}</div>`}
const foto=id=>`<span style="display:flex;gap:6px;align-items:center">${F.has(id)?`<img src="/foto/${id}?w=120&t=${F.get(id)}" style="width:44px;height:44px;border-radius:8px;object-fit:cover;cursor:zoom-in" onclick="zoom('${id}')">`:''}<label class="chip" style="cursor:pointer">📷${F.has(id)?'':' Subir'}<input type="file" accept="image/*" hidden onchange="subirFoto('${id}',this.files[0])"></label></span>`;
function zoom(id){const d=document.createElement('div');d.style.cssText='position:fixed;inset:0;background:#000d;display:flex;align-items:center;justify-content:center;z-index:99;cursor:zoom-out;padding:20px';
d.innerHTML=`<img src="/foto/${id}?t=${F.get(id)}" style="max-width:100%;max-height:100%;border-radius:12px">`;d.onclick=()=>d.remove();document.body.appendChild(d)}
async function subirFoto(id,f){if(!f)return;const fd=new FormData();fd.append('archivo',f);const r=await fetch('/fotos/'+id,{method:'POST',body:fd});
if(r.ok){F.set(id,Date.now());pintar()}else alert('No se pudo subir la foto')}
function orden(id,c){const s=S[id];s.asc=s.sc===c?-s.asc:1;s.sc=c;pintar()}
function repintar(id){const p=$('#q').selectionStart;pintar();const q=$('#q');q.focus();q.setSelectionRange(p,p)}
function pintar(){$('#nav').innerHTML=Object.entries(T).map(([k,t])=>`<button class="${k===tab?'on':''}" onclick="tab='${k}';pintar()">${t.n}</button>`).join('');
$('#main').innerHTML=tab==='resumen'?resumen():vista(tab)}
async function cargar(){try{const r=await fetch('/api/datos');if(!r.ok)throw new Error((await r.json()).detail);D=(await r.json()).hojas;try{F=new Map((await(await fetch('/api/fotos')).json()).map(i=>[i,1]))}catch(e){}
$('#sub').textContent='Inventario, alumnas y finanzas · '+new Date().toLocaleDateString('es-CO');pintar()}catch(e){$('#main').innerHTML='<div class="card">⚠️ '+esc(e.message)+'</div>'}}
async function subir(f){if(!f)return;const fd=new FormData();fd.append('archivo',f);const r=await fetch('/subir-excel',{method:'POST',body:fd});r.ok?cargar():alert('No se pudo subir el Excel')}
cargar();
</script></body></html>'''

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)