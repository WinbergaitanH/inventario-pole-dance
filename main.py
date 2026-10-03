"""Pole Dance Rojas Sport — panel de inventario, alumnas y finanzas (v11).

Variables de entorno (en Render → Environment):
  APP_CLAVE              Clave para entrar a la app (si no existe, la app queda abierta).
  CLOUDINARY_CLOUD_NAME  ┐
  CLOUDINARY_API_KEY     ├ "Memoria" permanente: fotos y Excel se guardan en Cloudinary (gratis).
  CLOUDINARY_API_SECRET  ┘ Sin ellas todo se guarda en disco, que en Render gratis se borra.
"""
import os, io, re, glob, hashlib, hmac, time, datetime as dt
from contextlib import asynccontextmanager
import pandas as pd
import uvicorn
from fastapi import FastAPI, HTTPException, UploadFile, File, Request, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, Response, RedirectResponse

BASE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(BASE, "static")
EXCEL_PATH = os.path.join(BASE, "Control_Inventario_Pole_Dance.xlsx")
FOTOS_DIR = os.environ.get("FOTOS_DIR", os.path.join(STATIC, "fotos"))
os.makedirs(FOTOS_DIR, exist_ok=True)

CLAVE = os.environ.get("APP_CLAVE", "").strip()
CLD = (os.environ.get("CLOUDINARY_CLOUD_NAME"), os.environ.get("CLOUDINARY_API_KEY"),
       os.environ.get("CLOUDINARY_API_SECRET"))
USA_CLOUD = all(CLD)
# Carpeta "secreta" del Excel en Cloudinary (derivada del API secret: nadie más puede adivinar la ruta)
EXCEL_CLD = f"polesport/excel/{hashlib.sha1((CLD[2] or '').encode()).hexdigest()[:20]}" if USA_CLOUD else ""
_cache = {"mtime": None, "data": None}
_estado = {"excel_origen": "local", "excel_fecha": None}


# =====================================================================  Cloudinary
def _firma(params):
    base = "&".join(f"{k}={v}" for k, v in sorted(params.items()))
    return hashlib.sha1((base + CLD[2]).encode()).hexdigest()


def _cld_post(tipo, accion, params, files=None):
    import requests
    params = {**params, "timestamp": int(time.time())}
    r = requests.post(f"https://api.cloudinary.com/v1_1/{CLD[0]}/{tipo}/{accion}",
                      data={**params, "api_key": CLD[1], "signature": _firma(params)}, files=files, timeout=60)
    if not r.ok:
        raise HTTPException(502, "Error con Cloudinary: " + r.text[:200])
    return r.json()


def excel_a_nube(data: bytes):
    """Guarda el Excel como 'actual' y además una copia de respaldo con fecha."""
    if not USA_CLOUD:
        return
    sello = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    for pid in (f"{EXCEL_CLD}/actual.xlsx", f"{EXCEL_CLD}/respaldos/{sello}.xlsx"):
        _cld_post("raw", "upload", {"public_id": pid, "overwrite": "true", "invalidate": "true"},
                  files={"file": ("excel.xlsx", data)})


def excel_desde_nube():
    """Al arrancar, trae el último Excel guardado en Cloudinary (si existe)."""
    if not USA_CLOUD:
        return
    import requests
    try:
        r = requests.get(f"https://api.cloudinary.com/v1_1/{CLD[0]}/resources/raw/upload/{EXCEL_CLD}/actual.xlsx",
                         auth=(CLD[1], CLD[2]), timeout=30)
        if not r.ok:
            return  # todavía no se ha subido ningún Excel: se usa el del repositorio
        info = r.json()
        x = requests.get(info["secure_url"], timeout=60)  # URL con versión: siempre la última
        if x.ok and x.content[:2] == b"PK":
            with open(EXCEL_PATH, "wb") as f:
                f.write(x.content)
            _estado.update(excel_origen="nube", excel_fecha=info.get("created_at"))
    except Exception as e:  # sin internet o Cloudinary caído: seguimos con el Excel local
        print("No se pudo traer el Excel de Cloudinary:", e)


@asynccontextmanager
async def inicio(app):
    excel_desde_nube()
    yield

app = FastAPI(title="Pole Dance Rojas Sport", version="11.0.0", lifespan=inicio)


# =====================================================================  Clave de acceso
def _token():
    return hmac.new(CLAVE.encode(), b"polesport-sesion", hashlib.sha256).hexdigest()


PUBLICAS = ("/login", "/icono/", "/manifest.webmanifest", "/sw.js", "/favicon.ico", "/logo.png")


@app.middleware("http")
async def proteger(request: Request, call_next):
    p = request.url.path
    if not CLAVE or p.startswith(PUBLICAS) or hmac.compare_digest(request.cookies.get("sesion", ""), _token()):
        return await call_next(request)
    if p == "/":
        return RedirectResponse("/login")
    return JSONResponse({"detail": "Sesión vencida: vuelve a entrar con la clave."}, status_code=401)


@app.get("/login", response_class=HTMLResponse)
def login_form(error: int = 0):
    return LOGIN.replace("{{ERROR}}", '<p class="er">Clave incorrecta</p>' if error else "")


@app.post("/login")
def login(clave: str = Form(...)):
    if not CLAVE or not hmac.compare_digest(clave.strip(), CLAVE):
        time.sleep(1)
        return RedirectResponse("/login?error=1", status_code=303)
    r = RedirectResponse("/", status_code=303)
    r.set_cookie("sesion", _token(), max_age=60 * 60 * 24 * 180, httponly=True, samesite="lax",
                 secure=os.environ.get("RENDER") is not None)
    return r


@app.get("/salir")
def salir():
    r = RedirectResponse("/login" if CLAVE else "/")
    r.delete_cookie("sesion")
    return r


# =====================================================================  Excel
def limpio(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return str(v)[:10]
    if isinstance(v, dt.time):  # una fecha vacía o en 0 que Excel guarda como hora
        return None
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
        raise HTTPException(404, "No se encuentra el Excel. Súbelo con el botón ⬆️ Excel.")
    return JSONResponse({"hojas": leer_excel()})


@app.get("/api/estado")
def estado():
    return {"memoria": "nube" if USA_CLOUD else "temporal", "clave": bool(CLAVE), **_estado,
            "excel_modificado": dt.datetime.fromtimestamp(os.path.getmtime(EXCEL_PATH)).strftime("%Y-%m-%d %H:%M")
            if os.path.exists(EXCEL_PATH) else None}


@app.post("/subir-excel")
async def subir_excel(archivo: UploadFile = File(...)):
    if not archivo.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Sube un archivo .xlsx")
    data = await archivo.read()
    try:  # validar antes de reemplazar: así un archivo dañado nunca borra los datos buenos
        hojas = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None)
        if not any("ingresos" in n.lower() for n in hojas):
            raise ValueError("no tiene la hoja de Ingresos")
    except Exception as e:
        raise HTTPException(400, f"Ese Excel no se puede leer ({e}).")
    with open(EXCEL_PATH, "wb") as f:
        f.write(data)
    _cache["mtime"] = None
    excel_a_nube(data)
    _estado.update(excel_origen="nube" if USA_CLOUD else "local", excel_fecha=dt.datetime.now().isoformat())
    return {"ok": True, "guardado_en_nube": USA_CLOUD}


@app.get("/descargar-excel")
def descargar_excel():
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(404, "El Excel no existe")
    return FileResponse(EXCEL_PATH, filename="Control_Inventario_Pole_Dance.xlsx")


# =====================================================================  Fotos (productos ENT-xxx y alumnas A-xxx)
def _pid(pid):
    pid = re.sub(r"[^A-Za-z0-9_-]", "", pid)
    if not pid:
        raise HTTPException(400, "ID inválido")
    return pid


def _jpeg(data: bytes) -> bytes:
    """Reduce la foto a máx. 1200px (JPEG) para ahorrar espacio."""
    from PIL import Image, ImageOps
    try:
        img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    except Exception:
        raise HTTPException(400, "Imagen no válida")
    img.thumbnail((1200, 1200))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return buf.getvalue()


@app.get("/api/fotos")
def lista_fotos():
    if USA_CLOUD:
        import requests
        ids, cursor = {}, None
        while True:
            p = {"prefix": "polesport/", "max_results": 500}
            if cursor:
                p["next_cursor"] = cursor
            r = requests.get(f"https://api.cloudinary.com/v1_1/{CLD[0]}/resources/image/upload",
                             params=p, auth=(CLD[1], CLD[2]), timeout=30)
            if not r.ok:
                return ids
            j = r.json()
            ids.update({x["public_id"].split("/", 1)[1]: x["version"] for x in j.get("resources", [])
                        if x["public_id"].count("/") == 1})
            cursor = j.get("next_cursor")
            if not cursor:
                return ids
    return {os.path.splitext(f)[0]: int(os.path.getmtime(os.path.join(FOTOS_DIR, f)))
            for f in os.listdir(FOTOS_DIR) if f.endswith(".jpg")}


@app.get("/foto/{pid}")
def ver_foto(pid: str, w: int = 0, v: int = 0):
    pid = _pid(pid)
    if USA_CLOUD:
        t = f"c_fill,g_auto,w_{w},h_{w},q_auto,f_auto/" if w else "q_auto,f_auto/"
        return RedirectResponse(f"https://res.cloudinary.com/{CLD[0]}/image/upload/{t}v{v or 1}/polesport/{pid}")
    ruta = os.path.join(FOTOS_DIR, pid + ".jpg")
    if not os.path.exists(ruta):
        raise HTTPException(404, "Sin foto")
    return FileResponse(ruta, media_type="image/jpeg")


@app.post("/fotos/{pid}")
async def subir_foto(pid: str, archivo: UploadFile = File(...)):
    pid = _pid(pid)
    data = await archivo.read()
    if len(data) > 25 * 1024 * 1024:
        raise HTTPException(413, "Foto demasiado grande")
    data = _jpeg(data)
    if USA_CLOUD:
        j = _cld_post("image", "upload", {"public_id": f"polesport/{pid}", "overwrite": "true", "invalidate": "true"},
                      files={"file": (pid + ".jpg", data)})
        return {"ok": True, "v": j.get("version", 1)}
    with open(os.path.join(FOTOS_DIR, pid + ".jpg"), "wb") as f:
        f.write(data)
    return {"ok": True, "v": int(time.time())}


@app.delete("/fotos/{pid}")
def borrar_foto(pid: str):
    pid = _pid(pid)
    if USA_CLOUD:
        _cld_post("image", "destroy", {"public_id": f"polesport/{pid}", "invalidate": "true"})
    else:
        ruta = os.path.join(FOTOS_DIR, pid + ".jpg")
        if os.path.exists(ruta):
            os.remove(ruta)
    return {"ok": True}


# =====================================================================  Logo e ícono de la app (PWA)
def _archivo(*patrones):
    for d in (STATIC, BASE):
        for pat in patrones:
            for f in sorted(glob.glob(os.path.join(d, pat))):
                if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
                    return f
    return None


def _icono_png(n: int, margen: float) -> bytes:
    from PIL import Image
    src = _archivo("icono*", "[Ll]ogo*")
    if not src:
        raise HTTPException(404, "Pon el logo en static/icono.png")
    img = Image.open(src).convert("RGBA")
    bb = img.getchannel("A").getbbox()
    if bb:
        img = img.crop(bb)  # quita el borde transparente para que el logo se vea grande
    lado = int(n * margen)
    img.thumbnail((lado, lado), Image.LANCZOS)
    fondo = Image.new("RGBA", (n, n), (255, 255, 255, 255))
    fondo.paste(img, ((n - img.width) // 2, (n - img.height) // 2), img)
    buf = io.BytesIO()
    fondo.convert("RGB").save(buf, "PNG", optimize=True)
    return buf.getvalue()


_iconos = {}


@app.get("/icono/{nombre}.png")
def icono(nombre: str):
    # 192 / 512 = normal (logo grande) · m192 / m512 = "maskable" (Android lo recorta en círculo: más margen)
    m = re.fullmatch(r"(m?)(\d{2,4})", nombre)
    if not m or not 16 <= int(m.group(2)) <= 1024:
        raise HTTPException(404)
    if nombre not in _iconos:
        _iconos[nombre] = _icono_png(int(m.group(2)), 0.62 if m.group(1) else 0.86)
    return Response(_iconos[nombre], media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})


@app.get("/favicon.ico")
def favicon():
    return icono("48")


@app.get("/logo.png")
def logo():
    f = _archivo("logo*", "[Ll]ogo*")
    if not f:
        raise HTTPException(404)
    return FileResponse(f, headers={"Cache-Control": "public, max-age=86400"})


@app.get("/manifest.webmanifest")
def manifest():
    return JSONResponse({
        "id": "/", "name": "Pole Dance Rojas Sport", "short_name": "Rojas Sport",
        "description": "Inventario, alumnas y finanzas de la academia",
        "start_url": "/", "scope": "/", "display": "standalone", "orientation": "any",
        "background_color": "#070b24", "theme_color": "#070b24", "lang": "es-CO",
        "icons": [
            {"src": "/icono/192.png", "sizes": "192x192", "type": "image/png", "purpose": "any"},
            {"src": "/icono/512.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
            {"src": "/icono/m192.png", "sizes": "192x192", "type": "image/png", "purpose": "maskable"},
            {"src": "/icono/m512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"},
        ]}, media_type="application/manifest+json")


@app.get("/sw.js")
def sw():
    # Guarda la última versión de la app y de los datos: si no hay internet, abre con lo último que vio.
    js = """const C='rojas-v11';
self.addEventListener('install',e=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x))))));
self.addEventListener('fetch',e=>{const u=new URL(e.request.url);
 if(e.request.method!=='GET'||u.origin!==location.origin||u.pathname.startsWith('/foto'))return;
 e.respondWith(fetch(e.request).then(r=>{if(r.ok&&!r.redirected){const c=r.clone();caches.open(C).then(x=>x.put(e.request,c))}return r})
 .catch(()=>caches.match(e.request)))});"""
    return Response(js, media_type="application/javascript", headers={"Cache-Control": "no-cache"})


@app.get("/", response_class=HTMLResponse)
def interfaz():
    return HTML


HEAD = '''<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Pole Dance Rojas Sport</title>
<link rel="manifest" href="/manifest.webmanifest"><meta name="theme-color" content="#070b24">
<link rel="icon" type="image/png" sizes="192x192" href="/icono/192.png"><link rel="icon" href="/favicon.ico">
<link rel="apple-touch-icon" sizes="180x180" href="/icono/180.png">
<meta name="apple-mobile-web-app-capable" content="yes"><meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Rojas Sport">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">'''

LOGIN = '''<!DOCTYPE html><html lang="es"><head>''' + HEAD + '''<style>
*{box-sizing:border-box;margin:0;font-family:system-ui,-apple-system,"Segoe UI",sans-serif}
body{min-height:100vh;display:grid;place-items:center;background:linear-gradient(135deg,#070b24,#0f1a4a 55%,#123a3a);padding:20px}
form{background:#fff;border-radius:24px;padding:32px 28px;width:100%;max-width:340px;text-align:center;box-shadow:0 20px 50px -10px #000a}
img{width:120px;margin-bottom:18px}h1{font-size:1.1rem;color:#1e1b4b;margin-bottom:18px}
input{width:100%;padding:13px 16px;border:1.5px solid #cbd5e1;border-radius:14px;font-size:1rem;margin-bottom:12px;outline:none}
input:focus{border-color:#22a94b}button{width:100%;padding:13px;border:0;border-radius:14px;font-weight:700;font-size:1rem;color:#fff;
background:linear-gradient(135deg,#22a94b,#1e3a8a);cursor:pointer}.er{color:#dc2626;font-size:.85rem;margin-bottom:10px}
</style></head><body><form method="post" action="/login"><img src="/logo.png" alt="Rojas Sport"><h1>Panel de la academia</h1>
{{ERROR}}<input type="password" name="clave" placeholder="Clave" autofocus autocomplete="current-password" required>
<button>Entrar</button></form></body></html>'''

HTML = r'''<!DOCTYPE html><html lang="es"><head>''' + HEAD + r'''<style>
:root{--p:#22a94b;--a:#1e3a8a;--ok:#10b981;--bad:#ef4444;--warn:#f59e0b;--tx:#1e1b4b;--mu:#64748b}
*{box-sizing:border-box;margin:0;padding:0;font-family:system-ui,-apple-system,"Segoe UI",sans-serif}
body{background:linear-gradient(135deg,#070b24,#0f1a4a 55%,#123a3a) fixed;min-height:100vh;color:var(--tx);
padding:env(safe-area-inset-top) 12px 40px}
.w{max-width:1100px;margin:0 auto}
header{display:flex;align-items:center;gap:14px;padding:16px 4px;color:#fff;flex-wrap:wrap}
header img{width:56px;height:56px;border-radius:14px;object-fit:contain;background:#fff;padding:4px;box-shadow:0 0 20px #22c55e55}
header h1{font-size:1.25rem;font-weight:800;background:linear-gradient(135deg,#fff,#86efac,#22c55e);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
header small{color:#cbd5e1;display:block}.sp{flex:1}.acc{display:flex;gap:8px;flex-wrap:wrap}
.bt{background:#ffffff22;color:#fff;border:1px solid #ffffff44;padding:8px 14px;border-radius:30px;font-weight:600;font-size:.8rem;cursor:pointer;text-decoration:none;display:inline-block}
.mem{font-size:.72rem;padding:4px 10px;border-radius:20px;font-weight:700;margin-top:4px;display:inline-block}
.mem.ok{background:#14532d;color:#bbf7d0}.mem.no{background:#7c2d12;color:#fed7aa;cursor:help}
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
th{position:sticky;top:0;background:#f8fafc;font-size:.7rem;text-transform:uppercase;color:#475569;padding:10px 12px;text-align:left;cursor:pointer;white-space:nowrap;border-bottom:1px solid #e2e8f0;z-index:1}
td{padding:8px 12px;border-bottom:1px solid #f1f5f9;font-size:.85rem;white-space:nowrap}
tr:hover td{background:#f0fdf4}
.chip{padding:3px 10px;border-radius:20px;font-size:.72rem;font-weight:700;background:#dcfce7;color:#166534}
.chip.r{background:#fee2e2;color:#991b1b}
.neg{color:var(--bad);font-weight:700}.pos{color:var(--ok);font-weight:700}
.mu{color:var(--mu);font-size:.8rem;margin:8px 2px}.more{width:100%;margin-top:10px;padding:10px;border:1.5px dashed #cbd5e1;border-radius:10px;background:#f8fafc;cursor:pointer;font-weight:700;color:var(--mu)}
.av{width:44px;height:44px;border-radius:10px;object-fit:cover;cursor:zoom-in;background:#f1f5f9}
.prog{height:16px;border-radius:9px;background:#e2e8f0;overflow:hidden}.prog i{display:block;height:100%;border-radius:9px;background:linear-gradient(90deg,var(--p),var(--a))}
.kpi small{display:block;margin-top:2px}
.cont .ln{display:flex;justify-content:space-between;gap:14px;align-items:flex-start;padding:11px 2px;border-bottom:1px solid #f1f5f9}
.cont .ln span{font-size:.9rem;font-weight:600}.cont .ln small,.cont .nt{display:block;color:var(--mu);font-size:.75rem;margin-top:3px;line-height:1.35}
.cont .ln b{font-size:1rem;white-space:nowrap}.cont .tot{border-bottom:0;border-top:2px solid var(--tx);margin-top:4px;background:#f8fafc;border-radius:8px;padding:12px 8px}
.cont .tot span{font-weight:800}.cont .tot b{font-size:1.1rem}.cont .nt{padding:0 8px 6px}
#toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#0f172a;color:#fff;padding:12px 18px;border-radius:14px;font-size:.9rem;z-index:100;display:none;box-shadow:0 10px 30px #0008}
</style></head><body><div class="w">
<header><img src="/icono/192.png" alt="">
<div><h1>Pole Dance Rojas Sport</h1><small id="sub">Cargando…</small><span id="mem"></span></div><div class="sp"></div>
<div class="acc"><label class="bt">⬆️ Excel<input type="file" accept=".xlsx" hidden onchange="subir(this.files[0]);this.value=''"></label>
<a class="bt" href="/descargar-excel">⬇️ Excel</a><button class="bt" onclick="cargar()">🔄</button><a class="bt" id="salir" href="/salir" hidden>🚪</a></div></header>
<nav id="nav"></nav><main id="main"><div class="card">Cargando…</div></main></div><div id="toast"></div>
<script>
if('serviceWorker'in navigator)navigator.serviceWorker.register('/sw.js');
const T={resumen:{n:'📊 Resumen'},alumnas:{n:'🎓 Alumnas',kw:'alumnas',req:'Nombre',fc:'Estado',foto:'ID Alumna'},
ingresos:{n:'🤸 Ingresos',kw:'ingresos',req:'Alumna',fc:'Tipo'},gastos:{n:'💸 Gastos',kw:'gastos',req:'Valor ($)',fc:'Categoría'},
stock:{n:'👗 Stock'},merc:{n:'📦 Entradas',kw:'marcancia',req:'Código Prenda',fc:'Clasificación',foto:'ID Entrada'},
activos:{n:'🏢 Activos',kw:'activos',req:'Nombre del Activo',fc:'Estado'},ventas:{n:'🛒 Ventas',kw:'ventas',req:'Código Prenda',fc:'Método Pago'},
cont:{n:'📑 Contable',kw:'contable'}};
// Columnas de apoyo del Excel que no se muestran en la app
const OCULTAS=/^(Clave \(auto\)|Alumna \(lista, auto\)|Nombre \+ Apellidos \(auto\)|Orden lista \(auto\)|ID Alumna \(auto\))$/;
let D={},tab='resumen',S={},F=new Map();
const $=s=>document.querySelector(s),esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmt=n=>(n<0?'-$':'$')+Math.abs(Math.round(n||0)).toLocaleString('es-CO'),num=v=>Number(v)||0;
const titulo=c=>c.replace(/ \(auto\)$/,'');
function toast(t,ms=3500){const e=$('#toast');e.textContent=t;e.style.display='block';clearTimeout(e._t);e._t=setTimeout(()=>e.style.display='none',ms)}
const hoja=kw=>{const k=Object.keys(D).find(n=>n.toLowerCase().includes(kw));return k?D[k]:null};
function tabla(id){const t=T[id];if(id==='stock')return stock();const h=hoja(t.kw);if(!h)return{cols:[],rows:[]};
// solo filas con ID en la 1.ª columna: así las filas de TOTAL dentro de las hojas no se cuentan dos veces
const rows=h.filas.filter(f=>f[0]!=null&&f[0]!=='').map(f=>Object.fromEntries(h.cols.map((c,i)=>[c,f[i]]))).filter(r=>r[t.req]!=null&&r[t.req]!=='');
return{cols:h.cols.filter(c=>!OCULTAS.test(c)&&c!=='nan'),rows}}
function stock(){const m={};tabla('merc').rows.forEach(r=>{const k=String(r['Código Prenda']).trim().toUpperCase();
const x=m[k]??={'Código':k,'Descripción':String(r['Descripción']||'').trim(),Entradas:0,Vendidas:0,Stock:0,'Costo total ($)':0};x.Entradas+=num(r['Cantidad (+)']);x['Costo total ($)']+=num(r['Costo Total ($)'])});
tabla('ventas').rows.forEach(r=>{const k=String(r['Código Prenda']).trim().toUpperCase();if(m[k])m[k].Vendidas+=num(r['Cant. Vendida (-)'])});
const rows=Object.values(m).map(x=>({...x,Stock:x.Entradas-x.Vendidas}));return{cols:['Código','Descripción','Entradas','Vendidas','Stock','Costo total ($)'],rows}}
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
const sinAsignar=ing.filter(r=>!String(r['ID Alumna (auto)']||'').startsWith('A-')).length;
const K=(t,v,id,c='')=>`<div class="card kpi" onclick="ir('${id}')"><h4>${t}</h4><b class="${c}">${v}</b></div>`;
const nom=a=>[a.Nombre,a.Apellidos].filter(Boolean).join(' ');
return `<div class="g k">${K('Ingresos',fmt(I),'ingresos','pos')}${K('Gastos',fmt(G),'gastos','neg')}${K('Resultado',fmt(I-G),'cont',I-G>=0?'pos':'neg')}
${K('Alumnas',al.length,'alumnas')}${K('Nuevas / Recurrentes',nuevas+' / '+rec,'alumnas')}${K('Prendas en stock',sum(st,'Stock'),'stock')}${K('Valor activos',fmt(sum(ac,'Valor Total ($)')),'activos')}</div>
${sinAsignar?`<div class="card" style="margin-bottom:14px;border-left:5px solid var(--warn)">⚠️ ${sinAsignar} ingreso(s) sin alumna reconocida. Revísalos en el Excel (celda roja en la columna Alumna).</div>`:''}
<div class="g c2"><div class="card" style="grid-column:1/-1"><h3>📅 Ingresos vs Gastos por mes <small class="mu">(🟢 ingresos · 🔴 gastos)</small></h3><div class="mes">
${ms.map(m=>`<div><div class="bb"><i title="${fmt(meses[m].i)}" style="height:${meses[m].i/mx*100}%;background:var(--p)"></i><i title="${fmt(meses[m].g)}" style="height:${meses[m].g/mx*100}%;background:var(--bad)"></i></div>${m}</div>`).join('')}</div></div>
<div class="card"><h3>💸 Gastos por categoría</h3>${barras(group(ga,'Categoría','Valor ($)'),'gastos','Categoría',1)}</div>
<div class="card"><h3>🤸 Ingresos por tipo</h3>${barras(group(ing,'Tipo','Total ($)'),'ingresos','Tipo',1)}</div>
<div class="card"><h3>💳 Ingresos por método de pago</h3>${barras(group(ing,'Método de Pago','Total ($)'),'ingresos','Método de Pago',1)}</div>
<div class="card"><h3>🎓 Alumnas por estado</h3>${barras(group(al.map(a=>({...a,Estado:a.Estado||'Sin pagos'})),'Estado'),'alumnas','Estado')}</div>
<div class="card"><h3>🏆 Alumnas que más han pagado</h3>${barras(al.map(a=>[nom(a),num(a['Total pagado ($)'])]).sort((a,b)=>b[1]-a[1]),'alumnas','',1)}</div>
<div class="card"><h3>🏢 Activos por estado</h3>${barras(group(ac,'Estado','Valor Total ($)'),'activos','Estado',1)}</div></div>
<p class="mu" style="color:#cbd5e1">Toca cualquier tarjeta o barra para ver el detalle filtrado.</p>`}
function vista(id){const t=T[id],s=S[id]??={q:'',f:'',sc:null,asc:1,n:100};
if(id==='cont')return contable();
let {cols,rows}=tabla(id);if(t.foto)cols=['Foto',...cols];const fc=t.fc;
const opts=fc&&cols.includes(fc)?[...new Set(rows.map(r=>r[fc]).filter(v=>v!=null))].sort():[];
let r=rows.filter(x=>(!s.f||String(x[fc])===s.f||(!fc||!cols.includes(fc))&&Object.values(x).includes(s.f))&&(!s.q||Object.values(x).join(' ').toLowerCase().includes(s.q.toLowerCase())));
if(s.sc)r.sort((a,b)=>((a[s.sc]??'')>(b[s.sc]??'')?1:-1)*s.asc);
const cell=(c,v,x)=>c==='Foto'?foto(x[t.foto]):v==null?'':typeof v==='number'&&/\(\$\)|precio|valor/i.test(c)?fmt(v)
 :/^(Varias|No encontrada)/.test(String(v))?`<span class="chip r">${esc(v)}</span>`:c==='Estado'||c===fc?`<span class="chip">${esc(v)}</span>`:esc(v);
return `<div class="card"><div class="tb"><input placeholder="🔍 Buscar en ${t.n}…" value="${esc(s.q)}" oninput="S['${id}'].q=this.value;S['${id}'].n=100;repintar('${id}')" id="q">
${opts.length?`<select onchange="S['${id}'].f=this.value;pintar()"><option value="">Todos (${fc})</option>${opts.map(o=>`<option ${o===s.f?'selected':''}>${esc(o)}</option>`).join('')}</select>`:''}</div>
<div class="mu">${r.length} de ${rows.length} registros${id==='ingresos'||id==='gastos'?' · Total filtrado: <b>'+fmt(sum(r,id==='gastos'?'Valor ($)':'Total ($)'))+'</b>':''}</div>
<div class="tw"><table><thead><tr>${cols.map(c=>`<th onclick="orden('${id}',this.dataset.c)" data-c="${esc(c)}">${esc(titulo(c))}${s.sc===c?(s.asc>0?' ▲':' ▼'):''}</th>`).join('')}</tr></thead>
<tbody>${r.slice(0,s.n).map(x=>`<tr>${cols.map(c=>`<td>${cell(c,x[c],x)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
${r.length>s.n?`<button class="more" onclick="S['${id}'].n+=200;pintar()">Ver más</button>`:''}</div>`}
// ---- Contable: tarjetas por sección, cifra a la derecha y explicación en gris debajo ----
function contable(){const h=hoja('contable');if(!h)return '<div class="card">No hay hoja contable.</div>';
const secs=[];let intro='',cur=null;const val={};
h.filas.forEach(f=>{const t=String(f[0]??'').trim(),n=f.find(x=>typeof x==='number'),nota=String(f.slice(1).find(x=>typeof x==='string')||'').replace(/^←\s*/,'');
 if(!t)return;
 if(/^\d+\.\s/.test(t)){const m=t.match(/^(\d+)\.\s*([^(]*)(\((.*)\))?/);cur={t:m[2].trim(),sub:m[4]||'',rows:[]};secs.push(cur);return}
 if(n===undefined){if(!cur&&!/RESUMEN/.test(t))intro=t;return}
 val[t]=n;(cur??=(secs.push({t:'',sub:'',rows:[]}),secs[secs.length-1])).rows.push({t,n,nota,tot:/^[A-ZÁÉÍÓÚÑ\/ ]{6,}/.test(t)})});
const busca=re=>{const k=Object.keys(val).find(k=>re.test(k));return k?val[k]:0};
const util=busca(/^UTILIDAD/),inv=busca(/^TOTAL INVERTIDO/),bal=busca(/^BALANCE NETO/);
const pct=inv>0?Math.max(0,Math.min(100,util/inv*100)):0,cls=v=>v<0?'neg':'pos';
const K=(t,v,sub,c)=>`<div class="card kpi" style="cursor:default"><h4>${t}</h4><b class="${c}">${fmt(v)}</b><small class="mu">${sub}</small></div>`;
const icon=['💼','🏗️','⚖️'];
return `<div class="g k">${K('Utilidad operativa',util,'el día a día',cls(util))}${K('Invertido para abrir',inv,'por recuperar','')}${K('Balance neto',bal,'caja real hasta hoy',cls(bal))}</div>
${inv>0?`<div class="card" style="margin-bottom:14px"><h3>🎯 Recuperación de la inversión</h3>
<div class="prog"><i style="width:${pct}%"></i></div>
<p class="mu">Con la utilidad operativa actual llevas <b>${pct.toFixed(1)}%</b> recuperado de ${fmt(inv)}.${util>0&&pct<100?` Te faltan <b>${fmt(inv-util)}</b>.`:''}</p></div>`:''}
<div class="g c2">${secs.map((s,i)=>`<div class="card cont"><h3>${icon[i]||'📑'} ${esc(s.t[0]+s.t.slice(1).toLowerCase())}</h3>${s.sub?`<p class="mu" style="margin-top:-8px">${esc(s.sub)}</p>`:''}
${s.rows.map(r=>`<div class="ln${r.tot?' tot':''}"><div><span>${esc(r.tot?r.t.replace(/^[^(]+/,x=>x[0]+x.slice(1).toLowerCase()):r.t)}</span>${r.nota&&!r.tot?`<small>${esc(r.nota)}</small>`:''}</div>
<b class="${r.n<0?'neg':r.tot?'':''}">${fmt(r.n)}</b></div>${r.nota&&r.tot?`<small class="nt">${esc(r.nota)}</small>`:''}`).join('')}</div>`).join('')}</div>
${intro?`<p class="mu" style="color:#cbd5e1">ℹ️ ${esc(intro)}</p>`:''}`}
const foto=id=>!id?'':`<span style="display:flex;gap:6px;align-items:center">${F.has(id)?`<img class="av" loading="lazy" src="/foto/${id}?w=120&v=${F.get(id)}" onclick="zoom('${id}')">`:''}<label class="chip" style="cursor:pointer">📷${F.has(id)?'':' Subir'}<input type="file" accept="image/*" hidden onchange="subirFoto('${id}',this.files[0])"></label>${F.has(id)?`<span class="chip r" style="cursor:pointer" onclick="borrarFoto('${id}')">✕</span>`:''}</span>`;
function zoom(id){const d=document.createElement('div');d.style.cssText='position:fixed;inset:0;background:#000d;display:flex;align-items:center;justify-content:center;z-index:99;cursor:zoom-out;padding:20px';
d.innerHTML=`<img src="/foto/${id}?v=${F.get(id)}" style="max-width:100%;max-height:100%;border-radius:12px">`;d.onclick=()=>d.remove();document.body.appendChild(d)}
async function subirFoto(id,f){if(!f)return;toast('Subiendo foto…',60000);const fd=new FormData();fd.append('archivo',f);
const r=await fetch('/fotos/'+id,{method:'POST',body:fd});if(r.ok){F.set(id,(await r.json()).v);pintar();toast('✅ Foto guardada')}else toast('❌ No se pudo subir la foto')}
async function borrarFoto(id){if(!confirm('¿Borrar esta foto?'))return;const r=await fetch('/fotos/'+id,{method:'DELETE'});if(r.ok){F.delete(id);pintar()}}
function orden(id,c){const s=S[id];s.asc=s.sc===c?-s.asc:1;s.sc=c;pintar()}
function repintar(id){const p=$('#q').selectionStart;pintar();const q=$('#q');q.focus();q.setSelectionRange(p,p)}
function pintar(){$('#nav').innerHTML=Object.entries(T).map(([k,t])=>`<button class="${k===tab?'on':''}" onclick="tab='${k}';pintar()">${t.n}</button>`).join('');
$('#main').innerHTML=tab==='resumen'?resumen():vista(tab)}
async function api(u){const r=await fetch(u);if(r.status===401){location.href='/login';throw new Error('Sesión vencida')}
if(!r.ok)throw new Error((await r.json()).detail);return r.json()}
async function cargar(){try{D=(await api('/api/datos')).hojas;
try{F=new Map(Object.entries(await api('/api/fotos')))}catch(e){}
try{const e=await api('/api/estado');$('#salir').hidden=!e.clave;
$('#mem').innerHTML=e.memoria==='nube'?'<span class="mem ok">☁️ Memoria permanente activa</span>'
:'<span class="mem no" title="Sin Cloudinary: las fotos y el Excel que subas se borran cuando Render reinicia la app.">⚠️ Memoria temporal</span>';
$('#sub').textContent='Inventario, alumnas y finanzas · Excel del '+(e.excel_modificado||'—')}catch(e){}
pintar()}catch(e){$('#main').innerHTML='<div class="card">⚠️ '+esc(e.message)+'</div>'}}
async function subir(f){if(!f)return;toast('Subiendo Excel…',60000);const fd=new FormData();fd.append('archivo',f);
const r=await fetch('/subir-excel',{method:'POST',body:fd});const j=await r.json().catch(()=>({}));
if(r.ok){toast(j.guardado_en_nube?'✅ Excel guardado (con respaldo en la nube)':'✅ Excel cargado (memoria temporal)');cargar()}else toast('❌ '+(j.detail||'No se pudo subir el Excel'),6000)}
cargar();
</script></body></html>'''

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))