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
    try:
        from PIL import Image, ImageOps
    except ImportError:
        return data
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
    # 1) íconos ya hechos en static/iconos (no necesitan nada en el servidor)
    m = re.fullmatch(r"(m?)(\d{2,4})", nombre)
    if not m or not 16 <= int(m.group(2)) <= 1024:
        raise HTTPException(404)
    if nombre in ICONOS:  # íconos incluidos dentro de este archivo
        return Response(_icono_incluido(nombre), media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})
    hecho = os.path.join(STATIC, "iconos", nombre + ".png")
    if os.path.exists(hecho):
        return FileResponse(hecho, media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})
    # 2) si no existe ese tamaño, se intenta generar; y si falla, se entrega el logo tal cual
    try:
        if nombre not in _iconos:
            _iconos[nombre] = _icono_png(int(m.group(2)), 0.62 if m.group(1) else 0.86)
        return Response(_iconos[nombre], media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})
    except Exception as e:
        print("No se pudo generar el ícono:", repr(e))
        f = _archivo("icono*", "logo*", "[Ll]ogo*")
        if not f:
            raise HTTPException(404, "Falta el logo")
        return FileResponse(f)


@app.get("/favicon.ico")
def favicon():
    return icono("192")


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
    js = """const C='rojas-v12';
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
<a class="bt" href="/descargar-excel">⬇️ Excel</a><button class="bt" onclick="cargar()">🔄</button><button class="bt" id="inst" hidden onclick="instalar()">📲 Instalar</button><a class="bt" id="salir" href="/salir" hidden>🚪</a></div></header>
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
// ---- Instalar como app en el celular ----
let promptInst=null;const enApp=matchMedia('(display-mode: standalone)').matches||navigator.standalone;
const esIOS=/iphone|ipad|ipod/i.test(navigator.userAgent);
addEventListener('beforeinstallprompt',e=>{e.preventDefault();promptInst=e;$('#inst').hidden=false});
addEventListener('appinstalled',()=>{$('#inst').hidden=true;toast('✅ App instalada: búscala en tu pantalla de inicio')});
if(!enApp&&esIOS)$('#inst').hidden=false;
async function instalar(){if(promptInst){promptInst.prompt();await promptInst.userChoice;promptInst=null;$('#inst').hidden=true;return}
const d=document.createElement('div');d.style.cssText='position:fixed;inset:0;background:#000b;display:flex;align-items:flex-end;justify-content:center;z-index:99;padding:16px';
d.innerHTML=`<div class="card" style="max-width:420px;width:100%"><h3>📲 Instalar en iPhone</h3>
<p style="margin:10px 0;line-height:1.6">1. Abre esta página en <b>Safari</b>.<br>2. Toca el botón <b>Compartir</b> (cuadro con flecha ⬆️).<br>3. Elige <b>"Agregar a pantalla de inicio"</b> y luego <b>Agregar</b>.</p>
<button class="more" onclick="this.closest('div[style]').remove()">Entendido</button></div>`;document.body.appendChild(d)}
cargar();
</script></body></html>'''


# =====================================================================  Íconos de la app (logo "S"), ya listos
import base64


def _icono_incluido(n):
    return base64.b64decode("".join(ICONOS[n].split()))


ICONOS = {
    "180": """
iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAABgFBMVEX////+///+//79//7+/v/+/v79/v79/v39/f79/f38/v38
/vz8/f78/f38/fz8/P37/vz7/P34/fn4+vvv+fLy9vft7/Tl9uji7+zS8NrB7Mni5e3P4+Gw5rqd4KvR1+Ot1c2zwtKF2ZRw1H9v
y5JdzHY/z1BMymA5zEtBxldAv2M4y003xUo3wkk3wkY2wkY3wkU2wEs2wUc3wUUyxlc0wU01wUgywFAvvVcpvmgyv1AvvlMtuWEt
ulsouWkuvFctuligrcWNnLplmqZrhqYqs3AptGAqq3Usk4YlsXQkqnIjpXYjoHkjmnsjk30ii38hhIBkeKFQZZQmeYMvZIcfeoAf
bn8fZX4eXX03UIUiUnwbT3gjRXohPngYQXQXPHIaOHIWOHEXNHATNHITMnAUMHASL28RMHIRLm8RLW8SLG8QLXAQK28QLG0PLHAP
K28PKm8PKm0PKW0QKWsPKWsNKW8OKWwOKWsNKG0OKGsNKGoMKGsNJ2oMJ2kLJ2gMJmgT3MliAAAR+UlEQVR42u2c+1caWRaFr2AB
A2GQt4BQ8hIUMT4wiSaagBix25kMZTrpSjAiCPgCUcBXtPOvzzm3Cih8zQ/ThbgWeyVppRW+2rXPuedeEgnpq6+++uqrr7766uup
xRCL59lBq4hjxkIUzwtaSZyvHETb+6ADEmN1xBN2w5+9LgUj/YQEwz6illjf+1FRENP4ZEhHmI7a7PlIO+bmonap1b1v9JDaHZ6b
dCp0zwKXUSp11F5POBr2EJ1CrVP2tt2KJp/JbveHp8JBU7OXKHuVWI39weRw+/xBUHR6bm7KD/K57fR6FL1oMvxhd1NIp8NusY8D
9LTTDtcQDAV9DixOde8hm9z+oOAqyoFGw/KCsjh9oaDbQoha1VvIdl/Qj1xEAYWnGlK4w9MA7WOG1LQuTU5/yAcXpO6VkACUxRcS
E6BoLuJjU3Nz4vKiUGN2HD4BuyeQYfHzhPyOZiUK3ivswfDcXNivFNd2ym33j/ssRKXuBZvdoZDzdpmpiT0ELW9I0Y6DCrAdwXH3
E5vNKHVDYF/YZyKmIbW6Y7YYIu7wpKOTDz32RPxPubgLS7QjFHK0xlFJmSkUpvFQ5+AnZMkeHHc+2ciHbkETngwH/UFozh4nbXbt
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
X3311Vdf9+u/6iWSGFXltsIAAAAASUVORK5CYII=
""",
    "192": """
iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAMAAABlApw1AAABgFBMVEX////+///+//79//7+/v/+/v79/v79/v39/f79/f38/v38
/vz8/f78/f38/P37/vz7/P34/fn4+vru+fHw9vXr7vPg9eTc7+jN79a+68Xf5OvI49ys5beZ36fQ1+PDy9qj0cKsu8191o5i0HRi
xohMx2g6z0w6yUs8wVk4wkY3x0s2wkk3wkY2wkY3wkU2wEw2wUg2wUYyxlczwU01wUgzwE8wvlkvvVUvvlQrvmMyv1AwvlIuu1ou
vFUtuV0ouGsquF2cq8SIlrZMqJJph6YosnIpsmInpHgnjYEkr3MjqHgjo3Yjnnkil3whkX0hhoBedp5KYZEmeYMlaIMeeYAebH8f
Yn4eW301ToQhUHscUXkaSHYfQncePHUXP3MWOnIYNnEVN3EWM3AUNHETMnATMG8TL28RMHIRLm8SLG4RLG8QLW8QK28QLG0PLHAP
K28PK2wQKmwPKm0PKW0PKWsNKW8OKWwOKWsOKGsNKGwNKGoNJ2oOKGkMJ2oMJ2gMJmivCh1TAAAURklEQVR42u2c6UMaydbGK0AD
VwZENkFZBAQU4sZkTExMDC0a9U5mwEwmbSIjooIbIMguJv/6Pae6gcYt74dLK+/lScwEjfD86ix1qpuRkIEGGmiggQYaaKCBBvqp
5MRuIco+BlARn4No+te/jIzMeYCirzzLGHEGWRZ88v7y35XwGmIPzlnEIXjWXzga4gnO2Lpy6Jmsv2rA+zzYx1UsI8aphaC3z6q4
7V4lG5ZbZxdmp0wquUzWZ941wqLbAtMLMxb6OY28TwKh0vCrbbRYbXZPcHY24LBajCM8xJNnkNFlNlodHi/KMzU9DTkED3xejx0j
oVIxT3nxsclb7ODcYeMX3RuYnoUq1hgtdo9vyuugDE+1HlSw+Ea71+exGVubsjEwPT09O23hHwOaz2tjniiCDFbf4vF57eheqVEp
ZTIVsUzPggI2RiOX083A4oB/MfIEEXj7fo+VRkLWmoRsfiCY9gtbmVIDc4YGwuAwkie2N2DyOAJeSBW5aG0hhbyB2dnnAXvbrgpJ
bL4pxwiRy56S/xG732cVqrhrmoYqDjiIyCsDhIx9ymd7MkHAXLf6/LCmwxrZrWHOGgj4hm9kPMYL0s34FAhUGlxzO80eYReTdXv1
+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
XAD5Z1pYAAAAAElFTkSuQmCC
""",
    "512": """
iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAMAAADDpiTIAAABgFBMVEX////+///+//79//7+/v/+/v79/v79/v39/f79/f37/fz2
+/jx9fbk8+rM79PZ5Oiq5LjB0dqrt8yG2ZZzzo9VympAxVM4xkg3wkc3wkU3wkQ2xUo2wkg2wkY2wkVVtn83wEc2wUk2wUc2wUU2
wEs0w0w0wUo1wUkywU41wUgzwE40wEoxv1Izv00wvFkuvFYxvlUuvVUwvlIsvFstvFcruGQqs2kltGstulkquV0rtl0nrXgnrWYj
rHUmpHcio3dujKknkIAliYAhmnkgkXsfh34ef38/c48icIA5YIkgYH0edX8eaX0dXnwdVnooTH0bS3cZR3UcP3QcOnIXPHEVOXEW
NW4UNW8VMm4TM28TMG8RMXARL24SLm8RLm4RLW8RLG0QLW4QLG4QK24PLG4PK20QKm0PKm4PKm0PKW0PKWwOKm0OKWwOKWsNKWsP
KGsOKGsOKGoNKGsNKGoNKGkNJ2oNJmgMKGoMJ2kLJ2sMJ2gMJmkMJmgMJmcMJWgLJmhcivTqAABLyklEQVR42u2diUMT1/bHQ2AC
ZCELYbGgQMW6o9a1WnVs32uavvalj7jEYEgghEAWhAACan/+679zzr2zJZNkAqFVc76vBayKzzmfe7Z77h2Xi8VisVgsFovFYrFY
LBaLxWKxWCwWi8VisVgsFovFYrFYLBaLxWKxWCwWi8VisVgsFovFYrFYLBaLxWKxWCwW68uTNzDID6F35XaFxkZcjEDv2t8/djXk
8vCT6FF5XCPXr0fdisKPolcBCF29OuYHT9DGUbC+1hAwevXqfLskQGEAvlIpLm/06nlOAr7mFd42Bzx/9fwo+/iv0vietmaFHHD8
/PnLY14m4Osz/2B/+9S+3xU6DxoPcCfgazO/d9DJmoYcEAGYHQEUWF/T6vc6yurcLt8YAgBJgINf38cP9ksp7gcdJvWDrsA42v9s
dNDVNl70eb2cKHwRy7/PsZ36XSOz5AHat4IorHi5WvwC7N/BMh10hc4RAOMjzjoBHg/3jL/swt8ixeWJThAA50NOs0A3I/A1weIf
O0P2P9tBK4gTga9GWg54/vy5MR8btvek54Dnz3ErqDc9QOicJIBbQT0oxeUelTng+YkQA9CDOaBvTAcg2neC7J57hF9sDnhGhoAJ
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
""",
    "m192": """
iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAMAAABlApw1AAABgFBMVEX////+///+//79//3+/v/+/v79/v79/v38/v38/vz9/f79
/f38/f78/f38/P37/vz7/P37/Pz4/fn4+vvv+vH09vjw8vbp+Ovn9PDY893p7PHa6ujI7s/C6tDX3+fHz92258Gh4quf17u1wNOK
25xy1YJ0yKBYy248z1A9yk9Gx104x0o+wV03wUo3wkY4wUYxx10vwl41xU40wE02wkY1wkc2wUc1wUk1wEorvGwuvFsvvFctvFcy
vlEuvVWrtcuXrcGOn7uElLQzuWcruF8wtW1GpJEntm4nsXAlrXEkqHUko3gknXskl30kkYB0h6thdZ8th4gyc4kgh38ffIEgc4Eg
bYBHXo8pYYMeYn4eWHwrUX8oR30bTXgaRXYgP3cdOnQWPHMWN3IXNXEUNXEVMnETMnIUL28SL28RLnASLXARLW4QLXAQLG8QK24P
K28QKm4PKm8PKm0PKW0QKmwPKWoOKm8OKW0NKW4OKWsOKGsNKGsNJ2oNKGkMJ2oMJmir36IaAAANZ0lEQVR42u2b+1cS6xrHH5EB
4QQiV7nIRS6DTKBi21JLLhm6y2LMigrbKAjIRUxB1Aj718/zzgx4yR/2OWedM8M682kt77W+3/e5vpMAyMjIyMjIyMjIyMjIyMjI
yMjIyMjIyMjIyMjI/BdQgNeOb0YXNdBeGB/Fg+ehQPE4AJqR009RwgcqsP1BD+2MjV4ujcMMEzGPahFQinHjuI954hrXKKjRE69R
cx8EGcbHVbPw+YjUMNd3jDaX1xuJMUGX3TZJCkI9GqlEEfVGly+A+HzzsdicLxAMBn12UhMjYIFkiisQDMzYyEe22HwsagYwu3zB
oNcofQuYJUZfMOAycj3IqPEy8zHGq+YmgT0Q9NlA0mNNgYfupQNcspCEHwcvE5sP+0CjUBAPZh8dMINasuWMZ+uig3YSB2owkX0M
Q5sVXN6o0IM5QPukGQRKjdltDODmAxqNRiFkOr6nGTsMjpyEwU7TNgkuF9yZojaXzWbmv8Bb0IBvTqG4KVxiwRfxgtQ6qpp0Hpph
5rBkIzS2TOz74xT3DVdEc2eVQGeuOR+oFNI6fzvNhJkn2PRjMYYJh2O00G+IgQmghBLnRU+CPRIEhUJK+rFU56Pzyyh/PjaPxJhw
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
""",
    "m512": """
iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAMAAADDpiTIAAABgFBMVEX////+///+//79//7+/v/+/v79/v79/v39/f79/f38/v35
/fr8/P30+vfs9vDo7PHP79bO4d+t5LmT3Z/Bz9qtuM141It+vK5QzGFFxF04yEo3xko3wkY3wkQ4wUlGtoI0xU80wkw2wkY0wkY2
wkU0wUw1wUk1wUg2wUY1wUcwvlUxv1ItvVguvVQ0wEw1wEg1v00yv08tuWIsulsot2gquF8vvFcuvFQtu1gtuVgps20ps2AksXAo
q3IlqnQjqnZ0lq0nln4kmHsionchm3ggknsfiX1leKE9cY8kfIEjbIAef38edX8ea34fY30yWIQeV3scWXsbUHknRnwbR3UYRXUY
P3IaO3IWOnEXNm8WM28TNXETM28TMm8SMW8UL3ASL24SLm8RL3ARLm4RLW8QLW8SLG8QK28QK20PLG4PK24PK2wPKm4PKm0PKmwO
Km0PKW0PKWwOKWwPKWoOKWsOKGoNKW0NKGsNKGoNJ2oMJ2kMJ2gMJmgMJmcMJWgLJmhwxkFrAAAznUlEQVR42u19iV9a1xPvFbzA
Y99dsGjSxFhrY2oTm2bFps0leW3przQlaiKugEYBRUE0efnX38ycey+X1Q3M4nwTt6x++H7PzHfmzDlXkhgMBoPBYDAYDAaDwWAw
GAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPB
YDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwGg8FgMBgMBoPBYDAYDAaDwWAwGAwG
g8FgMBhXEl63ZOJX4arCInnu+CUrvxBXFVbJPxZmAVxdmKXAtUFH6xxgkvn1+dohS6bwt3e8HAK+fpja/Kp7aHqaTcBXb/UsbTOA
d3r62wAL4KuG3W7q4AG/nf520Cpxuv96V7/d0qkKDIMAhtxtQ4C9j1/BLzz3WzsaA8fgjYmJ6fYu0G4382v41Vk/QwZwRyemJzqa
AJOF08NXC6vknQBEwh2VInOn+OsVQCACArg+6OLtgCuaIcKR6YmJb6NeiVP9leTfMYgRYOI6t4KuaAZwD91AAQwHOAVcUQ84MkEC
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
m4CJIwAAAABJRU5ErkJggg==
""",
}


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
