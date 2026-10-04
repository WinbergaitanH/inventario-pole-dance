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
            {"src": "/icono/192.png?v=3", "sizes": "192x192", "type": "image/png", "purpose": "any"},
            {"src": "/icono/512.png?v=3", "sizes": "512x512", "type": "image/png", "purpose": "any"},
            {"src": "/icono/m192.png?v=3", "sizes": "192x192", "type": "image/png", "purpose": "maskable"},
            {"src": "/icono/m512.png?v=3", "sizes": "512x512", "type": "image/png", "purpose": "maskable"},
        ]}, media_type="application/manifest+json")


@app.get("/sw.js")
def sw():
    # Guarda la última versión de la app y de los datos: si no hay internet, abre con lo último que vio.
    js = """const C='rojas-v16';
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
<link rel="icon" type="image/png" sizes="192x192" href="/icono/192.png?v=3"><link rel="icon" href="/favicon.ico">
<link rel="apple-touch-icon" sizes="180x180" href="/icono/180.png?v=3">
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
.chip.r{background:#fee2e2;color:#991b1b}.chip.a{background:#fef3c7;color:#92400e}
.neg{color:var(--bad);font-weight:700}.pos{color:var(--ok);font-weight:700}
.mu{color:var(--mu);font-size:.8rem;margin:8px 2px}.more{width:100%;margin-top:10px;padding:10px;border:1.5px dashed #cbd5e1;border-radius:10px;background:#f8fafc;cursor:pointer;font-weight:700;color:var(--mu)}
.av{width:44px;height:44px;border-radius:10px;object-fit:cover;cursor:zoom-in;background:#f1f5f9}
.prog{height:16px;border-radius:9px;background:#e2e8f0;overflow:hidden}.prog i{display:block;height:100%;border-radius:9px;background:linear-gradient(90deg,var(--p),var(--a))}
.kpi small{display:block;margin-top:2px}
.cont .ln{display:flex;justify-content:space-between;gap:14px;align-items:flex-start;padding:11px 2px;border-bottom:1px solid #f1f5f9}
.cont .ln span{font-size:.9rem;font-weight:600}.cont .ln small,.cont .nt{display:block;color:var(--mu);font-size:.75rem;margin-top:3px;line-height:1.35}
.cont .ln b{font-size:1rem;white-space:nowrap}.cont .tot{border-bottom:0;border-top:2px solid var(--tx);margin-top:4px;background:#f8fafc;border-radius:8px;padding:12px 8px}
.cont .tot span{font-weight:800}.cont .tot b{font-size:1.1rem}.cont .nt{padding:0 8px 6px}
.res{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:0 0 12px}.res .sp{flex:1}.chip.g{background:#e0e7ff;color:#1e3a8a}
.lnk{border:0;background:none;color:var(--a);font-weight:700;cursor:pointer;font-size:.8rem;padding:4px 6px}
.cards{display:grid;gap:10px;grid-template-columns:repeat(auto-fill,minmax(260px,1fr))}
.it{border:1px solid #e2e8f0;border-radius:14px;padding:12px;display:flex;gap:10px;background:#fff}.itb{min-width:0;flex:1}
.itb>div:first-child{margin-bottom:6px;font-size:.92rem}
.it dl{display:grid;gap:3px;margin:0}.it dl div{display:flex;justify-content:space-between;gap:10px;font-size:.8rem;border-bottom:1px dashed #f1f5f9;padding:2px 0}
.it dt{color:var(--mu)}.it dd{margin:0;text-align:right;overflow-wrap:anywhere}
table.bal td,table.bal th{text-align:right}table.bal td:first-child,table.bal th:first-child{text-align:left}
table.bal .e{background:#ecfdf3;font-weight:700}table.bal .s{background:#fef2f2;font-weight:700}table.bal .iv{background:#f5f3ff;color:#6d28d9;font-style:italic}table.bal .b{background:#eff6ff;font-weight:700}
.seg{border:1.5px solid #cbd5e1;background:#f8fafc;border-radius:30px;padding:8px 14px;font-weight:700;cursor:pointer;color:var(--tx)}.seg.on{background:linear-gradient(135deg,var(--p),var(--a));color:#fff;border-color:transparent}
table.bal .tt td{font-weight:800;border-top:2px solid var(--tx);background:#f8fafc}
#toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#0f172a;color:#fff;padding:12px 18px;border-radius:14px;font-size:.9rem;z-index:100;display:none;box-shadow:0 10px 30px #0008}
</style></head><body><div class="w">
<header><img src="/icono/192.png?v=3" alt="">
<div><h1>Pole Dance Rojas Sport</h1><small id="sub">Cargando…</small><span id="mem"></span></div><div class="sp"></div>
<div class="acc"><label class="bt">⬆️ Excel<input type="file" accept=".xlsx" hidden onchange="subir(this.files[0]);this.value=''"></label>
<a class="bt" href="/descargar-excel">⬇️ Excel</a><button class="bt" onclick="cargar()">🔄</button><button class="bt" id="inst" hidden onclick="instalar()">📲 Instalar</button><a class="bt" id="salir" href="/salir" hidden>🚪</a></div></header>
<nav id="nav"></nav><main id="main"><div class="card">Cargando…</div></main></div><div id="toast"></div>
<script>
if('serviceWorker'in navigator)navigator.serviceWorker.register('/sw.js');
const T={resumen:{n:'📊 Resumen'},bal:{n:'💰 Balance'},alumnas:{n:'🎓 Alumnas',kw:'alumnas',req:'Nombre',fc:'Estado',foto:'ID Alumna'},
ingresos:{n:'🤸 Ingresos',kw:'ingresos',req:'Alumna',fc:'Tipo'},gastos:{n:'💸 Gastos',kw:'gastos',req:'Valor ($)',fc:'Categoría'},
stock:{n:'👗 Catálogo',kw:'catálogo',req:'Código',fc:'Estado',foto:'Código',orden:['Código','Descripción','Talla','Color','Stock','Estado','Precio venta ($)']},merc:{n:'📦 Entradas',kw:'marcancia',req:'Código Prenda',fc:'Clasificación',foto:'ID Entrada'},
activos:{n:'🏢 Activos',kw:'activos',req:'Nombre del Activo',fc:'Estado'},ventas:{n:'🛒 Ventas',kw:'ventas',req:'Código Prenda',fc:'Método Pago'},
cont:{n:'📑 Contable',kw:'contable'}};
// Columnas de apoyo del Excel que no se muestran en la app
const OCULTAS=/^(Prefijo \(auto\)|Clave modelo \(auto\)|Modelo nuevo \(auto\)|N° modelo \(auto\)|Clave \(auto\)|Alumna \(lista, auto\)|Nombre \+ Apellidos \(auto\)|Orden lista \(auto\)|ID Alumna \(auto\))$/;
let D={},tab='resumen',S={},F=new Map();
const $=s=>document.querySelector(s),esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmt=n=>(n<0?'-$':'$')+Math.abs(Math.round(n||0)).toLocaleString('es-CO'),num=v=>Number(v)||0;
const titulo=c=>c.replace(/ \(auto\)$/,'');
function toast(t,ms=3500){const e=$('#toast');e.textContent=t;e.style.display='block';clearTimeout(e._t);e._t=setTimeout(()=>e.style.display='none',ms)}
const hoja=kw=>{const k=Object.keys(D).find(n=>n.toLowerCase().includes(kw));return k?D[k]:null};
function tabla(id){const t=T[id];const h=t.kw?hoja(t.kw):null;if(id==='stock'&&!h)return stock();if(!h)return{cols:[],rows:[]};
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
function ir(id,col,val){S[id]={q:'',f:val&&col?val:'',col:col||'',mes:'',sc:null,asc:1,n:100,vista:''};tab=id;pintar()}
function resumen(){const al=tabla('alumnas').rows,ing=tabla('ingresos').rows,ga=tabla('gastos').rows,ac=tabla('activos').rows,st=tabla('stock').rows,ve=tabla('ventas').rows;
const IC=sum(ing,'Total ($)'),IV=sum(ve,'Total Venta ($)'),I=IC+IV,G=sum(ga,'Valor ($)'),nuevas=al.filter(a=>a.Estado==='Nueva').length,rec=al.filter(a=>a.Estado==='Recurrente').length;
const meses={};ing.forEach(r=>{const m=String(r.Fecha||'').slice(0,7);if(m)(meses[m]??={i:0,g:0}).i+=num(r['Total ($)'])});
ga.forEach(r=>{const m=String(r.Fecha||'').slice(0,7);if(m)(meses[m]??={i:0,g:0}).g+=num(r['Valor ($)'])});
ve.forEach(r=>{const m=String(r.Fecha||'').slice(0,7);if(m)(meses[m]??={i:0,g:0}).i+=num(r['Total Venta ($)'])});
const agot=st.filter(x=>x.Estado==='Agotado').length,ult=st.filter(x=>x.Estado==='Última unidad').length;
const masV={};ve.forEach(r=>{const k=(r['Código Prenda']||'')+' · '+String(r['Descripción']||'').trim();masV[k]=(masV[k]||0)+num(r['Cant. Vendida (-)'])});
const ms=Object.keys(meses).sort(),mx=Math.max(...ms.map(m=>Math.max(meses[m].i,meses[m].g)),1);
const sinAsignar=ing.filter(r=>!String(r['ID Alumna (auto)']||'').startsWith('A-')).length;
const K=(t,v,id,c='',sub='')=>`<div class="card kpi" onclick="ir('${id}')"><h4>${t}</h4><b class="${c}">${v}</b>${sub?`<small class="mu" style="display:block;margin:2px 0 0">${sub}</small>`:''}</div>`;
const nom=a=>[a.Nombre,a.Apellidos].filter(Boolean).join(' ');
return `<div class="g k">${K('Ingresos',fmt(I),'ingresos','pos','clases '+fmt(IC)+' · ventas '+fmt(IV))}${K('Gastos',fmt(G),'gastos','neg')}${K('Resultado',fmt(I-G),'cont',I-G>=0?'pos':'neg')}
${K('Alumnas',al.length,'alumnas')}${K('Nuevas / Recurrentes',nuevas+' / '+rec,'alumnas')}${K('Unidades en stock',sum(st,'Stock'),'stock','',st.length+' productos')}${K('Valor activos',fmt(sum(ac,'Valor Total ($)')),'activos')}</div>
${agot||ult?`<div class="card" style="margin-bottom:14px;border-left:5px solid var(--bad);cursor:pointer" onclick="ir('stock','Estado','Agotado')">👗 Inventario: <b>${agot}</b> producto(s) agotado(s) y <b>${ult}</b> en su última unidad. Toca para verlos.</div>`:''}
${sinAsignar?`<div class="card" style="margin-bottom:14px;border-left:5px solid var(--warn)">⚠️ ${sinAsignar} ingreso(s) sin alumna reconocida. Revísalos en el Excel (celda roja en la columna Alumna).</div>`:''}
<div class="g c2"><div class="card" style="grid-column:1/-1"><h3>📅 Ingresos vs Gastos por mes <small class="mu">(🟢 ingresos · 🔴 gastos)</small></h3><div class="mes">
${ms.map(m=>`<div><div class="bb"><i title="${fmt(meses[m].i)}" style="height:${meses[m].i/mx*100}%;background:var(--p)"></i><i title="${fmt(meses[m].g)}" style="height:${meses[m].g/mx*100}%;background:var(--bad)"></i></div>${m}</div>`).join('')}</div></div>
<div class="card"><h3>💸 Gastos por categoría</h3>${barras(group(ga,'Categoría','Valor ($)'),'gastos','Categoría',1)}</div>
<div class="card"><h3>🤸 Ingresos por tipo</h3>${barras(group(ing,'Tipo','Total ($)'),'ingresos','Tipo',1)}</div>
<div class="card"><h3>💳 Ingresos por método de pago</h3>${barras(group(ing,'Método de Pago','Total ($)'),'ingresos','Método de Pago',1)}</div>
<div class="card"><h3>🎓 Alumnas por estado</h3>${barras(group(al.map(a=>({...a,Estado:a.Estado||'Sin pagos'})),'Estado'),'alumnas','Estado')}</div>
<div class="card"><h3>🏆 Alumnas que más han pagado</h3>${barras(al.map(a=>[nom(a),num(a['Total pagado ($)'])]).sort((a,b)=>b[1]-a[1]),'alumnas','',1)}</div>
<div class="card"><h3>🛒 Productos más vendidos <small class="mu">(unidades)</small></h3>${barras(Object.entries(masV).sort((a,b)=>b[1]-a[1]),'ventas','',0)}</div>
<div class="card"><h3>👗 Inventario por estado</h3>${barras(group(st,'Estado'),'stock','Estado')}</div>
<div class="card"><h3>🏢 Activos por estado</h3>${barras(group(ac,'Estado','Valor Total ($)'),'activos','Estado',1)}</div></div>
<p class="mu" style="color:#cbd5e1">Toca cualquier tarjeta o barra para ver el detalle filtrado.</p>`}
function vista(id){const t=T[id],s=S[id]??={q:'',f:'',col:'',mes:'',sc:null,asc:1,n:100,vista:''};
if(id==='cont')return contable();if(id==='bal')return balance();
let {cols,rows}=tabla(id);if(t.orden)cols=[...t.orden.filter(c=>cols.includes(c)),...cols.filter(c=>!t.orden.includes(c))];
const fc=t.fc,colF=s.col||(cols.includes(fc)?fc:''),fcol=cols.find(c=>/^Fecha/i.test(c));
const dinero=cols.filter(c=>/\(\$\)|precio|valor/i.test(c)&&rows.some(x=>typeof x[c]==='number'));
// columnas que sirven para filtrar: pocas opciones distintas
const filtrables=cols.filter(c=>{const u=new Set(rows.map(x=>x[c]).filter(v=>v!=null&&v!==''));return u.size>1&&u.size<=40&&!dinero.includes(c)&&c!==fcol});
const opts=colF?[...new Set(rows.map(x=>x[colF]).filter(v=>v!=null&&v!==''))].map(String).sort():[];
const meses=fcol?[...new Set(rows.map(x=>String(x[fcol]||'').slice(0,7)).filter(Boolean))].sort().reverse():[];
let r=rows.filter(x=>(!s.f||!colF||String(x[colF])===s.f)&&(!s.mes||String(x[fcol]||'').startsWith(s.mes))
 &&(!s.q||Object.values(x).join(' ').toLowerCase().includes(s.q.toLowerCase())));
if(s.sc)r.sort((a,b)=>((a[s.sc]??'')>(b[s.sc]??'')?1:-1)*s.asc);
const C=t.foto?['Foto',...cols]:cols;
const cell=(c,v,x)=>c==='Foto'?foto(x[t.foto]):v==null?'':typeof v==='number'&&dinero.includes(c)?fmt(v)
 :/^(Varias|No encontrada)/.test(String(v))?`<span class="chip r">${esc(v)}</span>`:c==='Estado'||c===colF?`<span class="chip${/Agotado/.test(v)?' r':/Última/.test(v)?' a':''}">${esc(v)}</span>`:esc(v);
const tarjetas=(s.vista||(innerWidth<720?'t':'l'))==='t';
const hayF=s.q||s.f||s.mes;
const nomMes=m=>{const[y,mm]=m.split('-');return['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'][+mm-1]+' '+y};
const tk=['Descripción','Concepto / Descripción','Alumna','Nombre del Activo','Concepto','Nombre'].find(c=>cols.includes(c))||cols[0];
const titulo1=x=>esc(x[tk]??'')+(tk==='Nombre'&&x.Apellidos?' '+esc(x.Apellidos):'');
const lado=x=>dinero.length&&typeof x[dinero[dinero.length-1]]==='number'?`<span style="white-space:nowrap;font-weight:800;color:${id==='gastos'?'var(--bad)':'var(--ok)'}">${fmt(x[dinero[dinero.length-1]])}</span>`:'';
return `<div class="card"><div class="tb"><input placeholder="🔍 Buscar en ${t.n}…" value="${esc(s.q)}" oninput="S['${id}'].q=this.value;S['${id}'].n=100;repintar('${id}')" id="q">
${meses.length?`<select aria-label="Mes" onchange="S['${id}'].mes=this.value;pintar()"><option value="">📅 Todos los meses</option>${meses.map(m=>`<option value="${m}" ${m===s.mes?'selected':''}>${nomMes(m)}</option>`).join('')}</select>`:''}
${filtrables.length?`<select aria-label="Filtrar por" onchange="S['${id}'].col=this.value;S['${id}'].f='';pintar()">${filtrables.map(c=>`<option value="${esc(c)}" ${c===colF?'selected':''}>Filtrar por: ${esc(titulo(c))}</option>`).join('')}</select>
<select aria-label="Valor" onchange="S['${id}'].f=this.value;pintar()"><option value="">Todos</option>${opts.map(o=>`<option ${o===s.f?'selected':''}>${esc(o)}</option>`).join('')}</select>`:''}
</div><div class="res"><span class="chip">${r.length} de ${rows.length} registros</span>${dinero.map(c=>`<span class="chip g">${esc(titulo(c))}: <b>${fmt(sum(r,c))}</b></span>`).join('')}
<span class="sp"></span>${hayF?`<button class="lnk" onclick="S['${id}']={...S['${id}'],q:'',f:'',mes:''};pintar()">✕ Limpiar filtros</button>`:''}
<button class="lnk" onclick="S['${id}'].vista='${tarjetas?'l':'t'}';pintar()">${tarjetas?'☰ Ver tabla':'▦ Ver tarjetas'}</button></div>
${tarjetas?`<div class="cards">${r.slice(0,s.n).map(x=>`<div class="it">${t.foto?`<div class="itf">${foto(x[t.foto])}</div>`:''}<div class="itb"><div style="display:flex;justify-content:space-between;gap:8px"><b>${titulo1(x)}</b>${lado(x)}</div>
<dl>${C.filter(c=>c!=='Foto').map(c=>x[c]==null||x[c]===''?'':`<div><dt>${esc(titulo(c))}</dt><dd>${cell(c,x[c],x)}</dd></div>`).join('')}</dl></div></div>`).join('')||'<p class="mu">Sin resultados con estos filtros.</p>'}</div>`
:`<div class="tw"><table><thead><tr>${C.map(c=>`<th onclick="orden('${id}',this.dataset.c)" data-c="${esc(c)}">${esc(titulo(c))}${s.sc===c?(s.asc>0?' ▲':' ▼'):''}</th>`).join('')}</tr></thead>
<tbody>${r.slice(0,s.n).map(x=>`<tr>${C.map(c=>`<td>${cell(c,x[c],x)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`}
${r.length>s.n?`<button class="more" onclick="S['${id}'].n+=200;pintar()">Ver más (${r.length-s.n} restantes)</button>`:''}</div>`}
// ---- Balance mes a mes: operativo vs total; la inversión inicial se muestra aparte y no se resta ----
function balance(){const M={},add=(m,k,v)=>{if(!m||!v)return;(M[m]??={cl:0,ve:0,op:0,rep:0,pre:0,otr:0,inv:0})[k]+=v};
const mes=v=>String(v||'').slice(0,7),INV=/inversi[oó]n inicial/i;
tabla('ingresos').rows.forEach(r=>add(mes(r.Fecha),'cl',num(r['Total ($)'])));
tabla('ventas').rows.forEach(r=>add(mes(r.Fecha),'ve',num(r['Total Venta ($)'])));
tabla('gastos').rows.forEach(r=>{const c=String(r['Clasificación Contable']||'');add(mes(r.Fecha),INV.test(c)?'inv':/^operativo/i.test(c)?'op':/preoperativo/i.test(c)?'pre':'otr',num(r['Valor ($)']))});
tabla('merc').rows.forEach(r=>{const c=String(r['Clasificación']||'');add(mes(r.Fecha),INV.test(c)?'inv':/reposici/i.test(c)?'rep':'otr',num(r['Costo Total ($)']))});
tabla('activos').rows.forEach(r=>add(mes(r['Fecha Adquisición']),INV.test(String(r['Clasificación']||''))?'inv':'otr',num(r['Valor Total ($)'])));
const ms=Object.keys(M).sort();let aO=0,aT=0;
const L=ms.map(m=>{const x=M[m],e=x.cl+x.ve,gO=x.op+x.rep,gT=gO+x.pre+x.otr;aO+=e-gO;aT+=e-gT;return{m,...x,e,gO,gT,bO:e-gO,bT:e-gT,aO,aT}});
if(!L.length)return '<div class="card">No hay movimientos con fecha.</div>';
const v=S.balv||'tot',G=v==='op'?'gO':'gT',B=v==='op'?'bO':'bT',A=v==='op'?'aO':'aT';
const nomMes=m=>{const[y,mm]=m.split('-');return['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][+mm-1]+' '+y.slice(2)};
const tot=k=>L.reduce((s,x)=>s+x[k],0),mx=Math.max(...L.map(x=>Math.max(x.e,x[G])),1);
const st=Math.max(90,Math.floor(1000/L.length)),W=Math.max(L.length*st+40,300),H=210,vals=L.map(x=>x[A]),aMin=Math.min(0,...vals),aMax=Math.max(0,...vals),rg=(aMax-aMin)||1;
const px=i=>50+i*st,py=y=>30+(aMax-y)/rg*(H-70),Mf=y=>(y<0?'-':'')+'$'+(Math.abs(y)/1e6).toFixed(1).replace('.',',')+' M';
const svg=`<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="Acumulado"><line x1="0" x2="${W}" y1="${py(0)}" y2="${py(0)}" stroke="#94a3b8" stroke-dasharray="4 4"/><text x="4" y="${py(0)-5}" font-size="10" fill="#64748b">$0</text>
<polyline fill="none" stroke="#1e3a8a" stroke-width="3" points="${L.map((x,i)=>px(i)+','+py(x[A])).join(' ')}"/>
${L.map((x,i)=>`<circle cx="${px(i)}" cy="${py(x[A])}" r="5" fill="${x[A]<0?'#ef4444':'#22a94b'}"/><text x="${px(i)}" y="${py(x[A])-11}" font-size="10.5" text-anchor="middle" fill="#1e1b4b" font-weight="700">${Mf(x[A])}</text><text x="${px(i)}" y="${H-4}" font-size="11" text-anchor="middle" fill="#64748b">${nomMes(x.m)}</text>`).join('')}</svg>`;
const K=(t,val,c,sub)=>`<div class="card kpi" style="cursor:default"><h4>${t}</h4><b class="${c}">${fmt(val)}</b><small class="mu" style="display:block">${sub}</small></div>`;
const btn=(k,t)=>`<button class="seg${v===k?' on':''}" onclick="S.balv='${k}';pintar()">${t}</button>`;
return `<div class="card" style="margin-bottom:14px;display:flex;flex-wrap:wrap;gap:10px;align-items:center"><b>Ver el balance con:</b>${btn('op','Solo gastos operativos')}${btn('tot','Gastos totales')}
<span class="mu">En ninguno se resta la inversión inicial.</span></div>
<div class="g k">${K('Entradas',tot('e'),'pos','clases + ventas')}${K(v==='op'?'Gastos operativos':'Gastos totales',tot(G),'neg',v==='op'?'operativos + reposición':'operativos + preoperativos + otros')}
${K('Balance '+(v==='op'?'operativo':'total'),tot(B),tot(B)<0?'neg':'pos','acumulado hasta hoy')}${K('Inversión inicial',tot('inv'),'',"aparte: no se resta")}</div>
<div class="card" style="margin-bottom:14px"><h3>📊 Entradas vs ${v==='op'?'gastos operativos':'gastos totales'} <small class="mu">(🟢 entradas · 🔴 gastos)</small></h3>
<div class="mes" style="height:200px">${L.map(x=>`<div onclick="S.balm='${x.m}';pintar()" style="cursor:pointer"><div class="bb" style="height:160px"><i title="Entradas ${fmt(x.e)}" style="height:${x.e/mx*100}%;background:var(--p)"></i><i title="Gastos ${fmt(x[G])}" style="height:${x[G]/mx*100}%;background:var(--bad)"></i></div><b class="${x[B]<0?'neg':'pos'}" style="font-size:.68rem">${fmt(x[B])}</b><br>${nomMes(x.m)}</div>`).join('')}</div></div>
<div class="card" style="margin-bottom:14px"><h3>📈 Acumulado ${v==='op'?'operativo':'total'}</h3><div style="overflow-x:auto">${svg}</div></div>
<div class="card"><h3>🧾 Detalle mes a mes</h3><div class="tw"><table class="bal"><thead><tr><th>Mes</th><th class="e">Entradas</th><th>Gastos operativos</th><th>Reposición mercancía</th><th>Gastos preoperativos</th><th>Otros</th><th class="s">Total gastos</th><th class="iv">Inversión inicial</th><th class="b">Balance operativo</th><th class="b">Balance total</th><th>Acum. operativo</th><th>Acum. total</th></tr></thead>
<tbody>${L.map(x=>`<tr><td><b>${nomMes(x.m)}</b></td><td class="e">${fmt(x.e)}</td><td>${fmt(x.op)}</td><td>${fmt(x.rep)}</td><td>${fmt(x.pre)}</td><td>${fmt(x.otr)}</td><td class="s">${fmt(x.gT)}</td><td class="iv">${fmt(x.inv)}</td><td class="b ${x.bO<0?'neg':'pos'}">${fmt(x.bO)}</td><td class="b ${x.bT<0?'neg':'pos'}">${fmt(x.bT)}</td><td class="${x.aO<0?'neg':'pos'}">${fmt(x.aO)}</td><td class="${x.aT<0?'neg':'pos'}">${fmt(x.aT)}</td></tr>`).join('')}
<tr class="tt"><td>Total</td><td class="e">${fmt(tot('e'))}</td><td>${fmt(tot('op'))}</td><td>${fmt(tot('rep'))}</td><td>${fmt(tot('pre'))}</td><td>${fmt(tot('otr'))}</td><td class="s">${fmt(tot('gT'))}</td><td class="iv">${fmt(tot('inv'))}</td><td class="${tot('bO')<0?'neg':'pos'}">${fmt(tot('bO'))}</td><td class="${tot('bT')<0?'neg':'pos'}">${fmt(tot('bT'))}</td><td></td><td></td></tr></tbody></table></div>
<p class="mu"><b>Balance operativo</b> = entradas − gastos operativos − reposición de mercancía (¿el día a día se sostiene?). <b>Balance total</b> = entradas − todos los gastos (operativos, preoperativos y otros). Lo marcado <b>"Inversión Inicial"</b> en Gastos, Activos o Marcancia va en su propia columna y no se resta en ninguno. Toca una barra para ver el detalle del mes.</p></div>
${S.balm?detalleMes(S.balm):''}`}
function detalleMes(m){const NM=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'][+m.slice(5,7)-1]+' '+m.slice(0,4);
const f=r=>String(r.Fecha||r['Fecha Adquisición']||'').startsWith(m),INV=/inversi[oó]n inicial/i;
const ga=tabla('gastos').rows.filter(f),ing=tabla('ingresos').rows.filter(f),ve=tabla('ventas').rows.filter(f),me=tabla('merc').rows.filter(f),ac=tabla('activos').rows.filter(f);
const et=c=>INV.test(c)?'inv.':/^operativo/i.test(c)?'op.':/preoperativo/i.test(c)?'preop.':'otro';
const gastos=group(ga.filter(r=>!INV.test(String(r['Clasificación Contable']||''))).map(r=>({...r,k:(r['Categoría']||'(sin categoría)')+' ('+et(String(r['Clasificación Contable']||''))+')'})),'k','Valor ($)');
const rep=sum(me.filter(r=>/reposici/i.test(String(r['Clasificación']||''))),'Costo Total ($)');if(rep)gastos.push(['Reposición mercancía (op.)',rep]);
const invI=[['Gastos marcados Inversión Inicial',sum(ga.filter(r=>INV.test(String(r['Clasificación Contable']||''))),'Valor ($)')],['Activos fijos',sum(ac.filter(r=>INV.test(String(r['Clasificación']||''))),'Valor Total ($)')],['Mercancía inicial',sum(me.filter(r=>INV.test(String(r['Clasificación']||''))),'Costo Total ($)')]].filter(x=>x[1]);
const entradas=[...group(ing,'Tipo','Total ($)').map(([k,v])=>['Clases: '+k,v]),['Ventas de mercancía',sum(ve,'Total Venta ($)')]].filter(x=>x[1]).sort((a,b)=>b[1]-a[1]);
const lista=(it,c)=>{it=it.sort((a,b)=>b[1]-a[1]);const mx=Math.max(...it.map(i=>i[1]),1);return it.map(([k,v])=>`<div class="br" style="cursor:default"><span title="${esc(k)}">${esc(k)}</span><i style="width:${v/mx*45}%;background:${c}"></i><em>${fmt(v)}</em></div>`).join('')||'<p class="mu">Sin movimientos</p>'};
return `<div class="g c2" style="margin-top:14px"><div class="card"><h3>🟢 Qué entró en ${NM}</h3>${lista(entradas,'var(--p)')}</div><div class="card"><h3>🔴 Gastos de ${NM} <small class="mu">(op. = operativo · preop. = preoperativo)</small></h3>${lista(gastos,'var(--bad)')}</div>
${invI.length?`<div class="card"><h3>🟣 Inversión inicial de ${NM} <small class="mu">(no se resta)</small></h3>${lista(invI,'#7c3aed')}</div>`:''}</div>`}
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
iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAB4FBMVEX//////v/+///+//7+/v/+/v79//79/v79/v38/v37/vv9
/f78/f38/P36/Pzz+fXv8/Xh7urM7dXU3+TEzduc36qTzLRlzH1AxFU3wkk2wUk2wUc3wkU2wkY1wkQ2wE4yv1AvvlMzwEctv0CY
rsBEs3kuvVYtuFwptmQnsmsoulAlq20komUhs1ohqEshp0shpnwhpkwhpGwhpE0hpUwhnlMhoFAhn1AhnlAhoU8hoE8ho04ho00h
ok4hoU4bqV2BlrNrf6ZEfo0kmW4ljXQkeHghmnYhmFQhl1QhmlMhmVMhnVEhnFEhm1Ihk1chlFYhllUhkVghj1khiF8Yl1odjVwf
iWEfi1ogg2AchHsfh1wffGsef2Aed2oeb3IccWYeaXEvXoMfZHEfYXMfXXUfWncfWHgfVnkeWXotUIMeU3seUXweT30eTn4eTH8e
S4AXXXMYT3YkRIQeSYEeSIEeSoAeR4IeRYMeRIQeRIMeRoIeQoUeQYUeQIYeQIUeQ4QeQoQeQnIYRXYhPoIeP4cePocePYgePIge
P4YeO4keOokiOncfN3UYOnMSN3MUMm8SMXASLm4RLW8QLG4QK20OLG0OKmwPKWwOKWsOKGoNKGoNJ2oLJ2kMJmgFI2YBG2GhgLd7
AAAc2UlEQVR42tWd/X/aVpbG7yITwICQQICJ2zixIY7xeNebbtJpM5MgmmRmMunLTvLZjNPUCbW7CUlfsNOXmWY7bZh2261tMDb4
BWPxr+45VxJISFcSjvNDnjqNDUb6cvTcc869unbIv76CIq8k9L+9giL//grq1YT+j5ehN/TDv/FSDk/eOGZduvTGpTd6rPA1PnLM
IsdIe+mN32m6NK3pkv4IffrYoH/34gKeSxprLp1KxeNxMaZJhC9SqXROo+9954uJXDoGIQ7gAithCOgB/ZLG/aIiv38hXXrr8uXL
v8+lkyJP4QJBng+FQgFOUwC+4PlggD7Ji8l07vfwgrcuvdhZjwz91lvwAQAXcilxFJGCo8DKijTQjwbxk1ExlbuA3PjyI0O/dVRd
vHj58tu5pIgoEFsfcZUP4o5/i8nc25cvX7x45FOTt4+mi3BaiDE6whOwARy9AvGGt3zxiCcnF4fX2xcvvnn57XQcbMyF/GRo+UNg
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
""",
    "192": """
iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAMAAABlApw1AAAB4FBMVEX///////7+///+//7+//3+/v/+/v79//79/v79/v38/v37
/vz9/f78/f38/P36/Pvy+PXt8PTc7uXD6s7K3N6+ydiR26GEyqVWyW85wkw2wUs2wUc3wkU2wkY1wkQ2wUYyv1AvvlMwwESXrsBA
tG8tvFguvVMsuF8oulEnsmglqWskoWghslshqHchqEshp00hp0shpk0hpHghpU0hpE0hpUwhoHohnlEhnVEhoFAhn1AhnlAhoU8h
oE8ho04hok4aplxej5xRc5ckknIkdnghl1chmFQhl1QhmlMhmVMhnFIhm1IhkVkhklchlVYhlFYhj1ohkFghhWAak18gi10fil4f
hGYfgWYfhlwffG0dfmAgdmgbdnQebXMdcGgdaHAwYIIfY3IfYHQfXXYsUoMfWXgfV3kfVXoeVXoeU3seUnweUH0eT30eTX4eTH8e
S4AkRIQeSoAeSYEeSIIeR4IeRoMeRYMeRIQeQoUeQYUeQIYeQ4QeQoQeQHYYXnMYUnYXRXUgP4geP4cePocePYgePYcePIgeP4Ue
O4keOokjO3gfOHYZO3QUOnMUNHETMG8RMG8RLm4QLW4RLG4QKmwOLG4OKmwOKWwMKWoOKGsNKGoMJ2kMJmgJJ2gGImYAGWC1cuSf
AAAew0lEQVR42tWd+WPTZprH37U3iu34kGTLR1KOWoEkEPZoZ5hhKIcUWMq23S6dzjalhLrEXB3apgft7rTEZna77CSxjfERJ8Hy
v7rP80qyJVuvJCfhB56hFOzE+n7e53xfKR3yD6+5kdce4J9ecyP/+Job+efX3MhvX3Mjb78iG16pV3Ud8pvDtYsg/eLbv/ntb4ZX
Cl56+yJgXDzkC5JDFY/S0d6em8tms5m0YRn4y9zc2zoIfNFhQpCLh2co/uIs6E6KPE+GjOfFJJDMXkSIQ7zoIQBcor9d+u2luSxI
jxiCQ+FwZMqwSDgcMl6OAEZ2Dr74Uv9bDwhw6aD2zjug/tIiiE9QieGpSCQUDA57IBgMRSJTYfrnBEAswjddfOedA1+evHMwO3fu
0oVLc5kkb2jnAsTVApxBwSczc/Ct584dUAA5mPpzFy4Y6sNT4SDxacEwhUCGCxfOHYyBXNi3ofoL2bSAkR3xLb4PEcFsEdJZ85P2
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
tQf4f+/k2ca6P6EpAAAAAElFTkSuQmCC
""",
    "512": """
iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAMAAADDpiTIAAAB4FBMVEX///////7//v/+///+//7+/v/+/v79/v38/f3y+fXg6uvC
29SH059DxFg3wkY2wkY2wUk2wUc2wkU2wEwzv08zwEswv1I1wUgzwEeLr7YzuGAuvFcsu1kstl8ot2MmsWslqnMkom0hr3MhqHoh
qEshp00hp0shpnghpkwhpHkhpU0hpE0hpUwhoHkhnlEhnVEhoFAhn1AhnlAhok8hoU8hoE8ho04ho00hok4hoU4dpWg3mIIhmnch
lFYhllUhlVUhmFQhl1QhnFIhm1IhmlMhmVNcfJ0ng38je3ohkVghk1chjlkhgGYekGMgi10giGEgilwfgWsgfmcghV4fgl8fd24f
e2QfcHAfc2gfa24faHAcankfZHIcZHktVoMfYHQfXXUfW3cfWHgfVXkeVXoeU3seUXweUH0eTn4eTH8eS4AeTX4eSoAeSYEeSXYk
Q4IeR4IeRoIeRYMeSIEeRIQeQ4QeRIMeQYYeQIYeP4ceP4YeQoUeQYUeQIUeQX4aVHgZRHUjPXgePYgePIgePocePYcePXMeO4ke
OokeO4cfOHYYOnEUN3IUM24RMnESL24RLm4RLW4QLW4QLG4QK20PKm0PKmwPKWwOK2wOKWwOKGsNKGoNJ2kMJ2kMJmgII2aVQE8t
AABfl0lEQVR42u29iWPjxpX/iRgkCAJUt9Tqbgnq085ms/ntOPHYm2QySuIjMdSJPY4z49jxxbbdl28lcWYyScat1u5M71qiRFE8
xEMk/9Wt96oAFECAxFEAKYkvTutoNSXhfer7jrqk/21uZ9qk+SOYAzC3OQBzmwMwt7MJwP8+tzNtcwDmAMztTAPwnbmdaZP+j7md
aZsDMAdgbnMA5jYHYG5nFIB/mNuZtjkAZx2A/zW3M21zAOYAzO1MA/B/zu1M25kF4EcemwNw+j0Oived73wH3/6vH3sMP/kj8tf4
9kdzAE6P48Gv5I8f/8M/8A7/zjWXfYf/O/KV7B+dfhCkH59ao9LuePy5tYWFhRIxHUwrukzDT8Lfki9ae+47Qa9y6kz60WkzdNlz
zz3HnPbcd9YW0OnE41JIQx6AhbXvPMcwgBe0XvxU2ekDgLjqn6jrwfPE7wXbsYVCQSVWAFNyLlPwk+xvnX9ASAAOKAb/xCiYAzC7
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
""",
    "m192": """
iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAMAAABlApw1AAAB4FBMVEX///////7+///+//7+//3+/v/+/v79//79/v79/v38/v37
/vz9/f78/f38/P36/Pvy+PXt8PTc7uXD6s7K3N6+ydiR26GEyqVWyW85wkw2wUs2wUc3wkU2wkY1wkQ2wUYyv1AvvlMwwESXrsBA
tG8tvFguvVMsuF8oulEnsmglqWskoWghslshqHchqEshp00hp0shpk0hpHghpU0hpE0hpUwhoHohnlEhnVEhoFAhn1AhnlAhoU8h
oE8ho04hok4aplxej5xRc5ckknIkdnghl1chmFQhl1QhmlMhmVMhnFIhm1IhkVkhklchlVYhlFYhj1ohkFghhWAak18gi10fil4f
hGYfgWYfhlwffG0dfmAgdmgbdnQebXMdcGgdaHAwYIIfY3IfYHQfXXYsUoMfWXgfV3kfVXoeVXoeU3seUnweUH0eT30eTX4eTH8e
S4AkRIQeSoAeSYEeSIIeR4IeRoMeRYMeRIQeQoUeQYUeQIYeQ4QeQoQeQHYYXnMYUnYXRXUgP4geP4cePocePYgePYcePIgeP4Ue
O4keOokjO3gfOHYZO3QUOnMUNHETMG8RMG8RLm4QLW4RLG4QKmwOLG4OKmwOKWwMKWoOKGsNKGoMJ2kMJmgJJ2gGImYAGWC1cuSf
AAAew0lEQVR42tWd+WPTZprH37U3iu34kGTLR1KOWoEkEPZoZ5hhKIcUWMq23S6dzjalhLrEXB3apgft7rTEZna77CSxjfERJ8Hy
v7rP80qyJVuvJCfhB56hFOzE+n7e53xfKR3yD6+5kdce4J9ecyP/+Job+efX3MhvX3Mjb78iG16pV3Ud8pvDtYsg/eLbv/ntb4ZX
Cl56+yJgXDzkC5JDFY/S0d6em8tms5m0YRn4y9zc2zoIfNFhQpCLh2co/uIs6E6KPE+GjOfFJJDMXkSIQ7zoIQBcor9d+u2luSxI
jxiCQ+FwZMqwSDgcMl6OAEZ2Dr74Uv9bDwhw6aD2zjug/tIiiE9QieGpSCQUDA57IBgMRSJTYfrnBEAswjddfOedA1+evHMwO3fu
0oVLc5kkb2jnAsTVApxBwSczc/Ct584dUAA5mPpzFy4Y6sNT4SDxacEwhUCGCxfOHYyBXNi3ofoL2bSAkR3xLb4PEcFsEdJZ85P2
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
tQf4f+/k2ca6P6EpAAAAAElFTkSuQmCC
""",
    "m512": """
iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAMAAADDpiTIAAAB4FBMVEX///////7//v/+///+//7+/v/+/v79/v38/f3y+fXg6uvC
29SH059DxFg3wkY2wkY2wUk2wUc2wkU2wEwzv08zwEswv1I1wUgzwEeLr7YzuGAuvFcsu1kstl8ot2MmsWslqnMkom0hr3MhqHoh
qEshp00hp0shpnghpkwhpHkhpU0hpE0hpUwhoHkhnlEhnVEhoFAhn1AhnlAhok8hoU8hoE8ho04ho00hok4hoU4dpWg3mIIhmnch
lFYhllUhlVUhmFQhl1QhnFIhm1IhmlMhmVNcfJ0ng38je3ohkVghk1chjlkhgGYekGMgi10giGEgilwfgWsgfmcghV4fgl8fd24f
e2QfcHAfc2gfa24faHAcankfZHIcZHktVoMfYHQfXXUfW3cfWHgfVXkeVXoeU3seUXweUH0eTn4eTH8eS4AeTX4eSoAeSYEeSXYk
Q4IeR4IeRoIeRYMeSIEeRIQeQ4QeRIMeQYYeQIYeP4ceP4YeQoUeQYUeQIUeQX4aVHgZRHUjPXgePYgePIgePocePYcePXMeO4ke
OokeO4cfOHYYOnEUN3IUM24RMnESL24RLm4RLW4QLW4QLG4QK20PKm0PKmwPKWwOK2wOKWwOKGsNKGoNJ2kMJ2kMJmgII2aVQE8t
AABfl0lEQVR42u29iWPjxpX/iRgkCAJUt9Tqbgnq085ms/ntOPHYm2QySuIjMdSJPY4z49jxxbbdl28lcWYyScat1u5M71qiRFE8
xEMk/9Wt96oAFECAxFEAKYkvTutoNSXhfer7jrqk/21uZ9qk+SOYAzC3OQBzmwMwt7MJwP8+tzNtcwDmAMztTAPwnbmdaZP+j7md
aZsDMAdgbnMA5jYHYG5nFIB/mNuZtjkAZx2A/zW3M21zAOYAzO1MA/B/zu1M25kF4EcemwNw+j0Oived73wH3/6vH3sMP/kj8tf4
9kdzAE6P48Gv5I8f/8M/8A7/zjWXfYf/O/KV7B+dfhCkH59ao9LuePy5tYWFhRIxHUwrukzDT8Lfki9ae+47Qa9y6kz60WkzdNlz
zz3HnPbcd9YW0OnE41JIQx6AhbXvPMcwgBe0XvxU2ekDgLjqn6jrwfPE7wXbsYVCQSVWAFNyLlPwk+xvnX9ASAAOKAb/xCiYAzC7
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
""",
}


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))