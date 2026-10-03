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
            {"src": "/icono/192.png?v=2", "sizes": "192x192", "type": "image/png", "purpose": "any"},
            {"src": "/icono/512.png?v=2", "sizes": "512x512", "type": "image/png", "purpose": "any"},
            {"src": "/icono/m192.png?v=2", "sizes": "192x192", "type": "image/png", "purpose": "maskable"},
            {"src": "/icono/m512.png?v=2", "sizes": "512x512", "type": "image/png", "purpose": "maskable"},
        ]}, media_type="application/manifest+json")


@app.get("/sw.js")
def sw():
    # Guarda la última versión de la app y de los datos: si no hay internet, abre con lo último que vio.
    js = """const C='rojas-v13';
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
<link rel="icon" type="image/png" sizes="192x192" href="/icono/192.png?v=2"><link rel="icon" href="/favicon.ico">
<link rel="apple-touch-icon" sizes="180x180" href="/icono/180.png?v=2">
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
<header><img src="/icono/192.png?v=2" alt="">
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
iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAABgFBMVEX////+///+//79//7+/v/+/v79/v79/v39/f78/v38/f38
/fz6/fv8/P34+/ry+fXv8vXg9OXh5+3J7NWs5bfU2uW21tC4xNSossqS26R004tMzl45y0t5ubBJwXE6xUw5v1Y1x1E0w081wk03
wkY2wkY1wkY3wkU2wUo2wUgzwUw2wUc1wUczv1Ewv1Q2wEszwEwwvFsuvFYwvVUuvVUwvlIpumYrulwtu1grtmcrtl0mtGsorXkn
rmUjrXWOnLpXlp0mongmkYAjpXcinngil3sgjn1oe6NCdJIlgoEkc4Eeg34eeYAfb38faH4/W4sgXH0dXnwcVHouSIAdS3gaSncZ
Q3QgPXYYPXIXPXIYOHEXNHEUN3IUNG8UMm8RMnEUMG8SMG8SLm8RMHERLm8RLW8QLW8RK24QLG8QKm4QK20PLHAPK24PK2wPKm0O
Km4OKmwPKW0OKW0PKWoOKWsOKGsNKGwNKGoNJ2oMJ2kMJ2gNJmkMJmgMJmcMJWgLJ2nvr5UbAAAbaUlEQVR42u2ci18Tx9fGh4QN
AXOHkAshIQV/xVK5iKWKSFBrN9rqqsFGSZSQAIEEyIUQLtrXf/19zsxu7ggJ0ZbPp6cSLobd7555zjNnZrcy9l/8F/9FXej6GBv0
+kaGTLrWf2/y+H5E+EbsUtM7dDq71xfwWFlf3/dkNjBm905M4LznXRQzeUE9MWJnLcF0QyO+kUHGviN1HzN5Avyk+vPeomf2kYkf
fUOsv/U1MSuO4LWyft33UYauH2n2+ZBmU79Bd/6FQSAjlj7deX/NBke+ft3dy7GeBhSn8w7W/KgVmF5nD/g854+/joYLyra0Houu
hcTpTNYhSGMIMTg4aBeq1uv1zeXGvD6P1Pc1kfGrt39TahpH+5B3JOCbgC3w+DEwMuIdGuTkjfnuI+ivDr6OhE8S0eu+qckFQDox
Aegf6RX/TdAPAG43NWKTqOuhdT2NcP3MCvMbQp18qzTDlwl4YurHSkzw+HGCwD32egdrhib9GxqHw0DC/jbUemaBRQH4JsWPDUHg
Pt8I/ETSnwctqQ6vb7gMHd7m8xhY96kNJL6JiZtqCNImcJRnjUaaMq3TcXB9n6HeQYna2991as48OXXzZi11oJZ6QiiFvEDTiL5Z
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
""",
    "192": """
iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAMAAABlApw1AAABgFBMVEX////+///+//79//7+/v/+/v79/v79/v39/f78/v38/f36
/fv8/P37/P37+/31+/fu9/Lq7vLV8dvY4+e66Mib36nG0t2zvdGD1J6Ts8BdzHZFx1g5y0w4xkpPuoU4wUo3wUk3wkUzxlM2wkk2
wkY2wkU0wkw0wU01wUk1wUg2wUcyv1Mzv042wEszwEwtvlsuvVUvvlMvul4uu1guvFUouGkquV8rulsqtGsqtF8lsm8nrHwmrGgj
q3WNm7lPlpkmoXkmkIAjpHYinXghlXsgjH5oe6NMb5YmgYIlcoIegn8ed38fbn8fZ34/WYsgW30dXXwcVHowSYAdS3caSncZQ3Qc
PXQXPXIYOHAXNXAUN3EUNHAUMm8RMnIUMG8SMG8SLm4RMHARLm4TLW8RLW8SLW0QLW8QLHAQLG0QK28QK20PK28PK20PKm0QKmwO
Km0PKW0PKWoOKW0OKWsOKGsOKGkNKG0NKGsNKGoNJ2oMJ2kMJ2gMJmgMJmcMJWgLJ2mRxThHAAAetklEQVR42u2di18S2fvHj+Bw
E+V+FdS0tDRXN7PMtqH61kztFhUViQZeQLnlgKgoZr/+9d/znDMDA4yJSLW+XvuUeImFz/s813Nmagn5z/6z/+zyxukJcfpDYY9d
x2k/w2D2hK9du3bjWthv03iOTmf3T4SDTkKMv0M/vKndPxzC9z8TkTP6Q0AQDg4QnfZT3MHQsMdG9LrfoN/sGQ6F3LDOZ7+5jtiC
4INhO9Fr/7mOGOFVgvAq+l+sX08cwRAuv8H4Y0x3eCzkOTtGjDQMR8EJ5l8X+waj3kycw+B7psGoN57lBZ3ONhweder0P8wkN41E
468JIz3T4obld9vsdtuAIlWv5zRX2BOCCPqRmwDeCZmAwfgLIh9Xyen2BK9DZFMLBoN+j9s+wHzRtoh6zn0eAIHosUE18//8MMJI
cHqCo6FQ6Mb0tXAYfofD8E0oHA763U69QtiSBACg/qlGQQXhnnAoaP/JBJhwQRB8/fr09Rs3rl2bnr5WN/gxQthaEYzE2QoASWTg
2peGJcJPJOCMUPfDoevT09PXr1+/gQ3qBmLU7TowDHvszW1JCwD0Gs0tP+FYWXD/PAKOFgumHu2apoWhNA0QvVEdQqEmADmAdMbW
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
""",
    "512": """
iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAMAAADDpiTIAAABgFBMVEX////+//7+/v/+/v7+/v3+/f78/f33+vro9+zd7ubU2+Wt
2sJv0IZBxFU3xEc3wkc3wkU3wkQ2w0g2wkg2wkY1wkY2wkWXqsBAumQ2wUo2wUg2wUc2wUY2v0w1wUk1wUg1wUc1v040wUozwEwz
v08ywE4xv08wv1IwvlUxvlEuvVYuvVQuul0tulkqul4puGAwvFcuvFUuu1Ytu1crtWcrtV0ntWclsmsnrnYnrmUjrnMlp3UmoXki
pngioHYjmnsjkn1YfZ0lgoEkfIEeiH0efn8edX8fbn4yY4YgX30dZX0cW3sqUX8cT3gaTXcjQngZQnMXQHMZOnEYNm8UN3AUNG4U
M24UMm4UL28SMW8SL24SLm4QMHARLm4RLW4QLW4QLG0QK24PLG0PK20QKm0PKm4PKm0PKW0PKWwOKmwOKW0OKWwNKWwOKWsPKGsO
KGsOKGoNKGsNKGoNKGkOJ2oNJ2kNJmgMKGoMJ2kLJ2oMJ2gMJmkMJmgMJmcMJWgLJmhG8VFuAABqfUlEQVR42u2diUNTx/b4E65k
BfEr/nqJiktrW7dWrdba1r4afa8vr42NTVQiSEgkhIQQViOLtP7rv3POzNw7c5fswA3JaQWCqHDPZ84+Mz7fUIYylKEMZShDGcpQ
hjKUoQxlKEMZylCGMpSh9IH4/T4tHB0/fWZi4vSZ8Wg42Iu/MzQ2/n+SfAZyEWRiLBro8K/UgvBNnj0bi8XGx8ei4cBQcT1Tvy8U
HTsN+j9zenysJ+pHUQn4jBEQG+tKccHw2PjZS7FLsfGpsWjIP1Re79UP6tEOgwBmACbGwt387fBng1GBwBiZKm2owm5EY+oH639m
4jRY5xAQ0atHqvnCY6f/7//OSACcGYv6e/H9xtAPxJgfGBLQzfJnK+r0uYu4/KPhHqof/3YtPKYCMB4N9MJioR+YjF0iKzA0Al2t
0UAYVz/Y5tOnyfr31qv6A9Hx/5sQAFz8bHws1ANtAaMB9APCCASHiux0+Y+EouMXT19kyz/Ua/Xj3xeMjp85IwzAmbFwb/4J5gfA
ArBQAILBoRHoTDsQ+2Foht4/2FPrb/wb4bEzpgGIBnqEmMb9AASDl6anhhlhZ/E0LH9QPyxMDP57v/yZoBM4wwC4CA7A3zu0kN/p
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
""",
    "m192": """
iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAMAAABlApw1AAABgFBMVEX////+///+//79//7+/v/+/v79/v79/v39/f79/f38/v38
/vz8/f36/fv8/P36/Pv3+/n19vno9+zp7fLW8d7A68nc4eq+4NTEztuzvdGa36iH1p9r0X8+0FCXt8FYwohAyFQ/vmU2yVA1xE41
wk43wkY2wkY1wkY3wkU2wUo2wUc3wUY0wUs1wUg0v080wEw2wEowv1Uyv04wu1wvvFYuvVUwvlMpumcrulwtu1gstmcrt1wmtW8o
tmIor3oor2UkrnSPnrtMm5cmpXUmlX9yhalgc50qhoYvdYgjpnYin3gil3shjn4fhX8fe38gc4AgbH9FXo4iYH8dY30dWHszTIMf
TnobTngaR3YdQnYcO3QWPXMWOHEYNXAVNW8TNHEVMXATMW8SMHASL28SLW8RLm4RLW8QLXAQLG8PLG8QK28QK20PK28PK20PKm0Q
KmwOKm0PKW0OKWwPKWoOKWsOKGsNKG0NKGsNKGoNJ2oMJ2kMJ2gLJ2gNJmkMJmgLJmjzlgOqAAAZm0lEQVR42u2bi1tSWdvGl8BG
PkUOcpCDgmI0aVn6WtnRRAOrmWZmV9gLiqAiGKAgKB6amn/9u5+198YNqDHvtbHxurwvMycavX/rOaxnLXaM3ehGN7rRjW50oxvd
6EY3+tnqZWzQHxp1DejOfdE5eufunbvh4UFd6+u6AddoaNjJv8PPk07P7L7RkH/wgtf1zBWG/9HBc20O+sdDPhvr0/3M5XeNjoy6
yOlFf8Mfuht2nfs6/sw5PEJBMP4c9739rN8/MuK3s4GBfr3u3HXU65yj4bGB3nNf1PWyXt/4uG/gJ6SRvp8s2bCALtUf9rYbMTL8
nQsDpBd4EAavmkBPhgZdvrGR8DDJ7/e5nINSUbSsdb/gG/ELF/oT+pnNT0mov8JCQOSZ0zc8Gh4J370THgmFRqDQndFhv4sgmt2i
jEf8qgUWdEJbHflC4z68cGX+qfZCI6FwGP7vhBXdCYdGwqPDPifvTRcC4FVjs1UUDzqBv//CPNNYfWzANzpyOxyevH37jlp3AUQQ
VBVnRdsOgFf79C1pBIJh29UQ6Gn5b5N70p0WAeJOKKzenloABKlZGfVNnbMfG94VEQhYrNBE+LaiO5RETQh3w3dDo9ie9L1yirdE
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
""",
    "m512": """
iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAMAAADDpiTIAAABgFBMVEX////+///+//79//7+/v/+/v79/f77/fz3+vrp9+3f7+jW
3eaz5r+0z9CI155mzX5CxVU3xEg3wkY3wkU3wkQ2xEk2wkg2wkY1wkY2wkWXq8E+umE2wUo2wUg2wUc2wUY2v0w1wUk1wUg1wUc1
v040wUozwEwywE4yv08wv1IwvlUxvlEuvVYuvVQuulstu1cuvFUtvFYruV4ouWIsulortWcrtV0ltGwntWQnrnUnr2UjrnMmqHQm
o3cipXclnHwhnHcjlX1bf54mg4EkfoEeinwegH4ed38fb384ZIkhYn0dZn0dXHwpUn8cT3gaTngeQ3YYQnQaPHIZN3AVOnEVNm8U
NG8VMm4TMm4SMXAUMG8SL24SLm4QMHARLm4RLW8QLW4QLG0QK24QKm0PLG0PK20PKm4PKm0PKmwPKW0PKWwOKmwOKWwNKW0OKWsP
KGsOKGsOKGoNKGsNKGoNKGkNJ2oNJ2kNJmgMJ2oMJ2kMJ2gMJmkMJmgMJmcMJWgLJmh24pIbAABeBUlEQVR42u2diUNT19b2Q8zJ
SALqK34eAlLHOmvVWtvavo32vU1v08ZEQQMSkkgGEgKEQVBb//VvrbX3PmefISMBmuSse4UwFOU8v72mPblcjjnmmGOOOeaYY445
5phjjjnmmGOOOeaYY4455phjjjnmmGOOOeaYY4455phjjjnmmGOOOeaYY4455phjjjnmmGOOOeaYY4455phjjjnmmGOOOeaYY445
5phjjjnmmGOOOeaYY4451k9ze12uwPjE6TNnzpyejAR9h/hRvvGJ/9HtwoULFy9evDAZCbjcXf4gL/yDvjoXjc5NTowH3S634sh0
VKaANv7xicnTZ8+ePj0R9B/up3lDE/9z1gAA6d+LBSKT585Fv0IE4Ae4HaWOaPiD/MEIyX9mYtx/+Cc9PnH2rAQA6O/v5WfikId/
17mvEIEIcKk4CByF/OBdUf7TZ86eppHqVQ7rT5Txyf9hCFxAAM724P81NCGiRCEOXHKcwNENfyE/eH/i4fA/0h85/T9nNQcAP9bb
czzB3ITiQHSSvJOTCfRbfsz9QH4Y/v3w/vynBib+5wwDAALA+GF+quJlcSCKCEQwGXRU66f+Xib/GYj+wf55WLdrfJJiAHiAs70l
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
""",
}


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
