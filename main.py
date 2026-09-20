from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
import pandas as pd
import os
import shutil
import uvicorn

app = FastAPI(
    title="Pole Dance Rojas Sport - Inventario & Activos",
    description="Sistema de control de inventario y disponibilidad para Pole Dance Rojas Sport",
    version="9.0.0"
)

EXCEL_PATH = "Control_Inventario_Pole_Dance.xlsx"
FOTOS_DIR = "static/fotos"
os.makedirs(FOTOS_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

# Diccionarios globales para mantener el registro de movimientos en memoria
historial_entradas = {}  # {sku: cantidad}
historial_ventas = {}    # {sku: cantidad}

PALABRAS_CATALOGO = ["catálogo", "catalogo", "productos", "marcancia", "mercancia", "ropa"]
PALABRAS_ACTIVOS = ["activos", "equipamiento", "activos fijos"]


def _foto_url(sku: str) -> str | None:
    """Devuelve la URL de la foto del SKU si existe un archivo guardado para él."""
    for ext in ("jpg", "jpeg", "png", "webp"):
        ruta = os.path.join(FOTOS_DIR, f"{sku}.{ext}")
        if os.path.exists(ruta):
            return f"/{ruta}"
    return None


def cargar_inventario_desde_excel():
    inventario = {}

    if os.path.exists(EXCEL_PATH):
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH)

            # 1. Leer Catálogo / Mercancía
            for sheet in excel_file.sheet_names:
                if any(kw in sheet.lower() for kw in PALABRAS_CATALOGO):
                    df_cat = pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=0)
                    col_map = {str(c).strip().lower(): c for c in df_cat.columns}

                    col_sku = next((col_map[k] for k in col_map if 'sku' in k or 'código' in k or 'codigo' in k), df_cat.columns[0])
                    col_desc = next((col_map[k] for k in col_map if 'descrip' in k or 'estilo' in k or 'nombre' in k), None)
                    col_talla = next((col_map[k] for k in col_map if 'talla' in k), None)
                    col_color = next((col_map[k] for k in col_map if 'color' in k), None)
                    col_stock = next((col_map[k] for k in col_map if 'stock inicial' in k or 'inicial' in k or 'cantidad' in k or 'stock' in k), None)
                    col_precio_venta = next((col_map[k] for k in col_map if 'precio venta' in k or 'venta' in k), None)

                    conteo_codigos = {}

                    for _, row in df_cat.iterrows():
                        codigo = str(row[col_sku]).strip() if pd.notna(row[col_sku]) else ""
                        if not codigo or codigo.lower() in ['nan', 'none', 'sku', 'sku / código', 'código', 'codigo', '']:
                            continue

                        desc = str(row[col_desc]).strip() if col_desc and pd.notna(row[col_desc]) else ""
                        talla = str(row[col_talla]).strip() if col_talla and pd.notna(row[col_talla]) else ""
                        color = str(row[col_color]).strip() if col_color and pd.notna(row[col_color]) else ""

                        partes = []
                        if desc and desc.lower() != 'nan':
                            partes.append(desc)
                        if color and color.lower() != 'nan':
                            partes.append(f"- {color}")
                        if talla and talla.lower() != 'nan':
                            partes.append(f"- Talla {talla}")
                        nombre_completo = " ".join(partes) if partes else codigo

                        try:
                            val_stock = row[col_stock] if col_stock and pd.notna(row[col_stock]) else 0
                            stock_ini = int(float(val_stock))
                        except (ValueError, TypeError):
                            stock_ini = 0

                        try:
                            precio_venta = int(float(row[col_precio_venta])) if col_precio_venta and pd.notna(row[col_precio_venta]) else 0
                        except (ValueError, TypeError):
                            precio_venta = 0

                        conteo_codigos[codigo] = conteo_codigos.get(codigo, 0) + 1
                        sku_final = codigo if conteo_codigos[codigo] == 1 else f"{codigo}-{conteo_codigos[codigo]}"

                        inventario[sku_final] = {
                            "codigo_base": codigo,
                            "nombre": nombre_completo,
                            "categoria": "productos",
                            "stock_inicial": stock_ini,
                            "precio_venta": precio_venta,
                            "entradas": historial_entradas.get(sku_final, 0),
                            "ventas": historial_ventas.get(sku_final, 0),
                        }

            # 2. Leer Activos Fijos / Equipamiento
            for sheet in excel_file.sheet_names:
                if any(kw in sheet.lower() for kw in PALABRAS_ACTIVOS):
                    df_act = pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=0)
                    col_map = {str(c).strip().lower(): c for c in df_act.columns}

                    col_sku = next((col_map[k] for k in col_map if 'código' in k or 'codigo' in k or 'sku' in k), df_act.columns[0])
                    col_nombre = next((col_map[k] for k in col_map if 'nombre del activo' in k or 'nombre' in k or 'descrip' in k), None)
                    col_cant = next((col_map[k] for k in col_map if 'cantidad' in k or 'stock' in k), None)

                    for _, row in df_act.iterrows():
                        sku = str(row[col_sku]).strip() if pd.notna(row[col_sku]) else ""
                        if sku and sku.lower() not in ['nan', 'none', 'código activo', 'codigo activo', 'sku', '']:
                            nombre_activo = str(row[col_nombre]).strip() if col_nombre and pd.notna(row[col_nombre]) else sku

                            try:
                                stock_ini = int(float(row[col_cant])) if col_cant and pd.notna(row[col_cant]) else 1
                            except (ValueError, TypeError):
                                stock_ini = 1

                            inventario[sku] = {
                                "codigo_base": sku,
                                "nombre": nombre_activo if nombre_activo.lower() != 'nan' else sku,
                                "categoria": "equipamiento",
                                "stock_inicial": stock_ini,
                                "precio_venta": 0,
                                "entradas": historial_entradas.get(sku, 0),
                                "ventas": historial_ventas.get(sku, 0),
                            }

        except Exception as e:
            print(f"⚠️ Error leyendo Excel: {e}")

    if not inventario:
        items_base = {
            "EQ-TUBO": ("Tubo Pole Dance Profesional", "equipamiento"),
            "EQ-MAT": ("Mat / Colchoneta de Caída", "equipamiento")
        }
        for sku, (nombre, cat) in items_base.items():
            inventario[sku] = {
                "codigo_base": sku,
                "nombre": nombre,
                "categoria": cat,
                "stock_inicial": 5,
                "precio_venta": 0,
                "entradas": historial_entradas.get(sku, 0),
                "ventas": historial_ventas.get(sku, 0),
            }

    return inventario


def guardar_inventario_en_excel():
    """Actualiza el archivo Excel persistente manteniendo hojas existentes."""
    if not os.path.exists(EXCEL_PATH):
        return

    try:
        excel_file = pd.ExcelFile(EXCEL_PATH)
        df_dict = {}

        for sheet in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
            col_map = {str(c).strip().lower(): c for c in df.columns}

            # Actualizar hoja de productos
            if any(kw in sheet.lower() for kw in PALABRAS_CATALOGO):
                col_sku = next((col_map[k] for k in col_map if 'sku' in k or 'código' in k or 'codigo' in k), None)
                col_entradas = next((col_map[k] for k in col_map if 'entrada' in k), None)
                col_ventas = next((col_map[k] for k in col_map if 'venta' in k or 'salida' in k), None)

                if col_sku:
                    if not col_entradas:
                        df['Entradas'] = 0
                        col_entradas = 'Entradas'
                    if not col_ventas:
                        df['Ventas'] = 0
                        col_ventas = 'Ventas'

                    for idx, row in df.iterrows():
                        sku = str(row[col_sku]).strip() if pd.notna(row[col_sku]) else ""
                        if sku in inventario_db:
                            df.at[idx, col_entradas] = inventario_db[sku]["entradas"]
                            df.at[idx, col_ventas] = inventario_db[sku]["ventas"]

            df_dict[sheet] = df

        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            for sheet_name, df_data in df_dict.items():
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        print(f"⚠️ Error al guardar cambios en Excel: {e}")


# Carga inicial de datos
inventario_db = cargar_inventario_desde_excel()


class Movimiento(BaseModel):
    sku: str = Field(..., description="Código SKU del elemento")
    cantidad: int = Field(..., description="Cantidad de unidades", gt=0)
    registrado_por: str = Field(..., description="Nombre del responsable")


def _item_publico(sku: str, item: dict) -> dict:
    stock_actual = item["stock_inicial"] + item["entradas"] - item["ventas"]
    return {
        "sku": sku,
        "nombre": item["nombre"],
        "categoria": item.get("categoria", "productos"),
        "stock_actual": stock_actual,
        "precio_venta": item.get("precio_venta", 0),
        "foto_url": _foto_url(sku),
    }


# ==================== ENDPOINTS API ====================

@app.get("/buscar")
def buscar_items(q: str = "", categoria: str = "", limit: int = 25, offset: int = 0):
    query = q.lower().strip()
    resultados = []
    
    for sku, item in inventario_db.items():
        if categoria and item.get("categoria") != categoria:
            continue
        
        texto_busqueda = f"{sku} {item['nombre']}".lower()
        if not query or query in texto_busqueda:
            resultados.append(_item_publico(sku, item))
            
    total = len(resultados)
    paginados = resultados[offset:offset + limit]
    return {"total": total, "items": paginados}


@app.get("/kpis")
def obtener_kpis(categoria: str = "productos"):
    total_skus = 0
    total_stock = 0
    sin_stock = 0
    
    for sku, item in inventario_db.items():
        if item.get("categoria") == categoria:
            stock = item["stock_inicial"] + item["entradas"] - item["ventas"]
            total_skus += 1
            total_stock += max(0, stock)
            if stock <= 0:
                sin_stock += 1
                
    return {"total_skus": total_skus, "total_stock": total_stock, "sin_stock": sin_stock}


@app.post("/movimientos/{tipo}")
def registrar_movimiento(tipo: str, mov: Movimiento):
    if mov.sku not in inventario_db:
        raise HTTPException(status_code=404, detail="El SKU no existe")
    
    if tipo == "ventas":
        inventario_db[mov.sku]["ventas"] += mov.cantidad
        historial_ventas[mov.sku] = inventario_db[mov.sku]["ventas"]
    elif tipo == "entradas":
        inventario_db[mov.sku]["entradas"] += mov.cantidad
        historial_entradas[mov.sku] = inventario_db[mov.sku]["entradas"]
    else:
        raise HTTPException(status_code=400, detail="Tipo de movimiento inválido")
    
    guardar_inventario_en_excel()
    return {"mensaje": "Movimiento registrado", "item": _item_publico(mov.sku, inventario_db[mov.sku])}


@app.post("/fotos/{sku}")
async def subir_foto(sku: str, archivo: UploadFile = File(...)):
    if sku not in inventario_db:
        raise HTTPException(status_code=404, detail="SKU no encontrado")
    
    ext = archivo.filename.split(".")[-1].lower()
    if ext not in ["jpg", "jpeg", "png", "webp"]:
        raise HTTPException(status_code=400, detail="Formato de imagen no soportado")
    
    ruta_destino = os.path.join(FOTOS_DIR, f"{sku}.{ext}")
    with open(ruta_destino, "wb") as buffer:
        shutil.copyfileobj(archivo.file, buffer)
        
    return {"mensaje": "Foto subida con éxito", "foto_url": f"/{ruta_destino}"}


@app.get("/descargar-excel")
def descargar_excel():
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(status_code=404, detail="El archivo Excel no existe")
    return FileResponse(EXCEL_PATH, filename="Inventario_Pole_Dance_Actualizado.xlsx")


@app.get("/sincronizar")
def sincronizar_excel():
    global inventario_db
    inventario_db = cargar_inventario_desde_excel()
    return {"mensaje": "Inventario re-sincronizado desde el Excel exitosamente"}


@app.get("/", response_class=HTMLResponse, tags=["Interfaz"])
def interfaz_usuario():
    return '''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="apple-mobile-web-app-title" content="Pole Sport">
    <meta name="theme-color" content="#0f0c20">
    <link rel="apple-touch-icon" href="https://cdn-icons-png.flaticon.com/512/3081/3081559.png">

    <title>Pole Dance Rojas Sport</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-gradient: linear-gradient(135deg, #0f0c20 0%, #1a102f 50%, #261239 100%);
            --card-bg: rgba(255, 255, 255, 0.96);
            --primary: #d946ef;
            --primary-dark: #c026d3;
            --accent: #8b5cf6;
            --accent-glow: rgba(217, 70, 239, 0.35);
            --text-main: #1e1b4b;
            --text-muted: #64748b;
            --success: #10b981;
            --success-bg: #ecfdf5;
            --warning: #f59e0b;
            --warning-bg: #fffbeb;
            --danger: #ef4444;
            --danger-bg: #fef2f2;
            --radius-xl: 20px;
            --radius-lg: 14px;
            --radius-md: 10px;
            --shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.3);
        }

        * { box-sizing: border-box; font-family: 'Plus Jakarta Sans', sans-serif; margin: 0; padding: 0; }

        body {
            background: var(--bg-gradient);
            background-attachment: fixed;
            color: var(--text-main);
            min-height: 100vh;
            padding: 20px 15px 40px;
        }

        .container { max-width: 1000px; margin: 0 auto; }

        header {
            text-align: center;
            padding: 25px 20px 20px;
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(12px);
            border-radius: var(--radius-xl);
            border: 1px solid rgba(255, 255, 255, 0.12);
            margin-bottom: 25px;
            box-shadow: var(--shadow);
        }

        .brand-logo {
            font-size: 2.2rem;
            font-weight: 800;
            background: linear-gradient(135deg, #ffffff 0%, #f472b6 50%, #d946ef 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -0.5px;
            text-transform: uppercase;
            margin-bottom: 6px;
        }

        .brand-subtitle {
            color: #cbd5e1;
            font-size: 0.95rem;
            font-weight: 500;
            letter-spacing: 1px;
            margin-bottom: 15px;
        }

        .header-buttons { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; }

        .btn-top {
            display: inline-flex; align-items: center; gap: 8px;
            background: rgba(255, 255, 255, 0.15); color: #fff;
            padding: 8px 16px; border-radius: 30px; text-decoration: none;
            font-size: 0.85rem; font-weight: 600; border: 1px solid rgba(255, 255, 255, 0.25);
            transition: all 0.2s ease; cursor: pointer;
        }
        .btn-top:hover { background: rgba(255, 255, 255, 0.25); transform: translateY(-2px); }
        .btn-reload { background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%); border: none; }

        .glass-card {
            background: var(--card-bg); border-radius: var(--radius-xl); padding: 24px;
            box-shadow: var(--shadow); margin-bottom: 25px; border: 1px solid rgba(255, 255, 255, 0.8);
        }

        .card-title {
            font-size: 1.15rem; font-weight: 700; color: var(--text-main);
            margin-bottom: 18px; display: flex; align-items: center; gap: 10px;
        }

        .search-box-wrapper { position: relative; }
        .search-box-wrapper input {
            width: 100%; padding: 14px 16px; border: 2px solid #e2e8f0; border-radius: var(--radius-lg);
            font-size: 1rem; font-weight: 500; color: var(--text-main); background-color: #f8fafc; outline: none;
        }
        .search-box-wrapper input:focus {
            border-color: var(--primary); background-color: #fff; box-shadow: 0 0 0 4px var(--accent-glow);
        }

        .autocomplete-results {
            position: absolute; top: 105%; left: 0; right: 0; background: white;
            border-radius: var(--radius-lg); box-shadow: 0 10px 25px rgba(0,0,0,0.2);
            border: 1px solid #cbd5e1; max-height: 260px; overflow-y: auto; z-index: 100; display: none;
        }
        .autocomplete-item {
            padding: 10px 16px; cursor: pointer; border-bottom: 1px solid #f1f5f9;
            font-size: 0.9rem; display: flex; justify-content: space-between; align-items: center; gap: 10px;
        }
        .autocomplete-item:hover { background-color: #f0fdf4; color: var(--primary-dark); }
        .autocomplete-thumb {
            width: 34px; height: 34px; border-radius: 8px; object-fit: cover; flex-shrink: 0;
            background: #f1f5f9;
        }

        .status-card { margin-top: 18px; padding: 18px 20px; border-radius: var(--radius-lg); display: none; }
        .status-card-inner { display: flex; gap: 16px; align-items: center; }
        .status-foto { width: 72px; height: 72px; border-radius: var(--radius-md); object-fit: cover; background: #e2e8f0; flex-shrink: 0; }
        .status-card h3 { font-size: 1.3rem; font-weight: 800; margin-bottom: 4px; }
        .status-card p { font-size: 0.95rem; font-weight: 500; }
        .status-available { background: var(--success-bg); color: #047857; border: 1.5px solid #a7f3d0; }
        .status-low { background: var(--warning-bg); color: #b45309; border: 1.5px solid #fde68a; }
        .status-empty { background: var(--danger-bg); color: #b91c1c; border: 1.5px solid #fca5a5; }

        .foto-upload-btn {
            margin-top: 10px; font-size: 0.78rem; font-weight: 700; color: var(--primary-dark);
            background: rgba(217, 70, 239, 0.1); border: 1px solid var(--primary);
            padding: 5px 12px; border-radius: 20px; cursor: pointer; display: inline-block;
        }

        .tab-group {
            display: flex; gap: 10px; margin-bottom: 25px; background: rgba(0, 0, 0, 0.25);
            padding: 6px; border-radius: 50px; backdrop-filter: blur(8px);
        }
        .tab-btn {
            flex: 1; padding: 12px 18px; border: none; background: transparent; color: #94a3b8;
            font-weight: 700; font-size: 0.92rem; border-radius: 40px; cursor: pointer;
        }
        .tab-btn.active {
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%);
            color: white; box-shadow: 0 4px 15px var(--accent-glow);
        }

        .kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
        @media (max-width: 640px) { .kpi-grid { grid-template-columns: 1fr; } }
        .kpi-card {
            background: var(--card-bg); padding: 18px; border-radius: var(--radius-lg);
            text-align: center; box-shadow: var(--shadow); border: 1px solid rgba(255, 255, 255, 0.8);
        }
        .kpi-card h4 {
            font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.8px;
            color: var(--text-muted); margin-bottom: 6px; font-weight: 700;
        }
        .kpi-card .number { font-size: 1.8rem; font-weight: 800; color: var(--text-main); }
        .kpi-card.alert .number { color: var(--danger); }

        .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 25px; }
        @media (max-width: 768px) { .form-grid { grid-template-columns: 1fr; } }
        .form-group { margin-bottom: 14px; }
        .form-group label {
            display: block; font-size: 0.82rem; font-weight: 700; color: #475569;
            margin-bottom: 6px; text-transform: uppercase;
        }
        .form-group input, .form-group select {
            width: 100%; padding: 11px 14px; border: 1.5px solid #cbd5e1; border-radius: var(--radius-md);
            font-size: 0.95rem; outline: none; background: #f8fafc;
        }
        .btn-action {
            width: 100%; padding: 13px; border: none; border-radius: var(--radius-md);
            font-weight: 700; font-size: 0.95rem; cursor: pointer; color: white; margin-top: 6px;
        }
        .btn-salida { background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); }
        .btn-entrada { background: linear-gradient(135deg, #10b981 0%, #059669 100%); }

        .table-header {
            display: flex; justify-content: space-between; align-items: center;
            margin-bottom: 16px; flex-wrap: wrap; gap: 12px;
        }
        .search-input {
            padding: 10px 16px; border: 1.5px solid #cbd5e1; border-radius: 30px;
            font-size: 0.9rem; width: 260px; outline: none; background: #f8fafc;
        }

        .table-wrapper { overflow-x: auto; border-radius: var(--radius-lg); border: 1px solid #e2e8f0; }
        table { width: 100%; border-collapse: collapse; background: white; text-align: left; }
        th {
            background: #f8fafc; color: #475569; font-size: 0.8rem; text-transform: uppercase;
            font-weight: 700; padding: 14px 16px; border-bottom: 1px solid #e2e8f0;
        }
        td { padding: 10px 16px; border-bottom: 1px solid #f1f5f9; font-size: 0.92rem; color: #334155; vertical-align: middle; }

        .row-thumb { width: 42px; height: 42px; border-radius: 8px; object-fit: cover; background: #f1f5f9; display: block; }

        .row-action-btn {
            background: rgba(217, 70, 239, 0.1); color: var(--primary-dark); border: 1px solid var(--primary);
            padding: 4px 10px; border-radius: 12px; font-size: 0.78rem; font-weight: 700; cursor: pointer;
        }

        .badge { padding: 4px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 700; display: inline-block; }
        .badge-success { background: var(--success-bg); color: #047857; }
        .badge-warning { background: var(--warning-bg); color: #b45309; }
        .badge-danger { background: var(--danger-bg); color: #b91c1c; }

        .btn-cargar-mas {
            width: 100%; margin-top: 14px; padding: 12px; border-radius: var(--radius-md);
            border: 1.5px dashed #cbd5e1; background: #f8fafc; color: var(--text-muted);
            font-weight: 700; cursor: pointer;
        }
        .empty-hint { text-align: center; color: #94a3b8; padding: 30px 10px; font-size: 0.9rem; }

        #toast {
            position: fixed; bottom: 25px; right: 25px; padding: 14px 22px; border-radius: var(--radius-lg);
            color: white; font-weight: 600; font-size: 0.95rem; display: none;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.25); z-index: 1000;
        }
        .toast-success { background: #10b981; }
        .toast-error { background: #ef4444; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="brand-logo">✨ Pole Dance Rojas Sport</div>
            <div class="brand-subtitle">SISTEMA INTEGRAL DE INVENTARIO Y DISPONIBILIDAD</div>
            <div class="header-buttons">
                <a href="/descargar-excel" class="btn-top">📥 Descargar Excel</a>
                <button onclick="recargarDesdeExcel()" class="btn-top btn-reload">🔄 Sincronizar Cambios de Excel</button>
            </div>
        </header>

        <div class="glass-card lookup-box">
            <div class="card-title">🔍 Consulta Rápida</div>
            <div class="search-box-wrapper">
                <input type="text" id="lookup-input" placeholder="Escribe 'Short', 'Top', 'Velvet', 'Barra' o un SKU..." oninput="onLookupInput(this.value)" autocomplete="off">
                <div id="autocomplete-list" class="autocomplete-results"></div>
            </div>

            <div id="status-display" class="status-card">
                <div class="status-card-inner">
                    <img id="status-foto" class="status-foto" src="" alt="" style="display:none;">
                    <div style="flex:1;">
                        <h3 id="status-title">---</h3>
                        <p id="status-desc">---</p>
                        <label class="foto-upload-btn">
                            📷 Subir / cambiar foto
                            <input type="file" accept="image/*" capture="environment" id="foto-input" style="display:none;" onchange="subirFoto(this.files[0])">
                        </label>
                    </div>
                </div>
            </div>
        </div>

        <div class="tab-group">
            <button class="tab-btn active" id="btn-tab-productos" onclick="cambiarPestana('productos')">👗 Ropa & Productos</button>
            <button class="tab-btn" id="btn-tab-equipamiento" onclick="cambiarPestana('equipamiento')">🪑 Equipamiento & Activos</button>
        </div>

        <div class="kpi-grid">
            <div class="kpi-card"><h4>Catálogo Activo</h4><div class="number" id="kpi-total-skus">—</div></div>
            <div class="kpi-card"><h4>Stock Físico Total</h4><div class="number" id="kpi-total-stock">—</div></div>
            <div class="kpi-card alert"><h4>Ítems Agotados</h4><div class="number" id="kpi-sin-stock">—</div></div>
        </div>

        <div class="form-grid">
            <div class="glass-card">
                <div class="card-title" id="form-salida-title">🛍️ Registrar Venta / Salida</div>
                <div class="form-group">
                    <label>Buscar Ítem</label>
                    <input type="text" id="v_search" placeholder="Escribe para buscar (mín. 2 letras)..." oninput="buscarParaSelect('v_sku', this.value)" style="margin-bottom: 6px;">
                    <select id="v_sku" size="4" style="height: 110px;"></select>
                </div>
                <div class="form-group"><label>Cantidad a Descontar</label><input type="number" id="v_cant" value="1" min="1"></div>
                <div class="form-group"><label>Registrado por</label><input type="text" id="v_usuario" placeholder="Ej: Profesora María"></div>
                <button class="btn-action btn-salida" id="btn-salida-action" onclick="procesarMovimiento('ventas')">Descontar Unidad</button>
            </div>

            <div class="glass-card">
                <div class="card-title" id="form-entrada-title">📦 Registrar Entrada / Compra</div>
                <div class="form-group">
                    <label>Buscar Ítem</label>
                    <input type="text" id="e_search" placeholder="Escribe para buscar (mín. 2 letras)..." oninput="buscarParaSelect('e_sku', this.value)" style="margin-bottom: 6px;">
                    <select id="e_sku" size="4" style="height: 110px;"></select>
                </div>
                <div class="form-group"><label>Cantidad Ingresada</label><input type="number" id="e_cant" value="1" min="1"></div>
                <div class="form-group"><label>Registrado por</label><input type="text" id="e_usuario" placeholder="Ej: Admin"></div>
                <button class="btn-action btn-entrada" onclick="procesarMovimiento('entradas')">Ingresar al Stock</button>
            </div>
        </div>

        <div class="glass-card">
            <div class="table-header">
                <div class="card-title" id="tabla-titulo" style="margin-bottom:0;">📋 Listado de Productos</div>
                <input type="text" id="search" class="search-input" placeholder="🔍 Escribe una palabra..." onkeyup="onTablaSearch()">
            </div>
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr><th></th><th>SKU</th><th>Descripción / Producto</th><th>Estado</th><th>Stock</th><th>Acción</th></tr>
                    </thead>
                    <tbody id="tabla-body">
                        <tr><td colspan="6" class="empty-hint">Escribe algo arriba para buscar, o espera a que cargue la primera página...</td></tr>
                    </tbody>
                </table>
            </div>
            <button class="btn-cargar-mas" id="btn-cargar-mas" onclick="cargarMasTabla()" style="display:none;">Cargar más</button>
        </div>
    </div>

    <div id="toast"></div>

    <script>
        let categoriaActual = 'productos';
        let skuSeleccionadoParaFoto = null;
        let paginaTabla = 0;
        const PAGE_SIZE = 25;
        let debounceTimer = null;

        function mostrarToast(mensaje, esExito) {
            const toast = document.getElementById('toast');
            toast.innerText = mensaje;
            toast.className = esExito ? 'toast-success' : 'toast-error';
            toast.style.display = 'block';
            setTimeout(() => { toast.style.display = 'none'; }, 3500);
        }

        function debounce(fn, ms) {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(fn, ms);
        }

        async function cargarKpis() {
            try {
                const res = await fetch('/kpis?categoria=' + categoriaActual);
                const data = await res.json();
                document.getElementById('kpi-total-skus').innerText = data.total_skus;
                document.getElementById('kpi-total-stock').innerText = data.total_stock;
                document.getElementById('kpi-sin-stock').innerText = data.sin_stock;
            } catch(e) {}
        }

        async function buscar(query, categoria, limit, offset) {
            const params = new URLSearchParams({ q: query || '', categoria: categoria || '', limit, offset });
            const res = await fetch('/buscar?' + params.toString());
            if (!res.ok) throw new Error('error de búsqueda');
            return await res.json();
        }

        function onLookupInput(valor) {
            const term = valor.trim();
            const listDiv = document.getElementById('autocomplete-list');
            if (term.length < 2) { listDiv.style.display = 'none'; return; }
            debounce(async () => {
                try {
                    const data = await buscar(term, '', 8, 0);
                    if (data.items.length === 0) {
                        listDiv.innerHTML = '<div class="autocomplete-item" style="color:#94a3b8;">Sin resultados</div>';
                    } else {
                        listDiv.innerHTML = data.items.map(item => {
                            const catTag = item.categoria === 'productos' ? '👗' : '🪑';
                            const stockTag = item.stock_actual > 0 ? `<b>${item.stock_actual} ud.</b>` : '<span style="color:#ef4444;font-weight:bold;">Agotado</span>';
                            const thumb = item.foto_url ? `<img class="autocomplete-thumb" src="${item.foto_url}">` : '';
                            return `<div class="autocomplete-item" onclick='seleccionarDatoConsulta(${JSON.stringify(item.sku)})'>
                                        <span style="display:flex;align-items:center;gap:8px;">${thumb}${catTag} <b>[${item.sku}]</b> ${item.nombre}</span>
                                        <span>${stockTag}</span>
                                    </div>`;
                        }).join('');
                    }
                    listDiv.style.display = 'block';
                } catch(e) { mostrarToast("Error de conexión con el servidor", false); }
            }, 250);
        }

        async function seleccionarDatoConsulta(sku) {
            document.getElementById('autocomplete-list').style.display = 'none';
            try {
                const data = await buscar(sku, '', 1, 0);
                const item = data.items.find(i => i.sku === sku) || data.items[0];
                if (!item) return;
                document.getElementById('lookup-input').value = item.nombre;
                if (item.categoria !== categoriaActual) cambiarPestana(item.categoria);
                mostrarConsulta(item);
                document.getElementById('v_search').value = item.sku;
                document.getElementById('e_search').value = item.sku;
                buscarParaSelect('v_sku', item.sku);
                buscarParaSelect('e_sku', item.sku);
            } catch(e) { mostrarToast("Error de conexión con el servidor", false); }
        }

        function mostrarConsulta(item) {
            const display = document.getElementById('status-display');
            const title = document.getElementById('status-title');
            const desc = document.getElementById('status-desc');
            const foto = document.getElementById('status-foto');

            skuSeleccionadoParaFoto = item.sku;
            if (item.foto_url) { foto.src = item.foto_url; foto.style.display = 'block'; }
            else { foto.style.display = 'none'; }

            display.style.display = 'block';
            display.className = 'status-card ';
            const stock = item.stock_actual;
            if (stock > 2) {
                display.classList.add('status-available');
                title.innerHTML = '🟢 SÍ HAY DISPONIBILIDAD';
                desc.innerHTML = `Tenemos <b>${stock} unidades</b> disponibles de <i>${item.nombre}</i>.`;
            } else if (stock > 0) {
                display.classList.add('status-low');
                title.innerHTML = '⚠️ ÚLTIMAS UNIDADES';
                desc.innerHTML = `¡Atención! Quedan solo <b>${stock} unidad(es)</b> disponibles de <i>${item.nombre}</i>.`;
            } else {
                display.classList.add('status-empty');
                title.innerHTML = '🔴 AGOTADO';
                desc.innerHTML = `Actualmente hay <b>0 unidades</b> de <i>${item.nombre}</i>.`;
            }
        }

        async function subirFoto(file) {
            if (!file || !skuSeleccionadoParaFoto) return;
            const formData = new FormData();
            formData.append('archivo', file);
            try {
                const res = await fetch('/fotos/' + encodeURIComponent(skuSeleccionadoParaFoto), { method: 'POST', body: formData });
                const data = await res.json();
                if (res.ok) {
                    mostrarToast('📷 Foto guardada', true);
                    document.getElementById('status-foto').src = data.foto_url + '?t=' + Date.now();
                    document.getElementById('status-foto').style.display = 'block';
                    refrescarTabla();
                } else {
                    mostrarToast(data.detail || 'Error al subir foto', false);
                }
            } catch(e) { mostrarToast('Error de conexión', false); }
        }

        function buscarParaSelect(selectId, query) {
            if (query.trim().length < 2) return;
            debounce(async () => {
                try {
                    const data = await buscar(query, categoriaActual, 10, 0);
                    const select = document.getElementById(selectId);
                    select.innerHTML = '';
                    data.items.forEach(item => {
                        const opt = document.createElement('option');
                        opt.value = item.sku;
                        opt.innerText = `[${item.sku}] ${item.nombre} (Stock: ${item.stock_actual})`;
                        select.appendChild(opt);
                    });
                    if (data.items.length > 0) select.selectedIndex = 0;
                } catch(e) {}
            }, 250);
        }

        async function procesarMovimiento(tipo) {
            const prefix = tipo === 'ventas' ? 'v_' : 'e_';
            const select = document.getElementById(prefix + 'sku');
            const sku = select.value;
            const cant = parseInt(document.getElementById(prefix + 'cant').value);
            const resp = document.getElementById(prefix + 'usuario').value.trim();

            if (!sku) { mostrarToast('Selecciona un ítem de la lista', false); return; }
            if (!cant || cant <= 0) { mostrarToast('Ingresa una cantidad válida', false); return; }
            if (!resp) { mostrarToast('Ingresa el nombre del responsable', false); return; }

            try {
                const res = await fetch('/movimientos/' + tipo, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sku: sku, cantidad: cant, registrado_por: resp })
                });

                const data = await res.json();
                if (res.ok) {
                    mostrarToast('✅ Registrado con éxito', true);
                    cargarKpis();
                    refrescarTabla();
                    if (skuSeleccionadoParaFoto === sku) mostrarConsulta(data.item);
                } else {
                    mostrarToast(data.detail || 'Error al procesar', false);
                }
            } catch(e) { mostrarToast('Error de servidor', false); }
        }

        function cambiarPestana(cat) {
            categoriaActual = cat;
            document.getElementById('btn-tab-productos').classList.toggle('active', cat === 'productos');
            document.getElementById('btn-tab-equipamiento').classList.toggle('active', cat === 'equipamiento');
            document.getElementById('tabla-titulo').innerText = cat === 'productos' ? '📋 Listado de Productos' : '🪑 Listado de Equipamiento';
            
            document.getElementById('v_search').value = '';
            document.getElementById('e_search').value = '';
            document.getElementById('v_sku').innerHTML = '';
            document.getElementById('e_sku').innerHTML = '';

            cargarKpis();
            refrescarTabla();
        }

        function onTablaSearch() {
            debounce(() => refrescarTabla(), 300);
        }

        function refrescarTabla() {
            paginaTabla = 0;
            renderizarTabla(true);
        }

        function cargarMasTabla() {
            paginaTabla++;
            renderizarTabla(false);
        }

        async function renderizarTabla(limpiar) {
            const query = document.getElementById('search').value;
            const tbody = document.getElementById('tabla-body');
            const btnMas = document.getElementById('btn-cargar-mas');

            try {
                const data = await buscar(query, categoriaActual, PAGE_SIZE, paginaTabla * PAGE_SIZE);
                if (limpiar) tbody.innerHTML = '';

                if (data.items.length === 0 && limpiar) {
                    tbody.innerHTML = '<tr><td colspan="6" class="empty-hint">No se encontraron resultados.</td></tr>';
                    btnMas.style.display = 'none';
                    return;
                }

                data.items.forEach(item => {
                    let badgeClass = 'badge-success';
                    let estadoTexto = 'Disponible';
                    if (item.stock_actual <= 0) { badgeClass = 'badge-danger'; estadoTexto = 'Agotado'; }
                    else if (item.stock_actual <= 2) { badgeClass = 'badge-warning'; estadoTexto = 'Bajo Stock'; }

                    const thumb = item.foto_url ? `<img class="row-thumb" src="${item.foto_url}">` : '<div class="row-thumb" style="display:flex;align-items:center;justify-content:center;font-size:1.2rem;">🖼️</div>';

                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${thumb}</td>
                        <td><b>${item.sku}</b></td>
                        <td>${item.nombre}</td>
                        <td><span class="badge ${badgeClass}">${estadoTexto}</span></td>
                        <td><b>${item.stock_actual}</b></td>
                        <td><button class="row-action-btn" onclick='seleccionarDatoConsulta(${JSON.stringify(item.sku)})'>Ver / Editar</button></td>
                    `;
                    tbody.appendChild(row);
                });

                btnMas.style.display = (data.total > (paginaTabla + 1) * PAGE_SIZE) ? 'block' : 'none';

            } catch(e) { mostrarToast('Error cargando la lista', false); }
        }

        async function recargarDesdeExcel() {
            try {
                const res = await fetch('/sincronizar');
                if (res.ok) {
                    mostrarToast('🔄 Excel sincronizado con éxito', true);
                    cargarKpis();
                    refrescarTabla();
                }
            } catch(e) { mostrarToast('Error al re-sincronizar', false); }
        }

        window.onload = () => {
            cargarKpis();
            refrescarTabla();
        };
    </script>
</body>
</html>'''


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)