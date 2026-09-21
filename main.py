import os
import shutil
from datetime import datetime

import pandas as pd
import uvicorn
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

app = FastAPI(
    title="Pole Dance Rojas Sport - Inventario, Clases & Finanzas",
    description="Sistema integral de inventario, ventas de clases/paquetes y control financiero para Pole Dance Rojas Sport",
    version="10.0.0"
)

EXCEL_PATH = "Control_Inventario_Pole_Dance.xlsx"
FOTOS_DIR = "static/fotos"
os.makedirs(FOTOS_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

historial_entradas = {}  # {sku: cantidad}
historial_ventas = {}    # {sku: cantidad}

PALABRAS_CATALOGO = ["catálogo", "catalogo", "productos", "marcancia", "mercancia", "ropa"]
PALABRAS_ACTIVOS = ["activos", "equipamiento", "activos fijos"]
HOJAS_RESERVADAS = ("Clases_Ventas", "Gastos", "Config")


def _foto_url(sku: str) -> str | None:
    """Devuelve la URL de la foto del SKU si existe un archivo guardado para él."""
    for ext in ("jpg", "jpeg", "png", "webp"):
        ruta = os.path.join(FOTOS_DIR, f"{sku}.{ext}")
        if os.path.exists(ruta):
            return f"/{ruta}"
    return None


def _fecha_ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ==================== CARGA DESDE EXCEL ====================

def cargar_inventario_desde_excel():
    inventario = {}

    if os.path.exists(EXCEL_PATH):
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH)

            # 1. Leer Catálogo / Mercancía
            for sheet in excel_file.sheet_names:
                if sheet in HOJAS_RESERVADAS:
                    continue
                if any(kw in sheet.lower() for kw in PALABRAS_CATALOGO):
                    df_cat = pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=0)
                    col_map = {str(c).strip().lower(): c for c in df_cat.columns}

                    col_sku = next((col_map[k] for k in col_map if 'sku' in k or 'código' in k or 'codigo' in k), df_cat.columns[0])
                    col_desc = next((col_map[k] for k in col_map if 'descrip' in k or 'estilo' in k or 'nombre' in k), None)
                    col_talla = next((col_map[k] for k in col_map if 'talla' in k), None)
                    col_color = next((col_map[k] for k in col_map if 'color' in k), None)
                    col_stock = next((col_map[k] for k in col_map if 'stock inicial' in k or 'inicial' in k or 'cantidad' in k or 'stock' in k), None)
                    col_precio_venta = next((col_map[k] for k in col_map if 'precio venta' in k or 'venta' in k), None)
                    col_costo = next((col_map[k] for k in col_map if 'costo' in k or 'compra' in k), None)

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

                        try:
                            costo_unitario = float(row[col_costo]) if col_costo and pd.notna(row[col_costo]) else 0.0
                        except (ValueError, TypeError):
                            costo_unitario = 0.0

                        conteo_codigos[codigo] = conteo_codigos.get(codigo, 0) + 1
                        sku_final = codigo if conteo_codigos[codigo] == 1 else f"{codigo}-{conteo_codigos[codigo]}"

                        inventario[sku_final] = {
                            "codigo_base": codigo,
                            "nombre": nombre_completo,
                            "categoria": "productos",
                            "stock_inicial": stock_ini,
                            "precio_venta": precio_venta,
                            "costo_unitario_ref": costo_unitario,
                            "entradas": historial_entradas.get(sku_final, 0),
                            "ventas": historial_ventas.get(sku_final, 0),
                        }

            # 2. Leer Activos Fijos / Equipamiento
            for sheet in excel_file.sheet_names:
                if sheet in HOJAS_RESERVADAS:
                    continue
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
                                "costo_unitario_ref": 0.0,
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
                "costo_unitario_ref": 0.0,
                "entradas": historial_entradas.get(sku, 0),
                "ventas": historial_ventas.get(sku, 0),
            }

    return inventario


def cargar_clases_desde_excel():
    lista = []
    if os.path.exists(EXCEL_PATH):
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH)
            if "Clases_Ventas" in excel_file.sheet_names:
                df = pd.read_excel(EXCEL_PATH, sheet_name="Clases_Ventas")
                for _, row in df.iterrows():
                    if pd.notna(row.get("id")):
                        lista.append({
                            "id": int(row["id"]),
                            "fecha": str(row.get("fecha", "")),
                            "tipo_plan": str(row.get("tipo_plan", "")),
                            "nombre_cliente": str(row.get("nombre_cliente", "")),
                            "contacto": str(row.get("contacto", "")) if pd.notna(row.get("contacto")) else "",
                            "precio": float(row.get("precio", 0)) if pd.notna(row.get("precio")) else 0.0,
                            "metodo_pago": str(row.get("metodo_pago", "")) if pd.notna(row.get("metodo_pago")) else "",
                            "registrado_por": str(row.get("registrado_por", "")),
                            "notas": str(row.get("notas", "")) if pd.notna(row.get("notas")) else "",
                        })
        except Exception as e:
            print(f"⚠️ Error leyendo hoja de clases: {e}")
    return lista


def cargar_gastos_desde_excel():
    lista = []
    if os.path.exists(EXCEL_PATH):
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH)
            if "Gastos" in excel_file.sheet_names:
                df = pd.read_excel(EXCEL_PATH, sheet_name="Gastos")
                for _, row in df.iterrows():
                    if pd.notna(row.get("id")):
                        lista.append({
                            "id": int(row["id"]),
                            "fecha": str(row.get("fecha", "")),
                            "concepto": str(row.get("concepto", "")),
                            "categoria": str(row.get("categoria", "operativo")),
                            "monto": float(row.get("monto", 0)) if pd.notna(row.get("monto")) else 0.0,
                            "registrado_por": str(row.get("registrado_por", "")),
                            "notas": str(row.get("notas", "")) if pd.notna(row.get("notas")) else "",
                        })
        except Exception as e:
            print(f"⚠️ Error leyendo hoja de gastos: {e}")
    return lista


def cargar_config_desde_excel():
    config = {"inversion_inicial": 0.0}
    if os.path.exists(EXCEL_PATH):
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH)
            if "Config" in excel_file.sheet_names:
                df = pd.read_excel(EXCEL_PATH, sheet_name="Config")
                for _, row in df.iterrows():
                    if str(row.get("clave", "")).strip() == "inversion_inicial" and pd.notna(row.get("valor")):
                        config["inversion_inicial"] = float(row["valor"])
        except Exception as e:
            print(f"⚠️ Error leyendo config: {e}")
    return config


# ==================== GUARDADO EN EXCEL ====================

def guardar_todo_en_excel():
    """Actualiza el archivo Excel persistente: inventario, clases, gastos y config."""
    try:
        df_dict = {}

        if os.path.exists(EXCEL_PATH):
            excel_file = pd.ExcelFile(EXCEL_PATH)
            for sheet in excel_file.sheet_names:
                if sheet in HOJAS_RESERVADAS:
                    continue  # se reconstruyen más abajo desde memoria

                df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
                col_map = {str(c).strip().lower(): c for c in df.columns}

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

                elif any(kw in sheet.lower() for kw in PALABRAS_ACTIVOS):
                    col_sku = next((col_map[k] for k in col_map if 'código' in k or 'codigo' in k or 'sku' in k), None)
                    col_entradas = next((col_map[k] for k in col_map if 'entrada' in k), None)
                    col_salidas = next((col_map[k] for k in col_map if 'salida' in k or 'venta' in k), None)

                    if col_sku:
                        if not col_entradas:
                            df['Entradas'] = 0
                            col_entradas = 'Entradas'
                        if not col_salidas:
                            df['Salidas'] = 0
                            col_salidas = 'Salidas'

                        for idx, row in df.iterrows():
                            sku = str(row[col_sku]).strip() if pd.notna(row[col_sku]) else ""
                            if sku in inventario_db:
                                df.at[idx, col_entradas] = inventario_db[sku]["entradas"]
                                df.at[idx, col_salidas] = inventario_db[sku]["ventas"]

                df_dict[sheet] = df

        # Reconstruir hojas de Clases, Gastos y Config desde memoria
        df_dict["Clases_Ventas"] = pd.DataFrame(clases_db) if clases_db else pd.DataFrame(
            columns=["id", "fecha", "tipo_plan", "nombre_cliente", "contacto", "precio", "metodo_pago", "registrado_por", "notas"])
        df_dict["Gastos"] = pd.DataFrame(gastos_db) if gastos_db else pd.DataFrame(
            columns=["id", "fecha", "concepto", "categoria", "monto", "registrado_por", "notas"])
        df_dict["Config"] = pd.DataFrame([{"clave": "inversion_inicial", "valor": config_db.get("inversion_inicial", 0.0)}])

        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            for sheet_name, df_data in df_dict.items():
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        print(f"⚠️ Error al guardar cambios en Excel: {e}")


# Carga inicial de datos
inventario_db = cargar_inventario_desde_excel()
clases_db = cargar_clases_desde_excel()
gastos_db = cargar_gastos_desde_excel()
config_db = cargar_config_desde_excel()


# ==================== MODELOS ====================

class Movimiento(BaseModel):
    sku: str = Field(..., description="Código SKU del elemento")
    cantidad: int = Field(..., description="Cantidad de unidades", gt=0)
    registrado_por: str = Field(..., description="Nombre del responsable")
    costo_unitario: float = Field(0.0, description="Solo para entradas: costo de compra por unidad (opcional)")


class VentaClase(BaseModel):
    tipo_plan: str = Field(..., description="Ej: Clase Suelta, Paquete 4 Clases, Mensualidad")
    nombre_cliente: str = Field(..., description="Nombre de la alumna/o")
    contacto: str = ""
    precio: float = Field(..., ge=0)
    metodo_pago: str = ""
    registrado_por: str = Field(..., description="Quién registra la venta")
    notas: str = ""


class Gasto(BaseModel):
    concepto: str
    categoria: str = "operativo"
    monto: float = Field(..., gt=0)
    registrado_por: str
    notas: str = ""


class InversionInput(BaseModel):
    monto: float = Field(..., ge=0)


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


# ==================== ENDPOINTS: INVENTARIO ====================

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

    item = inventario_db[mov.sku]

    if tipo == "ventas":
        item["ventas"] += mov.cantidad
        historial_ventas[mov.sku] = item["ventas"]
    elif tipo == "entradas":
        item["entradas"] += mov.cantidad
        historial_entradas[mov.sku] = item["entradas"]

        # Si se indica costo unitario, se registra automáticamente como gasto de compra de mercancía
        if mov.costo_unitario and mov.costo_unitario > 0:
            nuevo_id = (max([g["id"] for g in gastos_db], default=0)) + 1
            gastos_db.append({
                "id": nuevo_id,
                "fecha": _fecha_ahora(),
                "concepto": f"Compra de mercancía: {item['nombre']} x{mov.cantidad}",
                "categoria": "compra_mercancia",
                "monto": round(mov.costo_unitario * mov.cantidad, 2),
                "registrado_por": mov.registrado_por,
                "notas": f"SKU {mov.sku}",
            })
    else:
        raise HTTPException(status_code=400, detail="Tipo de movimiento inválido")

    guardar_todo_en_excel()
    return {"mensaje": "Movimiento registrado", "item": _item_publico(mov.sku, item)}


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
    global inventario_db, clases_db, gastos_db, config_db
    inventario_db = cargar_inventario_desde_excel()
    clases_db = cargar_clases_desde_excel()
    gastos_db = cargar_gastos_desde_excel()
    config_db = cargar_config_desde_excel()
    return {"mensaje": "Datos re-sincronizados desde el Excel exitosamente"}


# ==================== ENDPOINTS: CLASES & PAQUETES ====================

@app.post("/clases")
def registrar_clase(venta: VentaClase):
    nuevo_id = (max([c["id"] for c in clases_db], default=0)) + 1
    registro = {
        "id": nuevo_id,
        "fecha": _fecha_ahora(),
        **venta.dict()
    }
    clases_db.append(registro)
    guardar_todo_en_excel()
    return {"mensaje": "Venta de clase/paquete registrada", "registro": registro}


@app.get("/clases")
def listar_clases(q: str = "", limit: int = 50, offset: int = 0):
    query = q.lower().strip()
    resultados = [
        c for c in clases_db
        if not query or query in f"{c['tipo_plan']} {c['nombre_cliente']} {c['registrado_por']}".lower()
    ]
    resultados = sorted(resultados, key=lambda x: x["id"], reverse=True)
    total = len(resultados)
    return {"total": total, "items": resultados[offset:offset + limit]}


@app.get("/clases/kpis")
def kpis_clases():
    total_ventas = len(clases_db)
    ingresos = sum(c["precio"] for c in clases_db)
    promedio = ingresos / total_ventas if total_ventas else 0
    return {"total_ventas": total_ventas, "ingresos_totales": ingresos, "ticket_promedio": round(promedio, 2)}


@app.delete("/clases/{id_venta}")
def eliminar_clase(id_venta: int):
    global clases_db
    antes = len(clases_db)
    clases_db = [c for c in clases_db if c["id"] != id_venta]
    if len(clases_db) == antes:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    guardar_todo_en_excel()
    return {"mensaje": "Registro eliminado"}


# ==================== ENDPOINTS: FINANZAS (GASTOS, COMPRAS, INVERSIÓN) ====================

@app.post("/gastos")
def registrar_gasto(gasto: Gasto):
    nuevo_id = (max([g["id"] for g in gastos_db], default=0)) + 1
    registro = {
        "id": nuevo_id,
        "fecha": _fecha_ahora(),
        **gasto.dict()
    }
    gastos_db.append(registro)
    guardar_todo_en_excel()
    return {"mensaje": "Gasto registrado", "registro": registro}


@app.get("/gastos")
def listar_gastos(q: str = "", categoria: str = "", limit: int = 50, offset: int = 0):
    query = q.lower().strip()
    resultados = [
        g for g in gastos_db
        if (not categoria or g["categoria"] == categoria)
        and (not query or query in f"{g['concepto']} {g['registrado_por']}".lower())
    ]
    resultados = sorted(resultados, key=lambda x: x["id"], reverse=True)
    total = len(resultados)
    return {"total": total, "items": resultados[offset:offset + limit]}


@app.delete("/gastos/{id_gasto}")
def eliminar_gasto(id_gasto: int):
    global gastos_db
    antes = len(gastos_db)
    gastos_db = [g for g in gastos_db if g["id"] != id_gasto]
    if len(gastos_db) == antes:
        raise HTTPException(status_code=404, detail="Registro no encontrado")
    guardar_todo_en_excel()
    return {"mensaje": "Gasto eliminado"}


@app.post("/finanzas/inversion")
def actualizar_inversion(data: InversionInput):
    config_db["inversion_inicial"] = data.monto
    guardar_todo_en_excel()
    return {"mensaje": "Inversión inicial actualizada", "inversion_inicial": data.monto}


@app.get("/finanzas/resumen")
def resumen_finanzas():
    ingresos_productos = 0.0
    for sku, item in inventario_db.items():
        if item.get("categoria") == "productos":
            ingresos_productos += item["ventas"] * item.get("precio_venta", 0)

    ingresos_clases = sum(c["precio"] for c in clases_db)
    total_gastos = sum(g["monto"] for g in gastos_db)
    inversion = config_db.get("inversion_inicial", 0.0)
    ingresos_totales = ingresos_productos + ingresos_clases
    ganancia_neta = ingresos_totales - total_gastos

    return {
        "inversion_inicial": inversion,
        "ingresos_productos": ingresos_productos,
        "ingresos_clases": ingresos_clases,
        "ingresos_totales": ingresos_totales,
        "total_gastos": total_gastos,
        "ganancia_neta": ganancia_neta,
        "retorno_inversion_pct": round((ganancia_neta / inversion * 100), 1) if inversion > 0 else None,
    }


# ==================== INTERFAZ ====================

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
    <link rel="apple-touch-icon" href="/static/logo.png">
    <link rel="icon" type="image/png" href="/static/logo.png">

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

        .brand-image {
            max-height: 110px;
            width: auto;
            display: block;
            margin: 0 auto 10px;
            filter: drop-shadow(0 4px 12px rgba(0,0,0,0.35));
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

        .nav-main {
            display: flex; gap: 10px; margin-bottom: 20px; background: rgba(0, 0, 0, 0.3);
            padding: 6px; border-radius: 50px; backdrop-filter: blur(8px);
        }
        .nav-btn {
            flex: 1; padding: 13px 16px; border: none; background: transparent; color: #cbd5e1;
            font-weight: 800; font-size: 0.9rem; border-radius: 40px; cursor: pointer;
        }
        .nav-btn.active {
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%);
            color: white; box-shadow: 0 4px 15px var(--accent-glow);
        }

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
            border: 1px solid #cbd5e1; max-height: 280px; overflow-y: auto; z-index: 100; display: none;
        }
        .autocomplete-item {
            padding: 10px 16px; cursor: pointer; border-bottom: 1px solid #f1f5f9;
            font-size: 0.9rem; display: flex; justify-content: space-between; align-items: center; gap: 12px;
        }
        .autocomplete-item:hover { background-color: #f0fdf4; color: var(--primary-dark); }
        .autocomplete-thumb {
            width: 44px; height: 44px; border-radius: 8px; object-fit: cover; flex-shrink: 0;
            background: #e2e8f0; border: 1px solid #cbd5e1; display: block;
        }

        .status-card { margin-top: 18px; padding: 18px 20px; border-radius: var(--radius-lg); display: none; }
        .status-card-inner { display: flex; gap: 16px; align-items: center; }
        .status-foto { width: 80px; height: 80px; border-radius: var(--radius-md); object-fit: cover; background: #e2e8f0; flex-shrink: 0; border: 1px solid rgba(0,0,0,0.1); }
        .status-card h3 { font-size: 1.3rem; font-weight: 800; margin-bottom: 4px; }
        .status-card p { font-size: 0.95rem; font-weight: 500; }
        .status-available { background: var(--success-bg); color: #047857; border: 1.5px solid #a7f3d0; }
        .status-low { background: var(--warning-bg); color: #b45309; border: 1.5px solid #fde68a; }
        .status-empty { background: var(--danger-bg); color: #b91c1c; border: 1.5px solid #fca5a5; }

        .foto-upload-btn {
            margin-top: 10px; font-size: 0.8rem; font-weight: 700; color: var(--primary-dark);
            background: rgba(217, 70, 239, 0.1); border: 1px solid var(--primary);
            padding: 6px 14px; border-radius: 20px; cursor: pointer; display: inline-block; text-align: center;
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
        .kpi-card .number { font-size: 1.6rem; font-weight: 800; color: var(--text-main); }
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

        .picker-wrapper { position: relative; }
        .picker-list {
            position: absolute; top: 100%; left: 0; right: 0; background: white; border-radius: var(--radius-md);
            box-shadow: 0 10px 25px rgba(0,0,0,0.2); border: 1px solid #cbd5e1; max-height: 220px; overflow-y: auto;
            z-index: 60; display: none;
        }
        .picker-selected { margin-top: 8px; font-size: 0.85rem; font-weight: 600; color: var(--text-muted); }

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
            font-weight: 700; padding: 14px 16px; border-bottom: 1px solid #e2e8f0; white-space: nowrap;
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

        .modal-overlay {
            display: none; position: fixed; inset: 0; background: rgba(15, 12, 32, 0.75); backdrop-filter: blur(4px);
            z-index: 2000; align-items: center; justify-content: center; padding: 20px;
        }
        .modal-box {
            background: white; border-radius: var(--radius-xl); padding: 26px; max-width: 420px; width: 100%;
            box-shadow: 0 20px 50px rgba(0,0,0,0.4); position: relative; max-height: 90vh; overflow-y: auto;
        }
        .modal-close {
            position: absolute; top: 14px; right: 16px; background: #f1f5f9; border: none; border-radius: 50%;
            width: 32px; height: 32px; font-size: 1.1rem; cursor: pointer; color: #64748b;
        }
        .modal-foto { width: 100%; height: 220px; object-fit: cover; border-radius: var(--radius-lg); margin-bottom: 16px; background: #f1f5f9; }
        .modal-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #f1f5f9; font-size: 0.95rem; }
        .modal-row span:first-child { color: #64748b; font-weight: 600; }
        .modal-row span:last-child { font-weight: 700; color: var(--text-main); }

        #toast {
            position: fixed; bottom: 25px; right: 25px; padding: 14px 22px; border-radius: var(--radius-lg);
            color: white; font-weight: 600; font-size: 0.95rem; display: none;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.25); z-index: 1000; max-width: 85vw;
        }
        .toast-success { background: #10b981; }
        .toast-error { background: #ef4444; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <img src="/static/logo.png" alt="Rojas Sport" class="brand-image" onerror="this.style.display='none';">
            <div class="brand-subtitle">INVENTARIO · CLASES & PAQUETES · FINANZAS</div>
            <div class="header-buttons">
                <a href="/descargar-excel" class="btn-top">📥 Descargar Excel</a>
                <button onclick="recargarDesdeExcel()" class="btn-top btn-reload">🔄 Sincronizar Cambios de Excel</button>
            </div>
        </header>

        <div class="nav-main">
            <button class="nav-btn active" id="nav-inventario" onclick="cambiarVista('inventario')">📦 Inventario</button>
            <button class="nav-btn" id="nav-clases" onclick="cambiarVista('clases')">🩰 Clases & Paquetes</button>
            <button class="nav-btn" id="nav-finanzas" onclick="cambiarVista('finanzas')">💰 Finanzas</button>
        </div>

        <!-- ==================== VISTA: INVENTARIO ==================== -->
        <div id="view-inventario">

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
                        <div class="picker-wrapper">
                            <input type="text" id="v_search" placeholder="Escribe para buscar (mín. 2 letras)..." oninput="buscarParaPicker('v', this.value)" autocomplete="off">
                            <div id="v_picker_list" class="picker-list"></div>
                        </div>
                        <div id="v_picker_selected" class="picker-selected">Ningún ítem seleccionado</div>
                        <input type="hidden" id="v_sku">
                    </div>
                    <div class="form-group"><label>Cantidad a Descontar</label><input type="number" id="v_cant" value="1" min="1"></div>
                    <div class="form-group"><label>Registrado por</label><input type="text" id="v_usuario" placeholder="Ej: Profesora María"></div>
                    <button class="btn-action btn-salida" id="btn-salida-action" onclick="procesarMovimiento('ventas')">Descontar Unidad</button>
                </div>

                <div class="glass-card">
                    <div class="card-title" id="form-entrada-title">📦 Registrar Entrada / Compra</div>
                    <div class="form-group">
                        <label>Buscar Ítem</label>
                        <div class="picker-wrapper">
                            <input type="text" id="e_search" placeholder="Escribe para buscar (mín. 2 letras)..." oninput="buscarParaPicker('e', this.value)" autocomplete="off">
                            <div id="e_picker_list" class="picker-list"></div>
                        </div>
                        <div id="e_picker_selected" class="picker-selected">Ningún ítem seleccionado</div>
                        <input type="hidden" id="e_sku">
                    </div>
                    <div class="form-group"><label>Cantidad Ingresada</label><input type="number" id="e_cant" value="1" min="1"></div>
                    <div class="form-group"><label>Costo Unitario de Compra (opcional)</label><input type="number" id="e_costo" min="0" step="500" placeholder="Ej: 25000"></div>
                    <div class="form-group"><label>Registrado por</label><input type="text" id="e_usuario" placeholder="Ej: Admin"></div>

                    <div class="form-group">
                        <label>Adjuntar / Tomar Foto del Producto</label>
                        <label class="foto-upload-btn" style="width: 100%; display: block;">
                            📸 Tomar o Subir Foto del Producto
                            <input type="file" accept="image/*" capture="environment" id="e_foto_input" style="display:none;" onchange="subirFotoDesdeEntrada(this.files[0])">
                        </label>
                    </div>

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
                            <tr><td colspan="6" class="empty-hint">Cargando datos...</td></tr>
                        </tbody>
                    </table>
                </div>
                <button class="btn-cargar-mas" id="btn-cargar-mas" onclick="cargarMasTabla()" style="display:none;">Cargar más</button>
            </div>
        </div>

        <!-- ==================== VISTA: CLASES & PAQUETES ==================== -->
        <div id="view-clases" style="display:none;">
            <div class="kpi-grid">
                <div class="kpi-card"><h4>Ventas Registradas</h4><div class="number" id="kpi-clases-total">—</div></div>
                <div class="kpi-card"><h4>Ingresos por Clases</h4><div class="number" id="kpi-clases-ingresos">—</div></div>
                <div class="kpi-card"><h4>Ticket Promedio</h4><div class="number" id="kpi-clases-promedio">—</div></div>
            </div>

            <div class="glass-card">
                <div class="card-title">🩰 Registrar Venta de Clase / Paquete</div>
                <div class="form-grid">
                    <div>
                        <div class="form-group">
                            <label>Tipo de Plan</label>
                            <select id="c_tipo">
                                <option>Clase Suelta</option>
                                <option>Paquete 4 Clases</option>
                                <option>Paquete 8 Clases</option>
                                <option>Mensualidad</option>
                                <option>Trimestre</option>
                                <option>Otro</option>
                            </select>
                        </div>
                        <div class="form-group"><label>Nombre de la Alumna/o</label><input type="text" id="c_nombre" placeholder="Nombre completo"></div>
                        <div class="form-group"><label>Contacto (Tel/Email)</label><input type="text" id="c_contacto" placeholder="Opcional"></div>
                    </div>
                    <div>
                        <div class="form-group"><label>Precio ($COP)</label><input type="number" id="c_precio" min="0" step="1000" placeholder="Ej: 120000"></div>
                        <div class="form-group">
                            <label>Método de Pago</label>
                            <select id="c_metodo"><option>Efectivo</option><option>Transferencia</option><option>Tarjeta</option><option>Otro</option></select>
                        </div>
                        <div class="form-group"><label>Registrado por</label><input type="text" id="c_usuario" placeholder="Ej: Profesora María"></div>
                    </div>
                </div>
                <div class="form-group"><label>Notas</label><input type="text" id="c_notas" placeholder="Opcional"></div>
                <button class="btn-action btn-entrada" onclick="registrarClase()">Registrar Venta</button>
            </div>

            <div class="glass-card">
                <div class="table-header">
                    <div class="card-title" style="margin-bottom:0;">📋 Historial de Clases & Paquetes</div>
                    <input type="text" id="c_search" class="search-input" placeholder="🔍 Buscar por nombre, plan..." onkeyup="onClasesSearch()">
                </div>
                <div class="table-wrapper">
                    <table>
                        <thead><tr><th>Fecha</th><th>Plan</th><th>Cliente</th><th>Contacto</th><th>Precio</th><th>Pago</th><th>Registró</th><th></th></tr></thead>
                        <tbody id="clases-tabla-body"><tr><td colspan="8" class="empty-hint">Cargando...</td></tr></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- ==================== VISTA: FINANZAS ==================== -->
        <div id="view-finanzas" style="display:none;">
            <div class="kpi-grid">
                <div class="kpi-card"><h4>Inversión Inicial</h4><div class="number" id="kpi-inversion">—</div></div>
                <div class="kpi-card"><h4>Ingresos Totales</h4><div class="number" id="kpi-ingresos-totales">—</div></div>
                <div class="kpi-card alert"><h4>Gastos Totales</h4><div class="number" id="kpi-gastos-totales">—</div></div>
            </div>
            <div class="kpi-grid">
                <div class="kpi-card"><h4>Ingresos Productos</h4><div class="number" id="kpi-ingresos-productos">—</div></div>
                <div class="kpi-card"><h4>Ingresos Clases</h4><div class="number" id="kpi-ingresos-clases">—</div></div>
                <div class="kpi-card"><h4>Ganancia Neta</h4><div class="number" id="kpi-ganancia-neta">—</div></div>
            </div>

            <div class="glass-card">
                <div class="card-title">🏦 Inversión Inicial</div>
                <div class="form-group"><label>Monto Total Invertido</label><input type="number" id="f_inversion" min="0" step="10000" placeholder="Ej: 5000000"></div>
                <button class="btn-action btn-entrada" onclick="actualizarInversion()">Guardar Inversión Inicial</button>
            </div>

            <div class="form-grid">
                <div class="glass-card">
                    <div class="card-title">🧾 Registrar Gasto</div>
                    <div class="form-group"><label>Concepto</label><input type="text" id="g_concepto" placeholder="Ej: Pago arriendo, servicios..."></div>
                    <div class="form-group">
                        <label>Categoría</label>
                        <select id="g_categoria">
                            <option value="operativo">Operativo</option>
                            <option value="arriendo">Arriendo</option>
                            <option value="servicios">Servicios (luz, agua, internet)</option>
                            <option value="compra_mercancia">Compra de Mercancía</option>
                            <option value="mantenimiento">Mantenimiento de Equipos</option>
                            <option value="marketing">Marketing / Publicidad</option>
                            <option value="otro">Otro</option>
                        </select>
                    </div>
                    <div class="form-group"><label>Monto ($COP)</label><input type="number" id="g_monto" min="0" step="1000"></div>
                    <div class="form-group"><label>Registrado por</label><input type="text" id="g_usuario"></div>
                    <div class="form-group"><label>Notas</label><input type="text" id="g_notas" placeholder="Opcional"></div>
                    <button class="btn-action btn-salida" onclick="registrarGasto()">Registrar Gasto</button>
                </div>

                <div class="glass-card">
                    <div class="card-title">📊 Resumen</div>
                    <p style="color:#64748b; font-size:0.9rem; line-height:1.6;">
                        Aquí puedes ver el estado financiero general de tu negocio: cuánto has invertido, cuánto has ganado en ventas de productos y clases, cuánto has gastado, y tu ganancia neta actual.
                    </p>
                    <p style="margin-top:14px; font-size:0.85rem; color:#94a3b8;">💡 Tip: cuando registras una entrada de mercancía con "costo unitario" en la pestaña Inventario, el gasto se contabiliza automáticamente aquí como "Compra de Mercancía".</p>
                </div>
            </div>

            <div class="glass-card">
                <div class="table-header">
                    <div class="card-title" style="margin-bottom:0;">📋 Historial de Gastos</div>
                    <input type="text" id="g_search" class="search-input" placeholder="🔍 Buscar concepto..." onkeyup="onGastosSearch()">
                </div>
                <div class="table-wrapper">
                    <table>
                        <thead><tr><th>Fecha</th><th>Concepto</th><th>Categoría</th><th>Monto</th><th>Registró</th><th></th></tr></thead>
                        <tbody id="gastos-tabla-body"><tr><td colspan="6" class="empty-hint">Cargando...</td></tr></tbody>
                    </table>
                </div>
            </div>
        </div>

    </div>

    <!-- Modal de detalle de producto -->
    <div id="modal-overlay" class="modal-overlay" onclick="if(event.target===this) cerrarModal()">
        <div class="modal-box">
            <button class="modal-close" onclick="cerrarModal()">✕</button>
            <img id="modal-foto" class="modal-foto" src="" alt="" style="display:none;">
            <h3 id="modal-nombre" style="font-size:1.3rem; font-weight:800; margin-bottom:14px;">---</h3>
            <div class="modal-row"><span>SKU</span><span id="modal-sku">---</span></div>
            <div class="modal-row"><span>Categoría</span><span id="modal-categoria">---</span></div>
            <div class="modal-row"><span>Stock Actual</span><span id="modal-stock">---</span></div>
            <div class="modal-row"><span>Precio de Venta</span><span id="modal-precio">---</span></div>
            <label class="foto-upload-btn" style="width:100%; display:block; text-align:center; margin-top:16px;">
                📷 Subir / Cambiar Foto
                <input type="file" accept="image/*" capture="environment" id="modal-foto-input" style="display:none;" onchange="subirFoto(this.files[0]); setTimeout(cerrarModal, 800);">
            </label>
        </div>
    </div>

    <div id="toast"></div>

    <script>
        let categoriaActual = 'productos';
        let vistaActual = 'inventario';
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

        function fmtMoney(n) {
            return '$' + Math.round(n || 0).toLocaleString('es-CO');
        }

        // ============ NAVEGACIÓN PRINCIPAL ============
        function cambiarVista(vista) {
            vistaActual = vista;
            document.getElementById('view-inventario').style.display = vista === 'inventario' ? 'block' : 'none';
            document.getElementById('view-clases').style.display = vista === 'clases' ? 'block' : 'none';
            document.getElementById('view-finanzas').style.display = vista === 'finanzas' ? 'block' : 'none';
            document.getElementById('nav-inventario').className = vista === 'inventario' ? 'nav-btn active' : 'nav-btn';
            document.getElementById('nav-clases').className = vista === 'clases' ? 'nav-btn active' : 'nav-btn';
            document.getElementById('nav-finanzas').className = vista === 'finanzas' ? 'nav-btn active' : 'nav-btn';

            if (vista === 'clases') { cargarKpisClases(); refrescarTablaClases(); }
            if (vista === 'finanzas') { cargarResumenFinanzas(); refrescarTablaGastos(); }
        }

        // ============ INVENTARIO: KPIs Y BÚSQUEDA ============
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
                            const stockTag = item.stock_actual > 0 ? `<b>${item.stock_actual} ud.</b>` : '<span style="color:#ef4444;">Agotado</span>';
                            const img = item.foto_url
                                ? `<img src="${item.foto_url}" class="autocomplete-thumb">`
                                : '<div class="autocomplete-thumb" style="display:flex;align-items:center;justify-content:center;font-size:1.2rem;">🖼️</div>';
                            return `<div class="autocomplete-item" onclick="seleccionarResultado('${item.sku}')">
                                        <div style="display:flex; align-items:center; gap:10px;">
                                            ${img}
                                            <div><strong>[${item.sku}]</strong> ${item.nombre} <small style="color:#64748b;">${catTag}</small></div>
                                        </div>
                                        <div>${stockTag}</div>
                                    </div>`;
                        }).join('');
                    }
                    listDiv.style.display = 'block';
                } catch(e) {}
            }, 250);
        }

        async function seleccionarResultado(sku) {
            document.getElementById('autocomplete-list').style.display = 'none';
            try {
                const data = await buscar(sku, '', 1, 0);
                if (data.items.length > 0) mostrarTarjetaEstado(data.items[0]);
            } catch(e) {}
        }

        function mostrarTarjetaEstado(item) {
            skuSeleccionadoParaFoto = item.sku;
            const box = document.getElementById('status-display');
            const title = document.getElementById('status-title');
            const desc = document.getElementById('status-desc');
            const foto = document.getElementById('status-foto');

            if (item.foto_url) {
                foto.src = item.foto_url + '?t=' + Date.now();
                foto.style.display = 'block';
            } else {
                foto.style.display = 'none';
            }

            title.innerText = item.nombre;
            const precioStr = item.precio_venta > 0 ? ` | ${fmtMoney(item.precio_venta)} COP` : '';

            if (item.stock_actual > 2) {
                box.className = 'status-card status-available';
                desc.innerText = `✅ Disponible: ${item.stock_actual} unidades (${item.sku})${precioStr}`;
            } else if (item.stock_actual > 0) {
                box.className = 'status-card status-low';
                desc.innerText = `⚠️ Últimas unidades: ${item.stock_actual} disponibles (${item.sku})${precioStr}`;
            } else {
                box.className = 'status-card status-empty';
                desc.innerText = `❌ AGOTADO: 0 unidades en inventario (${item.sku})${precioStr}`;
            }
            box.style.display = 'block';
        }

        // ============ MODAL DE DETALLE (BOTÓN "VER") ============
        async function verDetalle(sku) {
            try {
                const data = await buscar(sku, '', 1, 0);
                if (data.items.length > 0) mostrarModal(data.items[0]);
            } catch(e) { mostrarToast('No se pudo cargar el detalle', false); }
        }

        function mostrarModal(item) {
            document.getElementById('modal-nombre').innerText = item.nombre;
            document.getElementById('modal-sku').innerText = item.sku;
            document.getElementById('modal-categoria').innerText = item.categoria === 'productos' ? 'Ropa & Productos' : 'Equipamiento & Activos';
            document.getElementById('modal-stock').innerText = item.stock_actual + ' unidades';
            document.getElementById('modal-precio').innerText = item.precio_venta > 0 ? fmtMoney(item.precio_venta) + ' COP' : '—';
            const foto = document.getElementById('modal-foto');
            if (item.foto_url) { foto.src = item.foto_url + '?t=' + Date.now(); foto.style.display = 'block'; }
            else { foto.style.display = 'none'; }
            skuSeleccionadoParaFoto = item.sku;
            document.getElementById('modal-overlay').style.display = 'flex';
        }

        function cerrarModal() { document.getElementById('modal-overlay').style.display = 'none'; }

        // ============ FOTOS ============
        async function subirFoto(file) {
            if (!file || !skuSeleccionadoParaFoto) {
                mostrarToast('Selecciona primero un producto', false);
                return;
            }
            const formData = new FormData();
            formData.append('archivo', file);
            try {
                const res = await fetch(`/fotos/${skuSeleccionadoParaFoto}`, { method: 'POST', body: formData });
                const data = await res.json();
                if (res.ok) {
                    mostrarToast('📷 Foto actualizada con éxito', true);
                    seleccionarResultado(skuSeleccionadoParaFoto);
                    refrescarTabla();
                } else {
                    mostrarToast(data.detail || 'Error al subir la foto', false);
                }
            } catch(e) {
                mostrarToast('Error en la conexión al subir foto', false);
            }
        }

        async function subirFotoDesdeEntrada(file) {
            const sku = document.getElementById('e_sku').value;
            if (!sku) {
                mostrarToast('Por favor selecciona primero un producto en la lista de entradas', false);
                return;
            }
            if (!file) return;

            const formData = new FormData();
            formData.append('archivo', file);
            try {
                const res = await fetch(`/fotos/${sku}`, { method: 'POST', body: formData });
                const data = await res.json();
                if (res.ok) {
                    mostrarToast('📸 Foto del producto registrada con éxito', true);
                    refrescarTabla();
                } else {
                    mostrarToast(data.detail || 'Error al guardar la foto', false);
                }
            } catch(e) {
                mostrarToast('Error de conexión al subir la foto', false);
            }
        }

        // ============ TABS PRODUCTOS / EQUIPAMIENTO ============
        function cambiarPestana(cat) {
            categoriaActual = cat;
            document.getElementById('btn-tab-productos').className = cat === 'productos' ? 'tab-btn active' : 'tab-btn';
            document.getElementById('btn-tab-equipamiento').className = cat === 'equipamiento' ? 'tab-btn active' : 'tab-btn';
            document.getElementById('tabla-titulo').innerText = cat === 'productos' ? '📋 Listado de Productos' : '🪑 Listado de Equipamiento & Activos';
            document.getElementById('form-salida-title').innerText = cat === 'productos' ? '🛍️ Registrar Venta / Salida' : '🔻 Registrar Salida / Baja de Activo';
            document.getElementById('btn-salida-action').innerText = cat === 'productos' ? 'Descontar Unidad' : 'Registrar Salida';

            cargarKpis();
            refrescarTabla();
        }

        // ============ SELECTOR DE ÍTEM PARA MOVIMIENTOS (arreglado) ============
        function buscarParaPicker(prefix, term) {
            const listDiv = document.getElementById(prefix + '_picker_list');
            if (term.trim().length < 2) { listDiv.style.display = 'none'; return; }
            debounce(async () => {
                try {
                    const data = await buscar(term, categoriaActual, 15, 0);
                    if (data.items.length === 0) {
                        listDiv.innerHTML = '<div class="autocomplete-item" style="color:#94a3b8;">Sin resultados</div>';
                    } else {
                        listDiv.innerHTML = data.items.map(i => {
                            const nombreSeguro = i.nombre.replace(/'/g, "\\\\'");
                            return `<div class="autocomplete-item" onclick="seleccionarParaPicker('${prefix}', '${i.sku}', '${nombreSeguro}', ${i.stock_actual})">
                                        <div><strong>[${i.sku}]</strong> ${i.nombre}</div>
                                        <div>${i.stock_actual} ud.</div>
                                    </div>`;
                        }).join('');
                    }
                    listDiv.style.display = 'block';
                } catch(e) {}
            }, 250);
        }

        function seleccionarParaPicker(prefix, sku, nombre, stock) {
            document.getElementById(prefix + '_sku').value = sku;
            document.getElementById(prefix + '_search').value = '[' + sku + '] ' + nombre;
            document.getElementById(prefix + '_picker_selected').innerHTML = `✅ Seleccionado: <strong>${nombre}</strong> — Stock actual: ${stock}`;
            document.getElementById(prefix + '_picker_list').style.display = 'none';
        }

        async function procesarMovimiento(tipo) {
            const prefix = tipo === 'ventas' ? 'v' : 'e';
            const sku = document.getElementById(prefix + '_sku').value;
            const cant = parseInt(document.getElementById(prefix + '_cant').value);
            const usr = document.getElementById(prefix + '_usuario').value.trim();
            const costo = tipo === 'entradas' ? (parseFloat(document.getElementById('e_costo').value) || 0) : 0;

            if (!sku) { mostrarToast('Por favor selecciona un ítem de la lista de resultados', false); return; }
            if (!cant || cant < 1) { mostrarToast('La cantidad debe ser mayor a 0', false); return; }
            if (!usr) { mostrarToast('Por favor escribe tu nombre en "Registrado por"', false); return; }

            try {
                const res = await fetch(`/movimientos/${tipo}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sku: sku, cantidad: cant, registrado_por: usr, costo_unitario: costo })
                });
                const data = await res.json();
                if (res.ok) {
                    mostrarToast(`✅ Registrado correctamente: ${data.item.nombre}`, true);
                    document.getElementById(prefix + '_sku').value = '';
                    document.getElementById(prefix + '_search').value = '';
                    document.getElementById(prefix + '_picker_selected').innerHTML = 'Ningún ítem seleccionado';
                    if (tipo === 'entradas') document.getElementById('e_costo').value = '';
                    cargarKpis();
                    refrescarTabla();
                    if (tipo === 'entradas' && costo > 0) mostrarToast('💰 Se registró también un gasto de compra de mercancía', true);
                } else {
                    mostrarToast(`⚠️ Error: ${data.detail}`, false);
                }
            } catch(e) {
                mostrarToast('Error de red al procesar el movimiento', false);
            }
        }

        // ============ TABLA DE INVENTARIO ============
        async function cargarTabla(acumular = false) {
            if (!acumular) paginaTabla = 0;
            const query = document.getElementById('search').value;
            const offset = paginaTabla * PAGE_SIZE;

            try {
                const data = await buscar(query, categoriaActual, PAGE_SIZE, offset);
                const tbody = document.getElementById('tabla-body');

                if (!acumular && data.items.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="empty-hint">No se encontraron ítems.</td></tr>';
                    document.getElementById('btn-cargar-mas').style.display = 'none';
                    return;
                }

                const rowsHtml = data.items.map(item => {
                    let badgeClass = 'badge-success';
                    let badgeText = 'Disponible';
                    if (item.stock_actual <= 0) { badgeClass = 'badge-danger'; badgeText = 'Agotado'; }
                    else if (item.stock_actual <= 2) { badgeClass = 'badge-warning'; badgeText = 'Bajo Stock'; }

                    const img = item.foto_url
                        ? `<img src="${item.foto_url}" class="row-thumb">`
                        : '<div class="row-thumb" style="display:flex;align-items:center;justify-content:center;font-size:1.2rem;">🖼️</div>';

                    return `<tr>
                        <td>${img}</td>
                        <td><strong>${item.sku}</strong></td>
                        <td>${item.nombre}</td>
                        <td><span class="badge ${badgeClass}">${badgeText}</span></td>
                        <td><strong>${item.stock_actual}</strong></td>
                        <td><button class="row-action-btn" onclick="verDetalle('${item.sku}')">Ver</button></td>
                    </tr>`;
                }).join('');

                if (acumular) tbody.innerHTML += rowsHtml;
                else tbody.innerHTML = rowsHtml;

                const masDisponibles = (offset + data.items.length) < data.total;
                document.getElementById('btn-cargar-mas').style.display = masDisponibles ? 'block' : 'none';
            } catch(e) {
                document.getElementById('tabla-body').innerHTML = '<tr><td colspan="6" class="empty-hint">Error al cargar datos.</td></tr>';
            }
        }

        function onTablaSearch() { debounce(() => cargarTabla(false), 300); }
        function cargarMasTabla() { paginaTabla++; cargarTabla(true); }
        function refrescarTabla() { cargarTabla(false); }

        async function recargarDesdeExcel() {
            try {
                const res = await fetch('/sincronizar');
                if (res.ok) {
                    mostrarToast('🔄 Excel sincronizado con éxito', true);
                    cargarKpis();
                    refrescarTabla();
                    if (vistaActual === 'clases') { cargarKpisClases(); refrescarTablaClases(); }
                    if (vistaActual === 'finanzas') { cargarResumenFinanzas(); refrescarTablaGastos(); }
                }
            } catch(e) { mostrarToast('Error al re-sincronizar', false); }
        }

        // ============ CLASES & PAQUETES ============
        async function cargarKpisClases() {
            try {
                const res = await fetch('/clases/kpis');
                const data = await res.json();
                document.getElementById('kpi-clases-total').innerText = data.total_ventas;
                document.getElementById('kpi-clases-ingresos').innerText = fmtMoney(data.ingresos_totales);
                document.getElementById('kpi-clases-promedio').innerText = fmtMoney(data.ticket_promedio);
            } catch(e) {}
        }

        async function registrarClase() {
            const tipo_plan = document.getElementById('c_tipo').value;
            const nombre_cliente = document.getElementById('c_nombre').value.trim();
            const contacto = document.getElementById('c_contacto').value.trim();
            const precio = parseFloat(document.getElementById('c_precio').value);
            const metodo_pago = document.getElementById('c_metodo').value;
            const registrado_por = document.getElementById('c_usuario').value.trim();
            const notas = document.getElementById('c_notas').value.trim();

            if (!nombre_cliente) { mostrarToast('Escribe el nombre de la alumna/o', false); return; }
            if (!precio || precio <= 0) { mostrarToast('El precio debe ser mayor a 0', false); return; }
            if (!registrado_por) { mostrarToast('Escribe quién registra la venta', false); return; }

            try {
                const res = await fetch('/clases', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ tipo_plan, nombre_cliente, contacto, precio, metodo_pago, registrado_por, notas })
                });
                const data = await res.json();
                if (res.ok) {
                    mostrarToast('✅ Venta registrada con éxito', true);
                    document.getElementById('c_nombre').value = '';
                    document.getElementById('c_contacto').value = '';
                    document.getElementById('c_precio').value = '';
                    document.getElementById('c_notas').value = '';
                    cargarKpisClases();
                    refrescarTablaClases();
                } else {
                    mostrarToast(data.detail || 'Error al registrar', false);
                }
            } catch(e) { mostrarToast('Error de red', false); }
        }

        async function cargarTablaClases() {
            const q = document.getElementById('c_search').value;
            try {
                const res = await fetch('/clases?' + new URLSearchParams({ q, limit: 50, offset: 0 }));
                const data = await res.json();
                const tbody = document.getElementById('clases-tabla-body');
                if (data.items.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="8" class="empty-hint">Sin registros.</td></tr>';
                    return;
                }
                tbody.innerHTML = data.items.map(c => `
                    <tr>
                        <td>${c.fecha}</td>
                        <td>${c.tipo_plan}</td>
                        <td>${c.nombre_cliente}</td>
                        <td>${c.contacto || '—'}</td>
                        <td>${fmtMoney(c.precio)}</td>
                        <td>${c.metodo_pago || '—'}</td>
                        <td>${c.registrado_por}</td>
                        <td><button class="row-action-btn" onclick="eliminarClase(${c.id})">🗑️</button></td>
                    </tr>`).join('');
            } catch(e) {
                document.getElementById('clases-tabla-body').innerHTML = '<tr><td colspan="8" class="empty-hint">Error al cargar.</td></tr>';
            }
        }
        function refrescarTablaClases() { cargarTablaClases(); }
        function onClasesSearch() { debounce(() => cargarTablaClases(), 300); }

        async function eliminarClase(id) {
            if (!confirm('¿Eliminar este registro de venta?')) return;
            try {
                const res = await fetch('/clases/' + id, { method: 'DELETE' });
                if (res.ok) { mostrarToast('Registro eliminado', true); cargarKpisClases(); refrescarTablaClases(); }
            } catch(e) {}
        }

        // ============ FINANZAS ============
        async function cargarResumenFinanzas() {
            try {
                const res = await fetch('/finanzas/resumen');
                const data = await res.json();
                document.getElementById('kpi-inversion').innerText = fmtMoney(data.inversion_inicial);
                document.getElementById('kpi-ingresos-totales').innerText = fmtMoney(data.ingresos_totales);
                document.getElementById('kpi-gastos-totales').innerText = fmtMoney(data.total_gastos);
                document.getElementById('kpi-ingresos-productos').innerText = fmtMoney(data.ingresos_productos);
                document.getElementById('kpi-ingresos-clases').innerText = fmtMoney(data.ingresos_clases);
                document.getElementById('kpi-ganancia-neta').innerText = fmtMoney(data.ganancia_neta);
                document.getElementById('f_inversion').value = data.inversion_inicial || '';
            } catch(e) {}
        }

        async function actualizarInversion() {
            const monto = parseFloat(document.getElementById('f_inversion').value);
            if (isNaN(monto) || monto < 0) { mostrarToast('Ingresa un monto válido', false); return; }
            try {
                const res = await fetch('/finanzas/inversion', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ monto })
                });
                if (res.ok) { mostrarToast('✅ Inversión inicial actualizada', true); cargarResumenFinanzas(); }
            } catch(e) { mostrarToast('Error de red', false); }
        }

        async function registrarGasto() {
            const concepto = document.getElementById('g_concepto').value.trim();
            const categoria = document.getElementById('g_categoria').value;
            const monto = parseFloat(document.getElementById('g_monto').value);
            const registrado_por = document.getElementById('g_usuario').value.trim();
            const notas = document.getElementById('g_notas').value.trim();

            if (!concepto) { mostrarToast('Escribe el concepto del gasto', false); return; }
            if (!monto || monto <= 0) { mostrarToast('El monto debe ser mayor a 0', false); return; }
            if (!registrado_por) { mostrarToast('Escribe quién registra el gasto', false); return; }

            try {
                const res = await fetch('/gastos', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ concepto, categoria, monto, registrado_por, notas })
                });
                const data = await res.json();
                if (res.ok) {
                    mostrarToast('✅ Gasto registrado', true);
                    document.getElementById('g_concepto').value = '';
                    document.getElementById('g_monto').value = '';
                    document.getElementById('g_notas').value = '';
                    refrescarTablaGastos();
                    cargarResumenFinanzas();
                } else { mostrarToast(data.detail || 'Error', false); }
            } catch(e) { mostrarToast('Error de red', false); }
        }

        async function cargarTablaGastos() {
            const q = document.getElementById('g_search').value;
            try {
                const res = await fetch('/gastos?' + new URLSearchParams({ q, limit: 50, offset: 0 }));
                const data = await res.json();
                const tbody = document.getElementById('gastos-tabla-body');
                if (data.items.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="empty-hint">Sin registros.</td></tr>';
                    return;
                }
                tbody.innerHTML = data.items.map(g => `
                    <tr>
                        <td>${g.fecha}</td>
                        <td>${g.concepto}</td>
                        <td><span class="badge badge-warning">${g.categoria}</span></td>
                        <td>${fmtMoney(g.monto)}</td>
                        <td>${g.registrado_por}</td>
                        <td><button class="row-action-btn" onclick="eliminarGasto(${g.id})">🗑️</button></td>
                    </tr>`).join('');
            } catch(e) {
                document.getElementById('gastos-tabla-body').innerHTML = '<tr><td colspan="6" class="empty-hint">Error al cargar.</td></tr>';
            }
        }
        function refrescarTablaGastos() { cargarTablaGastos(); }
        function onGastosSearch() { debounce(() => cargarTablaGastos(), 300); }

        async function eliminarGasto(id) {
            if (!confirm('¿Eliminar este gasto?')) return;
            try {
                const res = await fetch('/gastos/' + id, { method: 'DELETE' });
                if (res.ok) { mostrarToast('Gasto eliminado', true); refrescarTablaGastos(); cargarResumenFinanzas(); }
            } catch(e) {}
        }

        window.onload = () => {
            cargarKpis();
            refrescarTabla();
        };
    </script>
</body>
</html>'''


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)