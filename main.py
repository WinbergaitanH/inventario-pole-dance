Python
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import pandas as pd
import os

app = FastAPI(title="Control de Inventario Pole Dance")

EXCEL_PATH = "Control_Inventario_Pole_Dance.xlsx"

# Función para cargar los datos del Excel al iniciar la app
def cargar_inventario_desde_excel():
    if not os.path.exists(EXCEL_PATH):
        return {}
    
    # Leer el catálogo del Excel
    df_catalogo = pd.read_excel(EXCEL_PATH, sheet_name='🏷️ Catálogo Productos', header=1)
    
    inventario = {}
    for index, row in df_catalogo.iterrows():
        sku = str(row.iloc[0]).strip()
        if pd.notna(sku) and sku != "nan":
            nombre = f"{row.iloc[2]} ({row.iloc[1]}) - Talla {row.iloc[3]}"
            stock_ini = int(row.iloc[6]) if pd.notna(row.iloc[6]) else 0
            inventario[sku] = {
                "nombre": nombre,
                "stock_inicial": stock_ini,
                "entradas": 0,
                "ventas": 0
            }
    return inventario

# Guardar base en memoria activa
inventario_db = cargar_inventario_desde_excel()

class Movimiento(BaseModel):
    sku: str
    cantidad: int
    registrado_por: str

@app.get("/")
def home():
    return {"mensaje": "Servidor de Inventario listo y conectado al Excel"}

# Consultar el stock actual
@app.get("/stock")
def ver_stock():
    resultado = {}
    for sku, item in inventario_db.items():
        stock_actual = item["stock_inicial"] + item["entradas"] - item["ventas"]
        resultado[sku] = {
            "nombre": item["nombre"],
            "stock_actual": stock_actual
        }
    return resultado

# Registrar Entradas desde el celular
@app.post("/entradas")
def registrar_entrada(mov: Movimiento):
    if mov.sku not in inventario_db:
        raise HTTPException(status_code=404, detail="El código de prenda no existe")
    
    inventario_db[mov.sku]["entradas"] += mov.cantidad
    stock = inventario_db[mov.sku]["stock_inicial"] + inventario_db[mov.sku]["entradas"] - inventario_db[mov.sku]["ventas"]
    return {"mensaje": f"Se ingresaron {mov.cantidad} unidades a {inventario_db[mov.sku]['nombre']}", "stock_actual": stock}

# Registrar Ventas y descontar del celular
@app.post("/ventas")
def registrar_venta(mov: Movimiento):
    if mov.sku not in inventario_db:
        raise HTTPException(status_code=404, detail="El código de prenda no existe")
    
    item = inventario_db[mov.sku]
    stock_disponible = item["stock_inicial"] + item["entradas"] - item["ventas"]
    
    if mov.cantidad > stock_disponible:
        raise HTTPException(status_code=400, detail=f"Stock insuficiente. Solo quedan {stock_disponible} unidades disponibles.")
    
    item["ventas"] += mov.cantidad
    nuevo_stock = stock_disponible - mov.cantidad
    return {"mensaje": f"Venta registrada por {mov.registrado_por}", "stock_restante": nuevo_