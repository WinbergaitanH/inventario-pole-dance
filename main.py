from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel, Field
import pandas as pd
import os
import uvicorn

app = FastAPI(
    title="Pole Dance Rojas Sport - Inventario & Activos",
    description="Sistema de control de inventario y disponibilidad para Pole Dance Rojas Sport",
    version="8.0.0"
)

EXCEL_PATH = "Control_Inventario_Pole_Dance.xlsx"

# Diccionarios globales para mantener el registro de movimientos en memoria
historial_entradas = {}  # {sku: cantidad}
historial_ventas = {}    # {sku: cantidad}

def cargar_inventario_desde_excel():
    inventario = {}
    
    if os.path.exists(EXCEL_PATH):
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH)
            
            # 1. Leer Catálogo de Productos (Encabezados en la Fila 1)
            for sheet in excel_file.sheet_names:
                if any(kw in sheet.lower() for kw in ["catálogo", "catalogo", "productos"]):
                    df_cat = pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=0)
                    
                    # Normalizar nombres de columnas a minúsculas y sin espacios
                    col_map = {str(c).strip().lower(): c for c in df_cat.columns}
                    
                    # Buscar columnas relevantes
                    col_sku = next((col_map[k] for k in col_map if 'sku' in k or 'código' in k or 'codigo' in k), df_cat.columns[0])
                    col_desc = next((col_map[k] for k in col_map if 'descrip' in k or 'estilo' in k or 'nombre' in k), None)
                    col_tipo = next((col_map[k] for k in col_map if 'tipo' in k or 'prenda' in k), None)
                    col_talla = next((col_map[k] for k in col_map if 'talla' in k), None)
                    col_stock = next((col_map[k] for k in col_map if 'stock inicial' in k or 'inicial' in k or 'stock' in k), None)

                    for _, row in df_cat.iterrows():
                        sku = str(row[col_sku]).strip() if pd.notna(row[col_sku]) else ""
                        if sku and sku.lower() not in ['nan', 'none', 'sku', 'sku / código', 'código', 'codigo', '']:
                            desc = str(row[col_desc]).strip() if col_desc and pd.notna(row[col_desc]) else ""
                            tipo = str(row[col_tipo]).strip() if col_tipo and pd.notna(row[col_tipo]) else ""
                            talla = str(row[col_talla]).strip() if col_talla and pd.notna(row[col_talla]) else ""

                            # Construir el nombre dinámico del producto
                            partes = []
                            if desc and desc.lower() != 'nan':
                                partes.append(desc)
                            if tipo and tipo.lower() != 'nan':
                                partes.append(f"({tipo})")
                            if talla and talla.lower() != 'nan':
                                partes.append(f"- Talla {talla}")
                            
                            nombre_completo = " ".join(partes) if partes else sku

                            try:
                                val_stock = row[col_stock] if col_stock and pd.notna(row[col_stock]) else 0
                                stock_ini = int(float(val_stock))
                            except (ValueError, TypeError):
                                stock_ini = 0

                            inventario[sku] = {
                                "nombre": nombre_completo,
                                "categoria": "productos",
                                "stock_inicial": stock_ini,
                                "entradas": historial_entradas.get(sku, 0),
                                "ventas": historial_ventas.get(sku, 0)
                            }

            # 2. Leer Activos Fijos / Equipamiento (Encabezados en la Fila 1)
            for sheet in excel_file.sheet_names:
                if any(kw in sheet.lower() for kw in ["activos", "equipamiento", "activos fijos"]):
                    df_act = pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=0)
                    col_map = {str(c).strip().lower(): c for c in df_act.columns}
                    
                    col_sku = next((col_map[k] for k in col_map if 'código' in k or 'codigo' in k or 'sku' in k or 'activo' in k), df_act.columns[0])
                    col_nombre = next((col_map[k] for k in col_map if 'nombre' in k or 'activo' in k or 'descrip' in k), None)
                    col_cant = next((col_map[k] for k in col_map if 'cantidad' in k or 'stock' in k or 'inicial' in k), None)

                    for _, row in df_act.iterrows():
                        sku = str(row[col_sku]).strip() if pd.notna(row[col_sku]) else ""
                        if sku and sku.lower() not in ['nan', 'none', 'código activo', 'codigo activo', 'sku', '']:
                            nombre_activo = str(row[col_nombre]).strip() if col_nombre and pd.notna(row[col_nombre]) else sku

                            try:
                                val_cant = row[col_cant] if col_cant and pd.notna(row[col_cant]) else 0
                                stock_ini = int(float(val_cant))
                            except (ValueError, TypeError):
                                stock_ini = 0

                            inventario[sku] = {
                                "nombre": nombre_activo if nombre_activo.lower() != 'nan' else sku,
                                "categoria": "equipamiento",
                                "stock_inicial": stock_ini,
                                "entradas": historial_entradas.get(sku, 0),
                                "ventas": historial_ventas.get(sku, 0)
                            }

        except Exception as e:
            print(f"⚠️ Error leyendo Excel: {e}")

    # Datos por defecto si el Excel no está presente o falla la lectura
    if not inventario:
        items_base = {
            "EQ-TUBO": ("Tubo Pole Dance Profesional", "equipamiento"),
            "EQ-MAT": ("Mat / Colchoneta de Caída", "equipamiento")
        }
        for sku, (nombre, cat) in items_base.items():
            inventario[sku] = {
                "nombre": nombre,
                "categoria": cat,
                "stock_inicial": 5,
                "entradas": historial_entradas.get(sku, 0),
                "ventas": historial_ventas.get(sku, 0)
            }

    return inventario

# Carga inicial de datos
inventario_db = cargar_inventario_desde_excel()

class Movimiento(BaseModel):
    sku: str = Field(..., description="Código SKU del elemento")
    cantidad: int = Field(..., description="Cantidad de unidades", gt=0)
    registrado_por: str = Field(..., description="Nombre del responsable")

@app.get("/", response_class=HTMLResponse, tags=["Interfaz"])
def interfaz_usuario():
    return '''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
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

        .header-buttons {
            display: flex;
            gap: 10px;
            justify-content: center;
            flex-wrap: wrap;
        }

        .btn-top {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(255, 255, 255, 0.15);
            color: #fff;
            padding: 8px 16px;
            border-radius: 30px;
            text-decoration: none;
            font-size: 0.85rem;
            font-weight: 600;
            border: 1px solid rgba(255, 255, 255, 0.25);
            transition: all 0.2s ease;
            cursor: pointer;
        }

        .btn-top:hover {
            background: rgba(255, 255, 255, 0.25);
            transform: translateY(-2px);
        }

        .btn-reload {
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%);
            border: none;
        }

        .glass-card {
            background: var(--card-bg);
            border-radius: var(--radius-xl);
            padding: 24px;
            box-shadow: var(--shadow);
            margin-bottom: 25px;
            border: 1px solid rgba(255, 255, 255, 0.8);
        }

        .card-title {
            font-size: 1.15rem;
            font-weight: 700;
            color: var(--text-main);
            margin-bottom: 18px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .search-box-wrapper { position: relative; }
        .search-box-wrapper input {
            width: 100%;
            padding: 14px 16px;
            border: 2px solid #e2e8f0;
            border-radius: var(--radius-lg);
            font-size: 1rem;
            font-weight: 500;
            color: var(--text-main);
            background-color: #f8fafc;
            outline: none;
        }

        .search-box-wrapper input:focus {
            border-color: var(--primary);
            background-color: #fff;
            box-shadow: 0 0 0 4px var(--accent-glow);
        }

        .autocomplete-results {
            position: absolute;
            top: 105%;
            left: 0; right: 0;
            background: white;
            border-radius: var(--radius-lg);
            box-shadow: 0 10px 25px rgba(0,0,0,0.2);
            border: 1px solid #cbd5e1;
            max-height: 220px;
            overflow-y: auto;
            z-index: 100;
            display: none;
        }

        .autocomplete-item {
            padding: 12px 16px;
            cursor: pointer;
            border-bottom: 1px solid #f1f5f9;
            font-size: 0.95rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .autocomplete-item:hover {
            background-color: #f0fdf4;
            color: var(--primary-dark);
        }

        .status-card {
            margin-top: 18px;
            padding: 18px 20px;
            border-radius: var(--radius-lg);
            display: none;
            text-align: center;
        }

        .status-card h3 { font-size: 1.3rem; font-weight: 800; margin-bottom: 4px; }
        .status-card p { font-size: 0.95rem; font-weight: 500; }

        .status-available { background: var(--success-bg); color: #047857; border: 1.5px solid #a7f3d0; }
        .status-low { background: var(--warning-bg); color: #b45309; border: 1.5px solid #fde68a; }
        .status-empty { background: var(--danger-bg); color: #b91c1c; border: 1.5px solid #fca5a5; }

        .tab-group {
            display: flex;
            gap: 10px;
            margin-bottom: 25px;
            background: rgba(0, 0, 0, 0.25);
            padding: 6px;
            border-radius: 50px;
            backdrop-filter: blur(8px);
        }

        .tab-btn {
            flex: 1;
            padding: 12px 18px;
            border: none;
            background: transparent;
            color: #94a3b8;
            font-weight: 700;
            font-size: 0.92rem;
            border-radius: 40px;
            cursor: pointer;
        }

        .tab-btn.active {
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%);
            color: white;
            box-shadow: 0 4px 15px var(--accent-glow);
        }

        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
            margin-bottom: 25px;
        }

        @media (max-width: 640px) { .kpi-grid { grid-template-columns: 1fr; } }

        .kpi-card {
            background: var(--card-bg);
            padding: 18px;
            border-radius: var(--radius-lg);
            text-align: center;
            box-shadow: var(--shadow);
            border: 1px solid rgba(255, 255, 255, 0.8);
        }

        .kpi-card h4 {
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: var(--text-muted);
            margin-bottom: 6px;
            font-weight: 700;
        }

        .kpi-card .number { font-size: 1.8rem; font-weight: 800; color: var(--text-main); }
        .kpi-card.alert .number { color: var(--danger); }

        .form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 25px;
        }

        @media (max-width: 768px) { .form-grid { grid-template-columns: 1fr; } }

        .form-group { margin-bottom: 14px; }
        .form-group label {
            display: block;
            font-size: 0.82rem;
            font-weight: 700;
            color: #475569;
            margin-bottom: 6px;
            text-transform: uppercase;
        }

        .form-group input, .form-group select {
            width: 100%;
            padding: 11px 14px;
            border: 1.5px solid #cbd5e1;
            border-radius: var(--radius-md);
            font-size: 0.95rem;
            outline: none;
            background: #f8fafc;
        }

        .btn-action {
            width: 100%;
            padding: 13px;
            border: none;
            border-radius: var(--radius-md);
            font-weight: 700;
            font-size: 0.95rem;
            cursor: pointer;
            color: white;
            margin-top: 6px;
        }

        .btn-salida { background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); }
        .btn-entrada { background: linear-gradient(135deg, #10b981 0%, #059669 100%); }

        .table-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            flex-wrap: wrap;
            gap: 12px;
        }

        .search-input {
            padding: 10px 16px;
            border: 1.5px solid #cbd5e1;
            border-radius: 30px;
            font-size: 0.9rem;
            width: 260px;
            outline: none;
            background: #f8fafc;
        }

        .table-wrapper { overflow-x: auto; border-radius: var(--radius-lg); border: 1px solid #e2e8f0; }
        table { width: 100%; border-collapse: collapse; background: white; text-align: left; }
        th { background: #f8fafc; color: #475569; font-size: 0.8rem; text-transform: uppercase; font-weight: 700; padding: 14px 16px; border-bottom: 1px solid #e2e8f0; }
        td { padding: 14px 16px; border-bottom: 1px solid #f1f5f9; font-size: 0.92rem; color: #334155; }
        
        .row-action-btn {
            background: rgba(217, 70, 239, 0.1);
            color: var(--primary-dark);
            border: 1px solid var(--primary);
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.78rem;
            font-weight: 700;
            cursor: pointer;
        }

        .badge { padding: 4px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 700; display: inline-block; }
        .badge-success { background: var(--success-bg); color: #047857; }
        .badge-warning { background: var(--warning-bg); color: #b45309; }
        .badge-danger { background: var(--danger-bg); color: #b91c1c; }

        #toast {
            position: fixed; bottom: 25px; right: 25px;
            padding: 14px 22px; border-radius: var(--radius-lg);
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
            <div class="card-title">🔍 Consulta Rápida ("Escribe una palabra...")</div>
            <div class="search-box-wrapper">
                <input type="text" id="lookup-input" placeholder="Escribe 'Short', 'Top', 'Velvet', 'Barra' o un SKU..." oninput="buscarAutocomplete(this.value)" autocomplete="off">
                <div id="autocomplete-list" class="autocomplete-results"></div>
            </div>

            <div id="status-display" class="status-card">
                <h3 id="status-title">---</h3>
                <p id="status-desc">---</p>
            </div>
        </div>

        <div class="tab-group">
            <button class="tab-btn active" id="btn-tab-productos" onclick="cambiarPestana('productos')">👗 Ropa & Productos</button>
            <button class="tab-btn" id="btn-tab-equipamiento" onclick="cambiarPestana('equipamiento')">🪑 Equipamiento & Activos</button>
        </div>

        <div class="kpi-grid">
            <div class="kpi-card">
                <h4>Catálogo Activo</h4>
                <div class="number" id="kpi-total-skus">0</div>
            </div>
            <div class="kpi-card">
                <h4>Stock Físico Total</h4>
                <div class="number" id="kpi-total-stock">0</div>
            </div>
            <div class="kpi-card alert">
                <h4>Ítems Agotados</h4>
                <div class="number" id="kpi-sin-stock">0</div>
            </div>
        </div>

        <div class="form-grid">
            <div class="glass-card">
                <div class="card-title" id="form-salida-title">🛍️ Registrar Venta / Salida</div>
                <div class="form-group">
                    <label>Buscar o Seleccionar Ítem</label>
                    <input type="text" id="v_search" placeholder="Escribe para filtrar lista..." oninput="filtrarSelect('v_sku', this.value)" style="margin-bottom: 6px;">
                    <select id="v_sku" size="4" style="height: 110px;"></select>
                </div>
                <div class="form-group">
                    <label>Cantidad a Descontar</label>
                    <input type="number" id="v_cant" value="1" min="1">
                </div>
                <div class="form-group">
                    <label>Registrado por</label>
                    <input type="text" id="v_usuario" placeholder="Ej: Profesora María">
                </div>
                <button class="btn-action btn-salida" id="btn-salida-action" onclick="procesarMovimiento('ventas')">Descontar Unidad</button>
            </div>

            <div class="glass-card">
                <div class="card-title" id="form-entrada-title">📦 Registrar Entrada / Compra</div>
                <div class="form-group">
                    <label>Buscar o Seleccionar Ítem</label>
                    <input type="text" id="e_search" placeholder="Escribe para filtrar lista..." oninput="filtrarSelect('e_sku', this.value)" style="margin-bottom: 6px;">
                    <select id="e_sku" size="4" style="height: 110px;"></select>
                </div>
                <div class="form-group">
                    <label>Cantidad Ingresada</label>
                    <input type="number" id="e_cant" value="1" min="1">
                </div>
                <div class="form-group">
                    <label>Registrado por</label>
                    <input type="text" id="e_usuario" placeholder="Ej: Admin">
                </div>
                <button class="btn-action btn-entrada" onclick="procesarMovimiento('entradas')">Ingresar al Stock</button>
            </div>
        </div>

        <div class="glass-card">
            <div class="table-header">
                <div class="card-title" id="tabla-titulo" style="margin-bottom:0;">📋 Listado de Productos</div>
                <input type="text" id="search" class="search-input" placeholder="🔍 Escribe una palabra..." onkeyup="filtrarTabla()">
            </div>
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>SKU</th>
                            <th>Descripción / Producto</th>
                            <th>Estado</th>
                            <th>Stock</th>
                            <th>Acción Rápida</th>
                        </tr>
                    </thead>
                    <tbody id="tabla-body">
                        <tr><td colspan="5" style="text-align:center;">Cargando inventario...</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="toast"></div>

    <script>
        let inventarioGlobal = {};
        let categoriaActual = 'productos';

        async function cargarInventario() {
            try {
                const res = await fetch('/stock');
                inventarioGlobal = await res.json();
                renderizarInterface();
            } catch(e) {
                mostrarToast("Error de conexión con el servidor", false);
            }
        }

        async function recargarDesdeExcel() {
            try {
                const res = await fetch('/recargar', { method: 'POST' });
                const data = await res.json();
                if (res.ok) {
                    mostrarToast(data.mensaje, true);
                    cargarInventario();
                } else {
                    mostrarToast("Error al recargar el archivo Excel", false);
                }
            } catch(e) {
                mostrarToast("Error de conexión con el servidor", false);
            }
        }

        function buscarAutocomplete(query) {
            const listDiv = document.getElementById('autocomplete-list');
            const term = query.trim().toLowerCase();
            
            if (term.length === 0) {
                listDiv.style.display = 'none';
                return;
            }

            let html = '';
            let coincidencias = 0;

            for (const [sku, item] of Object.entries(inventarioGlobal)) {
                const textoCompleto = `${sku} ${item.nombre}`.toLowerCase();
                if (textoCompleto.includes(term)) {
                    coincidencias++;
                    const catTag = item.categoria === 'productos' ? '👗' : '🪑';
                    const stockTag = item.stock_actual > 0 ? `<b>${item.stock_actual} ud.</b>` : '<span style="color:#ef4444;font-weight:bold;">Agotado</span>';
                    
                    html += `
                        <div class="autocomplete-item" onclick="seleccionarDatoConsulta('${sku}')">
                            <span>${catTag} <b>[${sku}]</b> ${item.nombre}</span>
                            <span>${stockTag}</span>
                        </div>
                    `;
                }
            }

            if (coincidencias > 0) {
                listDiv.innerHTML = html;
                listDiv.style.display = 'block';
            } else {
                listDiv.innerHTML = '<div class="autocomplete-item" style="color:#94a3b8;">Sin resultados</div>';
                listDiv.style.display = 'block';
            }
        }

        function seleccionarDatoConsulta(sku) {
            const item = inventarioGlobal[sku];
            document.getElementById('lookup-input').value = item.nombre;
            document.getElementById('autocomplete-list').style.display = 'none';
            
            if (item.categoria !== categoriaActual) {
                cambiarPestana(item.categoria);
            }

            ejecutarConsulta(sku);

            document.getElementById('v_sku').value = sku;
            document.getElementById('e_sku').value = sku;
        }

        function ejecutarConsulta(sku) {
            const display = document.getElementById('status-display');
            const title = document.getElementById('status-title');
            const desc = document.getElementById('status-desc');

            if (!sku || !inventarioGlobal[sku]) {
                display.style.display = 'none';
                return;
            }

            const item = inventarioGlobal[sku];
            const stock = item.stock_actual;

            display.style.display = 'block';
            display.className = 'status-card ';

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

        function filtrarSelect(selectId, query) {
            const select = document.getElementById(selectId);
            const term = query.toLowerCase();
            select.innerHTML = '';

            for (const [sku, item] of Object.entries(inventarioGlobal)) {
                if (item.categoria !== categoriaActual) continue;
                
                const optionText = `${sku} - ${item.nombre}`;
                if (optionText.toLowerCase().includes(term)) {
                    const opt = document.createElement('option');
                    opt.value = sku;
                    opt.innerText = optionText;
                    select.appendChild(opt);
                }
            }
            if (select.options.length > 0) {
                select.selectedIndex = 0;
            }
        }

        function cambiarPestana(cat) {
            categoriaActual = cat;
            document.getElementById('btn-tab-productos').classList.toggle('active', cat === 'productos');
            document.getElementById('btn-tab-equipamiento').classList.toggle('active', cat === 'equipamiento');
            
            if (cat === 'productos') {
                document.getElementById('tabla-titulo').innerText = '📋 Catálogo de Prendas y Productos';
                document.getElementById('form-salida-title').innerText = '🛍️ Registrar Venta';
                document.getElementById('btn-salida-action').innerText = 'Descontar Venta';
            } else {
                document.getElementById('tabla-titulo').innerText = '🪑 Equipamiento y Activos Fijos';
                document.getElementById('form-salida-title').innerText = '⚠️ Registrar Salida / Baja';
                document.getElementById('btn-salida-action').innerText = 'Descontar Equipamiento';
            }
            
            document.getElementById('v_search').value = '';
            document.getElementById('e_search').value = '';
            
            renderizarInterface();
        }

        function renderizarInterface() {
            const tbody = document.getElementById('tabla-body');
            const vSelect = document.getElementById('v_sku');
            const eSelect = document.getElementById('e_sku');
            
            tbody.innerHTML = '';
            vSelect.innerHTML = '';
            eSelect.innerHTML = '';

            let totalSkus = 0;
            let totalStock = 0;
            let sinStock = 0;

            for (const [sku, item] of Object.entries(inventarioGlobal)) {
                if (item.categoria !== categoriaActual) continue;

                totalSkus++;
                totalStock += item.stock_actual;
                if (item.stock_actual <= 0) sinStock++;

                const optionText = `${sku} - ${item.nombre}`;
                
                const optV = document.createElement('option');
                optV.value = sku;
                optV.innerText = optionText;
                vSelect.appendChild(optV);

                const optE = document.createElement('option');
                optE.value = sku;
                optE.innerText = optionText;
                eSelect.appendChild(optE);

                let badgeClass = 'badge-success';
                let estadoTexto = 'Disponible';
                
                if (item.stock_actual <= 0) {
                    badgeClass = 'badge-danger';
                    estadoTexto = 'Agotado';
                } else if (item.stock_actual <= 2) {
                    badgeClass = 'badge-warning';
                    estadoTexto = 'Stock Bajo';
                }
                
                tbody.innerHTML += `
                    <tr>
                        <td><b>${sku}</b></td>
                        <td>${item.nombre}</td>
                        <td><span class="badge ${badgeClass}">${estadoTexto}</span></td>
                        <td><b>${item.stock_actual} ud.</b></td>
                        <td><button class="row-action-btn" onclick="seleccionarDatoConsulta('${sku}')">⚡ Seleccionar</button></td>
                    </tr>
                `;
            }

            document.getElementById('kpi-total-skus').innerText = totalSkus;
            document.getElementById('kpi-total-stock').innerText = totalStock;
            document.getElementById('kpi-sin-stock').innerText = sinStock;
        }

        function filtrarTabla() {
            const query = document.getElementById('search').value.toLowerCase();
            const tbody = document.getElementById('tabla-body');
            tbody.innerHTML = '';

            for (const [sku, item] of Object.entries(inventarioGlobal)) {
                if (item.categoria !== categoriaActual) continue;

                if (sku.toLowerCase().includes(query) || item.nombre.toLowerCase().includes(query)) {
                    let badgeClass = item.stock_actual > 0 ? (item.stock_actual <= 2 ? 'badge-warning' : 'badge-success') : 'badge-danger';
                    let estadoTexto = item.stock_actual > 0 ? (item.stock_actual <= 2 ? 'Stock Bajo' : 'Disponible') : 'Agotado';
                    
                    tbody.innerHTML += `
                        <tr>
                            <td><b>${sku}</b></td>
                            <td>${item.nombre}</td>
                            <td><span class="badge ${badgeClass}">${estadoTexto}</span></td>
                            <td><b>${item.stock_actual} ud.</b></td>
                            <td><button class="row-action-btn" onclick="seleccionarDatoConsulta('${sku}')">⚡ Seleccionar</button></td>
                        </tr>
                    `;
                }
            }
        }

        async function procesarMovimiento(tipo) {
            const prefix = tipo === 'ventas' ? 'v_' : 'e_';
            const sku = document.getElementById(prefix + 'sku').value;
            const cantidad = parseInt(document.getElementById(prefix + 'cant').value);
            const registrado_por = document.getElementById(prefix + 'usuario').value;

            if (!sku) {
                mostrarToast("Por favor selecciona un producto de la lista", false);
                return;
            }

            if (!registrado_por.trim()) {
                mostrarToast("Por favor ingresa tu nombre", false);
                return;
            }

            try {
                const res = await fetch('/' + tipo, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sku, cantidad, registrado_por })
                });

                const data = await res.json();

                if (res.ok) {
                    mostrarToast(data.mensaje, true);
                    cargarInventario();
                } else {
                    mostrarToast(data.detail || "Error al registrar movimiento", false);
                }
            } catch(e) {
                mostrarToast("Error de conexión con el servidor", false);
            }
        }

        function mostrarToast(mensaje, esExito) {
            const toast = document.getElementById('toast');
            toast.innerText = mensaje;
            toast.className = esExito ? 'toast-success' : 'toast-error';
            toast.style.display = 'block';
            setTimeout(() => { toast.style.display = 'none'; }, 3500);
        }

        document.addEventListener('click', function(e) {
            if (!e.target.closest('.search-box-wrapper')) {
                document.getElementById('autocomplete-list').style.display = 'none';
            }
        });

        cargarInventario();
    </script>
</body>
</html>'''

@app.get("/stock", tags=["API"])
def ver_stock():
    resultado = {}
    for sku, item in inventario_db.items():
        stock_actual = item["stock_inicial"] + item["entradas"] - item["ventas"]
        resultado[sku] = {
            "nombre": item["nombre"],
            "categoria": item.get("categoria", "productos"),
            "stock_actual": stock_actual
        }
    return resultado

@app.post("/recargar", tags=["API"])
def recargar_excel():
    global inventario_db
    inventario_db = cargar_inventario_desde_excel()
    return {"mensaje": "✅ Datos sincronizados con el archivo Excel exitosamente."}

@app.get("/descargar-excel", tags=["API"])
def descargar_excel():
    if os.path.exists(EXCEL_PATH):
        return FileResponse(
            path=EXCEL_PATH, 
            filename="Control_Inventario_Pole_Dance.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        raise HTTPException(status_code=404, detail="El archivo Excel no se encuentra en el servidor.")

@app.post("/entradas", tags=["API"])
def registrar_entrada(mov: Movimiento):
    if mov.sku not in inventario_db:
        raise HTTPException(status_code=404, detail="El SKU no existe.")
    
    inventario_db[mov.sku]["entradas"] += mov.cantidad
    historial_entradas[mov.sku] = inventario_db[mov.sku]["entradas"]
    
    stock = inventario_db[mov.sku]["stock_inicial"] + inventario_db[mov.sku]["entradas"] - inventario_db[mov.sku]["ventas"]
    return {
        "mensaje": f"✅ Ingresadas {mov.cantidad} ud. a {inventario_db[mov.sku]['nombre']}",
        "stock_actual": stock
    }

@app.post("/ventas", tags=["API"])
def registrar_venta(mov: Movimiento):
    if mov.sku not in inventario_db:
        raise HTTPException(status_code=404, detail="El SKU no existe.")
    
    stock_actual = inventario_db[mov.sku]["stock_inicial"] + inventario_db[mov.sku]["entradas"] - inventario_db[mov.sku]["ventas"]
    
    if mov.cantidad > stock_actual:
        raise HTTPException(
            status_code=400, 
            detail=f"Stock insuficiente. Disponible: {stock_actual} ud."
        )
        
    inventario_db[mov.sku]["ventas"] += mov.cantidad
    historial_ventas[mov.sku] = inventario_db[mov.sku]["ventas"]
    
    nuevo_stock = stock_actual - mov.cantidad
    return {
        "mensaje": f"🛒 Registrada salida de {mov.cantidad} ud. de {inventario_db[mov.sku]['nombre']}",
        "stock_actual": nuevo_stock
    }

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)