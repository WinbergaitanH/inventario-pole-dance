from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
import pandas as pd
import os

app = FastAPI(
    title="Pole Dance Rojas Sport - Inventario & Activos",
    description="Sistema de control de inventario y disponibilidad para Pole Dance Rojas Sport",
    version="5.0.0"
)

EXCEL_PATH = "Control_Inventario_Pole_Dance.xlsx"

def cargar_inventario_desde_excel():
    inventario = {}
    
    if os.path.exists(EXCEL_PATH):
        # 1. Cargar Productos (Ropa / Accesorios)
        try:
            df_catalogo = pd.read_excel(EXCEL_PATH, sheet_name='🏷️ Catálogo Productos', header=1)
            for index, row in df_catalogo.iterrows():
                sku = str(row.iloc[0]).strip()
                if pd.notna(sku) and sku.lower() not in ["nan", "sku", "código", "codigo"]:
                    nombre = f"{row.iloc[2]} ({row.iloc[1]}) - Talla {row.iloc[3]}"
                    try:
                        stock_ini = int(row.iloc[6])
                    except (ValueError, TypeError):
                        stock_ini = 0

                    inventario[sku] = {
                        "nombre": nombre,
                        "categoria": "productos",
                        "stock_inicial": stock_ini,
                        "entradas": 0,
                        "ventas": 0
                    }
        except Exception:
            pass

        # 2. Cargar Equipamiento / Activos
        try:
            df_equip = pd.read_excel(EXCEL_PATH, sheet_name='🪑 Equipamiento', header=0)
            for index, row in df_equip.iterrows():
                sku = str(row.iloc[0]).strip()
                if pd.notna(sku) and sku.lower() not in ["nan", "sku", "código", "codigo"]:
                    nombre = str(row.iloc[1]).strip()
                    try:
                        stock_ini = int(row.iloc[2])
                    except (ValueError, TypeError):
                        stock_ini = 0

                    inventario[sku] = {
                        "nombre": nombre,
                        "categoria": "equipamiento",
                        "stock_inicial": stock_ini,
                        "entradas": 0,
                        "ventas": 0
                    }
        except Exception:
            pass

    # Equipamiento base por defecto
    items_base_equipamiento = {
        "EQ-TUBO": "Tubo Pole Dance Profesional",
        "EQ-MAT": "Mat / Colchoneta de Caída",
        "EQ-SILLA": "Silla de Entrenamiento",
        "EQ-MESA": "Mesa Auxiliar",
        "EQ-ESPEJO": "Espejo de Sala"
    }

    for sku, nombre in items_base_equipamiento.items():
        if sku not in inventario:
            inventario[sku] = {
                "nombre": nombre,
                "categoria": "equipamiento",
                "stock_inicial": 5,
                "entradas": 0,
                "ventas": 0
            }

    return inventario

inventario_db = cargar_inventario_desde_excel()

class Movimiento(BaseModel):
    sku: str = Field(..., description="Código SKU del elemento")
    cantidad: int = Field(..., description="Cantidad de unidades", gt=0)
    registrado_por: str = Field(..., description="Nombre del responsable")

@app.get("/", response_class=HTMLResponse, tags=["Interfaz"])
def interfaz_usuario():
    return """<!DOCTYPE html>
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

        /* HEADER BRANDING */
        header {
            text-align: center;
            padding: 25px 20px 30px;
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
        }

        /* CARD CONTAINERS */
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

        /* BUSCADOR DE DISPONIBILIDAD */
        .lookup-box select {
            width: 100%;
            padding: 14px 16px;
            border: 2px solid #e2e8f0;
            border-radius: var(--radius-lg);
            font-size: 1rem;
            font-weight: 500;
            color: var(--text-main);
            background-color: #f8fafc;
            transition: all 0.25s ease;
            outline: none;
        }

        .lookup-box select:focus {
            border-color: var(--primary);
            background-color: #fff;
            box-shadow: 0 0 0 4px var(--accent-glow);
        }

        .status-card {
            margin-top: 18px;
            padding: 18px 20px;
            border-radius: var(--radius-lg);
            display: none;
            text-align: center;
            animation: fadeIn 0.3s ease-in-out;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-5px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .status-card h3 { font-size: 1.3rem; font-weight: 800; margin-bottom: 4px; }
        .status-card p { font-size: 0.95rem; font-weight: 500; }

        .status-available { background: var(--success-bg); color: #047857; border: 1.5px solid #a7f3d0; }
        .status-low { background: var(--warning-bg); color: #b45309; border: 1.5px solid #fde68a; }
        .status-empty { background: var(--danger-bg); color: #b91c1c; border: 1.5px solid #fca5a5; }

        /* TABS */
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
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .tab-btn.active {
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%);
            color: white;
            box-shadow: 0 4px 15px var(--accent-glow);
        }

        /* KPIS GRID */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
            margin-bottom: 25px;
        }

        @media (max-width: 640px) {
            .kpi-grid { grid-template-columns: 1fr; }
        }

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

        .kpi-card .number {
            font-size: 1.8rem;
            font-weight: 800;
            color: var(--text-main);
        }

        .kpi-card.alert .number { color: var(--danger); }

        /* FORM GRID */
        .form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 25px;
        }

        @media (max-width: 768px) {
            .form-grid { grid-template-columns: 1fr; }
        }

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
            transition: border 0.2s;
            background: #f8fafc;
        }

        .form-group input:focus, .form-group select:focus {
            border-color: var(--primary);
            background: #fff;
        }

        .btn-action {
            width: 100%;
            padding: 13px;
            border: none;
            border-radius: var(--radius-md);
            font-weight: 700;
            font-size: 0.95rem;
            cursor: pointer;
            transition: all 0.25s ease;
            color: white;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            margin-top: 6px;
        }

        .btn-salida { background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); }
        .btn-salida:hover { box-shadow: 0 6px 18px rgba(239, 68, 68, 0.4); transform: translateY(-1px); }

        .btn-entrada { background: linear-gradient(135deg, #10b981 0%, #059669 100%); }
        .btn-entrada:hover { box-shadow: 0 6px 18px rgba(16, 185, 129, 0.4); transform: translateY(-1px); }

        /* TABLA DE PRODUCTOS */
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

        .search-input:focus { border-color: var(--primary); background: white; }

        .table-wrapper {
            overflow-x: auto;
            border-radius: var(--radius-lg);
            border: 1px solid #e2e8f0;
        }

        table { width: 100%; border-collapse: collapse; background: white; text-align: left; }
        th { background: #f8fafc; color: #475569; font-size: 0.8rem; text-transform: uppercase; font-weight: 700; padding: 14px 16px; border-bottom: 1px solid #e2e8f0; }
        td { padding: 14px 16px; border-bottom: 1px solid #f1f5f9; font-size: 0.92rem; color: #334155; }
        tr:last-child td { border-bottom: none; }

        .badge {
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 0.78rem;
            font-weight: 700;
            display: inline-block;
        }
        .badge-success { background: var(--success-bg); color: #047857; }
        .badge-warning { background: var(--warning-bg); color: #b45309; }
        .badge-danger { background: var(--danger-bg); color: #b91c1c; }

        /* TOAST NOTIFICATION */
        #toast {
            position: fixed;
            bottom: 25px;
            right: 25px;
            padding: 14px 22px;
            border-radius: var(--radius-lg);
            color: white;
            font-weight: 600;
            font-size: 0.95rem;
            display: none;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.25);
            z-index: 1000;
            animation: slideUp 0.3s ease-out;
        }
        @keyframes slideUp {
            from { transform: translateY(20px); opacity: 0; }
            to { transform: translateY(0); opacity: 1; }
        }
        .toast-success { background: #10b981; }
        .toast-error { background: #ef4444; }
    </style>
</head>
<body>
    <div class="container">
        <!-- MARCA & TITULO -->
        <header>
            <div class="brand-logo">✨ Pole Dance Rojas Sport</div>
            <div class="brand-subtitle">SISTEMA INTEGRAL DE INVENTARIO Y DISPONIBILIDAD</div>
        </header>

        <!-- BUSCADOR DE DISPONIBILIDAD ("¿SE TIENE?") -->
        <div class="glass-card lookup-box">
            <div class="card-title">🔍 Consultar Disponibilidad ("¿Se tiene?")</div>
            <select id="lookup-sku" onchange="ejecutarConsulta()">
                <option value="">-- Selecciona una prenda o equipo para consultar --</option>
            </select>
            <div id="status-display" class="status-card">
                <h3 id="status-title">---</h3>
                <p id="status-desc">---</p>
            </div>
        </div>

        <!-- PESTAÑAS (ROPA / EQUIPOS) -->
        <div class="tab-group">
            <button class="tab-btn active" id="btn-tab-productos" onclick="cambiarPestana('productos')">👗 Ropa & Productos</button>
            <button class="tab-btn" id="btn-tab-equipamiento" onclick="cambiarPestana('equipamiento')">🪑 Equipamiento & Activos</button>
        </div>

        <!-- KPIS DETALLADOS -->
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

        <!-- FORMULARIOS DE REGISTRO DE ENTRADAS Y SALIDAS -->
        <div class="form-grid">
            <div class="glass-card">
                <div class="card-title" id="form-salida-title">🛍️ Registrar Venta / Salida</div>
                <div class="form-group">
                    <label>Seleccionar Ítem</label>
                    <select id="v_sku"></select>
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
                    <label>Seleccionar Ítem</label>
                    <select id="e_sku"></select>
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

        <!-- TABLA PRINCIPAL DE INVENTARIO -->
        <div class="glass-card">
            <div class="table-header">
                <div class="card-title" id="tabla-titulo" style="margin-bottom:0;">📋 Listado de Productos</div>
                <input type="text" id="search" class="search-input" placeholder="🔍 Buscar por nombre o SKU..." onkeyup="filtrarTabla()">
            </div>
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>SKU</th>
                            <th>Descripción / Producto</th>
                            <th>Estado</th>
                            <th>Disponibles</th>
                        </tr>
                    </thead>
                    <tbody id="tabla-body">
                        <tr><td colspan="4" style="text-align:center;">Cargando inventario...</td></tr>
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
                actualizarListaConsulta();
                renderizarInterface();
            } catch(e) {
                mostrarToast("Error de conexión con el servidor", false);
            }
        }

        function actualizarListaConsulta() {
            const lookupSelect = document.getElementById('lookup-sku');
            const valorPrevio = lookupSelect.value;
            lookupSelect.innerHTML = '<option value="">-- Selecciona una prenda o equipo para consultar --</option>';

            for (const [sku, item] of Object.entries(inventarioGlobal)) {
                const catTag = item.categoria === 'productos' ? '👗' : '🪑';
                lookupSelect.innerHTML += `<option value="${sku}">${catTag} [${sku}] ${item.nombre}</option>`;
            }
            if (valorPrevio) {
                lookupSelect.value = valorPrevio;
                ejecutarConsulta();
            }
        }

        function ejecutarConsulta() {
            const sku = document.getElementById('lookup-sku').value;
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
                vSelect.innerHTML += `<option value="${sku}">${optionText}</option>`;
                eSelect.innerHTML += `<option value="${sku}">${optionText}</option>`;

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

        cargarInventario();
    </script>
</body>
</html>"""

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

@app.post("/entradas", tags=["API"])
def registrar_entrada(mov: Movimiento):
    if mov.sku not in inventario_db:
        raise HTTPException(status_code=404, detail="El SKU no existe.")
    
    inventario_db[mov.sku]["entradas"] += mov.cantidad
    stock = inventario_db[mov.sku]["stock_inicial"] + inventario_db[mov.sku]["entradas"] - inventario_db[mov.sku]["ventas"]
    return {
        "mensaje": f"✅ Ingresadas {mov.cantidad} ud. a {inventario_db[mov.sku]['nombre']}",
        "stock_actual": stock
    }

@app.post("/ventas", tags=["API"])
def registrar_venta(mov: Movimiento):
    if mov.sku not in inventario_db:
        raise HTTPException(status_code=404, detail="El SKU no existe.")
    
    item = inventario_db[mov.sku]
    stock_disponible = item["stock_inicial"] + item["entradas"] - item["ventas"]
    
    if mov.cantidad > stock_disponible:
        raise HTTPException(status_code=400, detail=f"⚠️ Stock insuficiente. Solo quedan {stock_disponible} unidades.")
    
    item["ventas"] += mov.cantidad
    nuevo_stock = stock_disponible - mov.cantidad
    return {
        "mensaje": f"Descuento registrado por {mov.registrado_por}",
        "stock_restante": nuevo_stock
    }