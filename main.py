from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
import pandas as pd
import os

app = FastAPI(
    title="Control de Inventario y Consultas Pole Dance",
    description="Sistema interactivo con registro de movimientos y consulta rápida de disponibilidad",
    version="4.0.0"
)

EXCEL_PATH = "Control_Inventario_Pole_Dance.xlsx"

def cargar_inventario_desde_excel():
    inventario = {}
    
    if os.path.exists(EXCEL_PATH):
        # 1. Cargar Productos (Ropa)
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

        # 2. Cargar Equipamiento
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

    # Ítems base de equipamiento por defecto si no existen en el Excel
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
    return """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Control de Inventario & Consultas</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            :root {
                --primary: #8e44ad;
                --primary-hover: #732d91;
                --success: #2ecc71;
                --danger: #e74c3c;
                --warning: #f39c12;
                --info: #3498db;
                --bg: #f8f9fa;
                --card-bg: #ffffff;
                --text: #2c3e50;
            }

            * { box-sizing: border-box; font-family: 'Inter', sans-serif; margin: 0; padding: 0; }
            body { background-color: var(--bg); color: var(--text); padding: 20px; }
            .container { max-width: 1050px; margin: 0 auto; }
            
            header { text-align: center; margin-bottom: 25px; }
            header h1 { color: var(--primary); font-size: 2.2rem; margin-bottom: 5px; }
            header p { color: #7f8c8d; font-size: 1rem; }

            /* Tabs navigation */
            .tab-nav { display: flex; justify-content: center; gap: 10px; margin-bottom: 25px; flex-wrap: wrap; }
            .tab-btn { padding: 12px 20px; border: none; background: #e0e0e0; border-radius: 30px; font-weight: 600; cursor: pointer; font-size: 0.95rem; transition: all 0.2s; color: #555; }
            .tab-btn.active { background: var(--primary); color: white; box-shadow: 0 4px 12px rgba(142, 68, 173, 0.3); }

            /* Consultation Box */
            .lookup-card { background: var(--card-bg); padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.06); margin-bottom: 30px; border-top: 5px solid var(--info); }
            .lookup-header { font-size: 1.2rem; font-weight: 700; color: #2c3e50; margin-bottom: 15px; display: flex; align-items: center; gap: 8px; }
            .lookup-flex { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
            .lookup-flex select { flex: 1; min-width: 250px; padding: 12px; font-size: 1rem; }
            
            .status-display { margin-top: 20px; padding: 18px; border-radius: 10px; display: none; text-align: center; }
            .status-display h3 { font-size: 1.5rem; margin-bottom: 5px; }
            .status-display p { font-size: 1.1rem; }
            .status-available { background-color: #e8f8f0; color: #1e8449; border: 1px solid #a3e4d7; }
            .status-low { background-color: #fef9e7; color: #b7950b; border: 1px solid #f9e79f; }
            .status-empty { background-color: #fdeaea; color: #922b21; border: 1px solid #f5b7b1; }

            /* KPIs Cards */
            .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 25px; }
            .kpi-card { background: var(--card-bg); padding: 18px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); text-align: center; border-left: 4px solid var(--primary); }
            .kpi-card.warning { border-left-color: var(--warning); }
            .kpi-card h4 { font-size: 0.85rem; color: #7f8c8d; text-transform: uppercase; margin-bottom: 5px; }
            .kpi-card .number { font-size: 1.6rem; font-weight: 700; color: var(--text); }

            /* Forms Layout */
            .actions-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }
            @media (max-width: 768px) { .actions-grid { grid-template-columns: 1fr; } }
            
            .card { background: var(--card-bg); padding: 22px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
            .card h3 { font-size: 1.1rem; margin-bottom: 15px; display: flex; align-items: center; gap: 8px; }
            
            label { display: block; font-size: 0.85rem; font-weight: 600; margin: 10px 0 4px; color: #34495e; }
            input, select { width: 100%; padding: 10px 12px; border: 1px solid #dcdfe6; border-radius: 8px; font-size: 0.95rem; transition: border 0.2s; }
            input:focus, select:focus { outline: none; border-color: var(--primary); }

            button { width: 100%; padding: 12px; border: none; border-radius: 8px; font-size: 0.95rem; font-weight: 600; cursor: pointer; margin-top: 15px; transition: all 0.2s; }
            .btn-venta { background: var(--danger); color: white; }
            .btn-venta:hover { background: #c0392b; }
            .btn-entrada { background: var(--success); color: white; }
            .btn-entrada:hover { background: #27ae60; }
            .btn-consultar { background: var(--info); color: white; width: auto; padding: 12px 25px; margin-top: 0; }
            .btn-consultar:hover { background: #2980b9; }

            /* Table Section */
            .table-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; flex-wrap: wrap; gap: 10px; }
            .search-box { width: 280px; }
            
            .table-container { background: var(--card-bg); border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); overflow-x: auto; }
            table { width: 100%; border-collapse: collapse; text-align: left; }
            th, td { padding: 14px 18px; border-bottom: 1px solid #f1f2f6; font-size: 0.92rem; }
            th { background: #fafbfc; color: #57606f; font-weight: 600; }
            
            .badge { padding: 5px 10px; border-radius: 20px; font-weight: 600; font-size: 0.8rem; display: inline-block; }
            .badge-success { background: #e8f8f0; color: #2ecc71; }
            .badge-warning { background: #fef9e7; color: #f39c12; }
            .badge-danger { background: #fdeaea; color: #e74c3c; }

            /* Toast Notifications */
            #toast { position: fixed; bottom: 20px; right: 20px; padding: 14px 20px; border-radius: 8px; color: white; font-weight: 500; display: none; box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 1000; }
            .toast-success { background: var(--success); }
            .toast-error { background: var(--danger); }
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <h1>✨ Inventario & Consultas Pole Dance</h1>
                <p>Gestión de Prendas, Mobiliario y Consultas de Disponibilidad</p>
            </header>

            <!-- 🔍 SECCIÓN DE CONSULTA RÁPIDA -->
            <div class="lookup-card">
                <div class="lookup-header">🔍 Consulta Rápida de Disponibilidad ("¿Se tiene?")</div>
                <div class="lookup-flex">
                    <select id="lookup-sku" onchange="ejecutarConsulta()">
                        <option value="">-- Selecciona una prenda o equipo para consultar --</option>
                    </select>
                </div>

                <div id="status-display" class="status-display">
                    <h3 id="status-title">---</h3>
                    <p id="status-desc">---</p>
                </div>
            </div>

            <!-- Navegación de Pestañas -->
            <div class="tab-nav">
                <button class="tab-btn active" id="btn-tab-productos" onclick="cambiarPestana('productos')">👗 Ropa y Productos</button>
                <button class="tab-btn" id="btn-tab-equipamiento" onclick="cambiarPestana('equipamiento')">🪑 Equipamiento y Activos</button>
            </div>

            <!-- KPIs -->
            <div class="kpi-grid">
                <div class="kpi-card">
                    <h4>Variedad de Ítems</h4>
                    <div class="number" id="kpi-total-skus">0</div>
                </div>
                <div class="kpi-card">
                    <h4>Stock Físico Total</h4>
                    <div class="number" id="kpi-total-stock">0</div>
                </div>
                <div class="kpi-card warning">
                    <h4>Agotados / Sin Stock</h4>
                    <div class="number" id="kpi-sin-stock">0</div>
                </div>
            </div>

            <!-- Formularios de Entradas / Salidas -->
            <div class="actions-grid">
                <div class="card">
                    <h3 id="form-salida-title">🛍️ Registrar Venta / Salida</h3>
                    <label>Seleccionar Ítem:</label>
                    <select id="v_sku"></select>
                    
                    <label>Cantidad a Descontar:</label>
                    <input type="number" id="v_cant" value="1" min="1">
                    
                    <label>Registrado por:</label>
                    <input type="text" id="v_usuario" placeholder="Ej: María">
                    
                    <button class="btn-venta" id="btn-salida-action" onclick="procesarMovimiento('ventas')">Descontar Unidad</button>
                </div>

                <div class="card">
                    <h3 id="form-entrada-title">📦 Registrar Entrada / Compra</h3>
                    <label>Seleccionar Ítem:</label>
                    <select id="e_sku"></select>
                    
                    <label>Cantidad Ingresada:</label>
                    <input type="number" id="e_cant" value="1" min="1">
                    
                    <label>Registrado por:</label>
                    <input type="text" id="e_usuario" placeholder="Ej: Admin">
                    
                    <button class="btn-entrada" onclick="procesarMovimiento('entradas')">Ingresar Stock</button>
                </div>
            </div>

            <!-- Tabla Stock -->
            <div class="table-header">
                <h2 id="tabla-titulo">📋 Catálogo de Productos</h2>
                <input type="text" id="search" class="search-box" placeholder="🔍 Buscar por nombre o SKU..." onkeyup="filtrarTabla()">
            </div>

            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>SKU</th>
                            <th>Descripción</th>
                            <th>Categoría</th>
                            <th>Estado</th>
                            <th>Stock Disponible</th>
                        </tr>
                    </thead>
                    <tbody id="tabla-body">
                        <tr><td colspan="5" style="text-align:center;">Cargando datos...</td></tr>
                    </tbody>
                </table>
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
                    mostrarToast("Error conectando con el servidor", false);
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
                display.className = 'status-display ';

                if (stock > 2) {
                    display.classList.add('status-available');
                    title.innerHTML = '🟢 SÍ HAY DISPONIBILIDAD';
                    desc.innerHTML = `Quedan <b>${stock} unidades</b> disponibles de <i>${item.nombre}</i> (${sku}).`;
                } else if (stock > 0) {
                    display.classList.add('status-low');
                    title.innerHTML = '⚠️ ÚLTIMAS UNIDADES';
                    desc.innerHTML = `¡Atención! Solo quedan <b>${stock} unidad(es)</b> de <i>${item.nombre}</i> (${sku}).`;
                } else {
                    display.classList.add('status-empty');
                    title.innerHTML = '🔴 NO HAY STOCK (AGOTADO)';
                    desc.innerHTML = `Actualmente hay <b>0 unidades</b> de <i>${item.nombre}</i> (${sku}).`;
                }
            }

            function cambiarPestana(cat) {
                categoriaActual = cat;
                document.getElementById('btn-tab-productos').classList.toggle('active', cat === 'productos');
                document.getElementById('btn-tab-equipamiento').classList.toggle('active', cat === 'equipamiento');
                
                if (cat === 'productos') {
                    document.getElementById('tabla-titulo').innerText = '📋 Catálogo de Productos (Ropa)';
                    document.getElementById('form-salida-title').innerText = '🛍️ Registrar Venta';
                    document.getElementById('btn-salida-action').innerText = 'Descontar Venta';
                } else {
                    document.getElementById('tabla-titulo').innerText = '🪑 Equipamiento y Mobiliario';
                    document.getElementById('form-salida-title').innerText = '⚠️ Registrar Salida / Baja / Daño';
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

                    const catTag = item.categoria === 'productos' ? 'Ropa' : 'Equipamiento';
                    
                    tbody.innerHTML += `
                        <tr>
                            <td><b>${sku}</b></td>
                            <td>${item.nombre}</td>
                            <td><small>${catTag}</small></td>
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
                        const catTag = item.categoria === 'productos' ? 'Ropa' : 'Equipamiento';
                        
                        tbody.innerHTML += `
                            <tr>
                                <td><b>${sku}</b></td>
                                <td>${item.nombre}</td>
                                <td><small>${catTag}</small></td>
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
    </html>
    """

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