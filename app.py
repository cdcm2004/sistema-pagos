import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN DE BASE DE DATOS ---
def conectar_db():
    conn = sqlite3.connect("control_pagos.db")
    return conn

def crear_tablas():
    conn = conectar_db()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS proveedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cuentas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proveedor_id INTEGER,
            factura TEXT,
            monto_total REAL,
            saldo_pendiente REAL,
            fecha_proximo_pago TEXT,
            estado TEXT DEFAULT 'Pendiente',
            FOREIGN KEY(proveedor_id) REFERENCES proveedores(id)
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS abonos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cuenta_id INTEGER,
            monto_abonado REAL,
            fecha_pago TEXT,
            FOREIGN KEY(cuenta_id) REFERENCES cuentas(id)
        )
    """)
    conn.commit()
    conn.close()

crear_tablas()

# --- INTERFAZ DE STREAMLIT ---
st.set_page_config(page_title="Control de Pagos a Proveedores", layout="wide")
st.title("📊 Sistema Rápido de Abonos y Próximos Pagos")

# Sidebar - Acciones Rápidas
st.sidebar.header("⚙️ Operaciones")
opcion = st.sidebar.radio("Ir a:", ["Ver Panel / Reportes", "Registrar Proveedor o Factura", "Registrar Abono (Pagar Poco a Poco)"])

# --- OPCIÓN 1: REPORTES Y PANEL PRINCIPAL ---
if opcion == "Ver Panel / Reportes":
    st.subheader("📆 Calendario de Próximos Pagos Pendientes")
    
    conn = conectar_db()
    query = """
        SELECT c.id AS [ID Sistema], p.nombre AS Proveedor, c.factura AS [N° Factura], 
               c.monto_total AS [Monto Original], c.saldo_pendiente AS [Saldo Pendiente], 
               c.fecha_proximo_pago AS [Próximo Pago], c.estado AS Estado
        FROM cuentas c
        JOIN proveedores p ON c.proveedor_id = p.id
        WHERE c.estado != 'Pagado'
        ORDER BY c.fecha_proximo_pago ASC
    """
    df_pendientes = pd.read_sql_query(query, conn)
    conn.close()
    
    if not df_pendientes.empty:
        st.dataframe(df_pendientes.style.highlight_max(axis=0, subset=["Saldo Pendiente"], color="#ffcccc"), use_container_width=True)
        
        # --- BOTÓN PARA EXPORTAR A EXCEL ---
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_pendientes.to_excel(writer, index=False, sheet_name='Pagos Pendientes')
            data_excel = output.getvalue()
            
            st.download_button(
                label="📥 Descargar Reporte en Excel",
                data=data_excel,
                file_name=f"Reporte_Pagos_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception:
            # Si no está instalado xlsxwriter en la nube, descarga un archivo compatible
            csv = df_pendientes.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Descargar Reporte (Formato CSV)",
                data=csv,
                file_name=f"Reporte_Pagos_{datetime.today().strftime('%Y-%m-%d')}.csv",
                mime="text/csv"
            )
        
        st.subheader("📉 Resumen de Saldos por Proveedor")
        df_grafico = df_pendientes.groupby("Proveedor")["Saldo Pendiente"].sum().reset_index()
        st.bar_chart(data=df_grafico, x="Proveedor", y="Saldo Pendiente")
    else:
        st.success("¡Al día! No hay pagos pendientes programados.")

# --- OPCIÓN 2: REGISTRAR PROVEEDOR O FACTURA ---
elif opcion == "Registrar Proveedor o Factura":
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("➕ Nuevo Proveedor")
        nuevo_prov = st.text_input("Nombre del Proveedor")
        if st.button("Guardar Proveedor"):
            if nuevo_prov:
                conn = conectar_db()
                conn.execute("INSERT INTO proveedores (nombre) VALUES (?)", (nuevo_prov,))
                conn.commit()
                conn.close()
                st.success(f"Proveedor '{nuevo_prov}' registrado.")
                st.rerun()

    with col2:
        st.subheader("🧾 Nueva Factura / Deuda")
        conn = conectar_db()
        prov_df = pd.read_sql_query("SELECT * FROM proveedores", conn)
        conn.close()
        
        if not prov_df.empty:
            prov_dict = dict(zip(prov_df['nombre'], prov_df['id']))
            prov_seleccionado = st.selectbox("Selecciona Proveedor", list(prov_dict.keys()))
            n_factura = st.text_input("Número de Factura")
            monto = st.number_input("Monto Total de la Deuda", min_value=0.0, step=10.0)
            fecha_prox = st.date_input("Fecha del Próximo Pago", min_value=datetime.today())
            
            if st.button("Guardar Factura"):
                if n_factura and monto > 0:
                    conn = conectar_db()
                    conn.execute("""
                        INSERT INTO cuentas (proveedor_id, factura, monto_total, saldo_pendiente, fecha_proximo_pago)
                        VALUES (?, ?, ?, ?, ?)
                    """, (prov_dict[prov_seleccionado], n_factura, monto, monto, str(fecha_prox)))
                    conn.commit()
                    conn.close()
                    st.success("Factura registrada exitosamente.")
        else:
            st.info("Primero debes registrar al menos un proveedor.")

# --- OPCIÓN 3: REGISTRAR ABONO ---
elif opcion == "Registrar Abono (Pagar Poco a Poco)":
    st.subheader("💸 Registrar Abono a Cuenta")
    
    conn = conectar_db()
    query_cuentas = """
        SELECT c.id, p.nombre || ' - Factura: ' || c.factura || ' (Saldo: $' || c.saldo_pendiente || ')' AS info, c.saldo_pendiente
        FROM cuentas c JOIN proveedores p ON c.proveedor_id = p.id WHERE c.estado != 'Pagado'
    """
    cuentas_df = pd.read_sql_query(query_cuentas, conn)
    conn.close()
    
    if not cuentas_df.empty:
        cuentas_dict = dict(zip(cuentas_df['info'], cuentas_df['id']))
        cuenta_sel = st.selectbox("Selecciona la cuenta a abonar", list(cuentas_dict.keys()))
        id_cuenta = cuentas_dict[cuenta_sel]
        
        saldo_actual = cuentas_df[cuentas_df['id'] == id_cuenta]['saldo_pendiente'].values
        
        monto_abono = st.number_input("Monto del Abono", min_value=0.01, max_value=float(saldo_actual), step=10.0)
        nueva_fecha = st.date_input("Siguiente Fecha de Pago (Si aún queda saldo)", min_value=datetime.today() + timedelta(days=1))
        
        if st.button("Aplicar Pago / Abono"):
            nuevo_saldo = saldo_actual - monto_abono
            nuevo_estado = "Pagado" if nuevo_saldo <= 0 else "Parcial"
            fecha_hoy = str(datetime.today().date())
            
            conn = conectar_db()
            conn.execute("INSERT INTO abonos (cuenta_id, monto_abonado, fecha_pago) VALUES (?, ?, ?)", (id_cuenta, monto_abono, fecha_hoy))
            conn.execute("""
                UPDATE cuentas 
                SET saldo_pendiente = ?, fecha_proximo_pago = ?, estado = ?
                WHERE id = ?
            """, (nuevo_saldo, str(nueva_fecha) if nuevo_saldo > 0 else "Liquidado", nuevo_estado, id_cuenta))
            conn.commit()
            conn.close()
            
            st.success(f"¡Abono aplicado! Nuevo saldo pendiente: ${nuevo_saldo:.2f}")
            st.rerun()
    else:
        st.info("No hay facturas pendientes para aplicar abonos.")
