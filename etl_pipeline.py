# ==============================================================================
# PIPELINE ETL AUTOMATIZADO - PROYECTO MATEANDO DATOS
# ==============================================================================

import pandas as pd
from sqlalchemy import create_engine

print("🚀 [INICIO] Arrancando el proceso ETL automatizado...")

# ------------------------------------------------------------------------------
# 1. CONEXIÓN AL MOTOR DE BASE DE DATOS (Infraestructura base)
# ------------------------------------------------------------------------------
# Aquí establecemos el puente con nuestra base de datos SQLite local.
engine = create_engine('sqlite:///mateando_datos.db')


# ------------------------------------------------------------------------------
# 2. FASE DE TRANSFORMACIÓN Y CONSULTAS (El corazón analítico)
# ------------------------------------------------------------------------------
# Definimos las consultas SQL avanzadas que procesan los datos crudos 
# aplicando los JOINs y agregaciones que fuimos aprendiendo.

query_ventas_diarias = """
    SELECT 
        fecha,
        COUNT(DISTINCT pedido) AS total_pedidos,
        SUM(cantidad) AS total_unidades_vendidas,
        SUM(total_venta) AS facturacion_total,
        AVG(total_venta) AS ticket_promedio
    FROM ventas
    WHERE estado = 'Pagada'
    GROUP BY fecha
    ORDER BY fecha ASC;
"""

query_ranking_productos = """
    SELECT sku, SUM(cantidad) AS total_unidades, SUM(total_venta) AS facturacion_total
    FROM ventas
    WHERE estado = 'Pagada'
    GROUP BY sku
    ORDER BY total_unidades DESC;
"""

query_reporte_completo = """
    SELECT 
        c.nombre_cliente,
        c.provincia,
        mp.nombre_metodo AS metodo_pago,
        COUNT(v.pedido) AS total_pedidos,
        SUM(v.total_venta) AS facturacion_total
    FROM ventas v
    JOIN clientes c ON v.id_cliente = c.id_cliente
    JOIN metodos_pago mp ON v.id_metodo = mp.id_metodo
    WHERE v.estado = 'Pagada'
    GROUP BY c.nombre_cliente, c.provincia, mp.nombre_metodo
    ORDER BY facturacion_total DESC;
"""


# ------------------------------------------------------------------------------
# 3. FASE DE EXTRACCIÓN Y CARGA (ETL en acción)
# ------------------------------------------------------------------------------
# Ejecutamos las consultas contra SQLite y traemos los datos limpios a DataFrames
print("⚙️ Procesando y transformando datos mediante consultas SQL...")
df_diarias = pd.read_sql(query_ventas_diarias, con=engine)
df_ranking = pd.read_sql(query_ranking_productos, con=engine)
df_completo = pd.read_sql(query_reporte_completo, con=engine)


# ------------------------------------------------------------------------------
# 4. FASE DE SALIDA / REPORTE FINAL (Load a Excel para Looker Studio)
# ------------------------------------------------------------------------------
# Empaquetamos automáticamente todos los dataframes en un archivo Excel multipestaña

df_estados = pd.read_sql("SELECT * FROM resumen_estados", con=engine)

nombre_archivo = "reporte_mateando_datos.xlsx"

print(f"💾 Guardando resultados limpios en el archivo: {nombre_archivo}...")
with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
    df_diarias.to_excel(writer, sheet_name='Ventas Diarias', index=False)
    df_ranking.to_excel(writer, sheet_name='Ranking Productos', index=False)
    df_estados.to_excel(writer, sheet_name='Resumen Estados', index=False)
    df_completo.to_excel(writer, sheet_name='Reporte Clientes y Pagos', index=False)

print("✅ [FIN] ¡Pipeline ETL ejecutado con éxito! Reporte listo.")