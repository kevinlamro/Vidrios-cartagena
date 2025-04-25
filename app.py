# Importo las librerías necesarias
import streamlit as st
import pandas as pd
import altair as alt
import io

# Configuración de la app: pantalla ancha y título
st.set_page_config(layout="wide")
st.title("Dashboard Completo - Vidrios Cartagena 2024")

st.subheader("Analisis Reslizado por : Kevin Lambraño, Dalmiro Barrios, Juan Madera, Eliot Fuentes")
st.write("Nota: en la parte izquierda puede desplegar la tabla y todos los graficos")

# Nombre del archivo de Excel que vamos a analizar
archivo = "Dashboard_Ventas2024-Act.xlsx"

# Uso try para manejar si el archivo no se encuentra
try:
    # Leo los datos desde el Excel
    df = pd.read_excel(archivo, sheet_name="Datos de Ventas 2024")

    # Convierto columnas clave a números, por si hay errores
    for col in ['Total Ventas (COP)', 'Cantidad de Ventas', 'Clientes Nuevos', 'Clientes Recurrentes']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Creo nuevos indicadores que voy a usar en los gráficos
    df["Ingreso Promedio por Venta"] = df["Total Ventas (COP)"] / df["Cantidad de Ventas"]
    df["% Retención"] = df["Clientes Recurrentes"] / (df["Clientes Nuevos"] + df["Clientes Recurrentes"])
    df["Conversión a Recurrentes"] = df["Clientes Recurrentes"] / df["Clientes Nuevos"]

    # Ordeno los meses para que los gráficos salgan bien
    orden_meses = ["Enero 2024", "Febrero 2024", "Marzo 2024", "Abril 2024", "Mayo 2024",
                "Junio 2024", "Julio 2024", "Agosto 2024", "Septiembre 2024", "Octubre 2024",
                "Noviembre 2024", "Diciembre 2024"]
    df["Mes"] = pd.Categorical(df["Mes"], categories=orden_meses, ordered=True)
    df = df.sort_values("Mes")

    # En el sidebar elijo qué gráficos quiero ver
    opciones = st.sidebar.multiselect(
        "Selecciona qué análisis mostrar:",
        ["Ventas por Mes", "Ingreso Promedio", "Clientes Nuevos vs Recurrentes", 
        "Tasa de Retención", "Conversión a Recurrentes", "Ranking Producto Más Vendido", 
        "Ventas Totales por Producto", "Top 3 Meses con Mayores Ventas", "Tabla Completa"],
        default=["Ventas por Mes", "Ranking Producto Más Vendido"]
    )

    # Gráfico de ventas por mes
    if "Ventas por Mes" in opciones:
        st.subheader("Total de Ventas por Mes (COP)")
        st.altair_chart(alt.Chart(df).mark_bar().encode(
            x='Mes', y='Total Ventas (COP)', color='Mes'
        ).properties(width=700), use_container_width=True)

    # Gráfico de ingreso promedio por venta
    if "Ingreso Promedio" in opciones:
        st.subheader("Ingreso Promedio por Venta")
        st.altair_chart(alt.Chart(df).mark_line(point=True).encode(
            x='Mes', y='Ingreso Promedio por Venta'
        ).properties(width=700), use_container_width=True)

    # Comparo clientes nuevos y recurrentes
    if "Clientes Nuevos vs Recurrentes" in opciones:
        st.subheader("Clientes Nuevos vs Recurrentes")
        df_melt = df.melt(id_vars=["Mes"], value_vars=["Clientes Nuevos", "Clientes Recurrentes"],
                        var_name="Tipo Cliente", value_name="Cantidad")
        st.altair_chart(alt.Chart(df_melt).mark_bar().encode(
            x='Mes', y='Cantidad', color='Tipo Cliente'
        ).properties(width=700), use_container_width=True)

    # Tasa de retención
    if "Tasa de Retención" in opciones:
        st.subheader("Tasa de Retención de Clientes (%)")
        st.altair_chart(alt.Chart(df).mark_line(point=True).encode(
            x='Mes', y=alt.Y('% Retención', axis=alt.Axis(format='.0%'))
        ).properties(width=700), use_container_width=True)

    # Conversión de nuevos a recurrentes
    if "Conversión a Recurrentes" in opciones:
        st.subheader("Conversión de Nuevos a Recurrentes")
        st.altair_chart(alt.Chart(df).mark_area(opacity=0.6).encode(
            x='Mes', y='Conversión a Recurrentes'
        ).properties(width=700), use_container_width=True)

    # Producto más vendido
    if "Ranking Producto Más Vendido" in opciones:
        st.subheader("Producto Más Vendido (Frecuencia)")
        ranking = df["Producto Más Vendido"].value_counts().reset_index()
        ranking.columns = ["Producto", "Frecuencia"]
        st.bar_chart(ranking.set_index("Producto"))

    # Ventas totales por producto más vendido
    if "Ventas Totales por Producto" in opciones:
        st.subheader("Ventas Totales por Producto Más Vendido")
        ventas_totales = df.groupby("Producto Más Vendido")["Total Ventas (COP)"].sum().reset_index()
        st.altair_chart(alt.Chart(ventas_totales).mark_bar().encode(
            x='Producto Más Vendido', y='Total Ventas (COP)', color='Producto Más Vendido'
        ).properties(width=700), use_container_width=True)

    # Tabla con los 3 meses de mejores ventas
    if "Top 3 Meses con Mayores Ventas" in opciones:
        st.subheader("Top 3 Meses con Más Ventas")
        top3 = df.sort_values(by="Total Ventas (COP)", ascending=False).head(3)
        st.table(top3[["Mes", "Total Ventas (COP)"]])

    # Muestro toda la tabla con los indicadores
    if "Tabla Completa" in opciones:
        st.subheader("Tabla Completa con KPIs Calculados")
        tabla_formateada = df.copy()
        tabla_formateada["% Retención"] = (tabla_formateada["% Retención"] * 100).round(2).astype(str) + " %"
        tabla_formateada["Conversión a Recurrentes"] = (tabla_formateada["Conversión a Recurrentes"] * 100).round(2).astype(str) + " %"
        st.dataframe(tabla_formateada)

    # EXPORTACIÓN
    st.markdown("---")
    st.subheader("Exportar Archivos")

    # Creo un Excel nuevo con los datos ya procesados
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False, sheet_name="Datos de Ventas 2024")
    excel_buffer.seek(0)

    # Botón para descargar el Excel actualizado
    st.download_button(
        label="Descargar Excel con los datos",
        data=excel_buffer,
        file_name="Dashboard_Ventas2024-Act.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Botón para descargar el informe en PDF
    with open("InformeVidriosCartagena2024.pdf", "rb") as pdf_file:
        pdf_data = pdf_file.read()

    st.download_button(
        label="Descargar Informe en PDF",
        data=pdf_data,
        file_name="InformeVidriosCartagena2024.pdf",
        mime="application/pdf"
    )

# En caso de que el archivo Excel no se encuentre
except FileNotFoundError:
    st.error(f"El archivo '{archivo}' no se encuentra en la carpeta.")
