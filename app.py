import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Bodega - ASIAM", page_icon="⚙", layout="wide"
)

# Búsqueda automática del archivo Excel en la carpeta
archivos = [f for f in os.listdir(".") if f.endswith(".xlsx")]
EXCEL_FILE = archivos[0] if archivos else "inventario.xlsx"


@st.cache_data
def cargar_datos():
  if os.path.exists(EXCEL_FILE):
    try:
      df = pd.read_excel(EXCEL_FILE, sheet_name="BUSCADOR", header=11)
      df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
      df = df.dropna(subset=["Id", "PRODUCTO"], how="all")
      df.columns = df.columns.astype(str).str.strip()
      return df
    except Exception as e:
      try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=0)
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        df.columns = df.columns.astype(str).str.strip()
        return df
      except:
        return pd.DataFrame()
  else:
    return pd.DataFrame()


df_inventario = cargar_datos()

st.title("📦 Sistema de Gestión y Cruce de Repuestos - ASIAM")
st.sidebar.header("Menú de Bodega")
opcion = st.sidebar.radio(
    "Seleccione una opción:",
    ["Inventario y Cruce de Repuestos", "Impresión Zebra (ZPL)"],
)

if opcion == "Inventario y Cruce de Repuestos":
  st.header("🔍 Consulta y Cruce Avanzado de Repuestos")

  if df_inventario.empty:
    st.error(f"⚠️ No se pudo leer el archivo '{EXCEL_FILE}'.")
  else:
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filtros de Bodega")

    if "Marca" in df_inventario.columns:
      marcas_disponibles = ["Todas"] + sorted(
          df_inventario["Marca"].dropna().astype(str).unique().tolist()
      )
      marca_sel = st.sidebar.selectbox("Filtrar por Marca", marcas_disponibles)
    else:
      marca_sel = "Todas"

    if "Categoria Primaria" in df_inventario.columns:
      categorias_disponibles = ["Todas"] + sorted(
          df_inventario["Categoria Primaria"]
          .dropna()
          .astype(str)
          .unique()
          .tolist()
      )
      cat_sel = st.sidebar.selectbox(
          "Filtrar por Categoría", categorias_disponibles
      )
    else:
      cat_sel = "Todas"

    df_filtrado = df_inventario.copy()
    if marca_sel != "Todas":
      df_filtrado = df_filtrado[
          df_filtrado["Marca"].astype(str) == marca_sel
      ]
    if cat_sel != "Todas":
      df_filtrado = df_filtrado[
          df_filtrado["Categoria Primaria"].astype(str) == cat_sel
      ]

    busqueda = st.text_input(
        "Buscar repuesto (ID, Producto, OEM, Código Asiam o Cruces):"
    )

    if busqueda:
      df_filtrado = df_filtrado[
          df_filtrado.astype(str)
          .apply(lambda row: row.str.contains(busqueda, case=False).any(), axis=1)
      ]

    st.info(f"Mostrando {len(df_filtrado)} repuestos encontrados:")
    st.dataframe(df_filtrado, use_container_width=True)

elif opcion == "Impresión Zebra (ZPL)":
  st.header("🖨️ Búsqueda y Generación de Etiquetas para Zebra (ZPL)")

  if not df_inventario.empty:
    texto_buscar = st.text_input(
        "🔎 Escribe parte del producto, código Asiam o ID para filtrar la"
        " etiqueta:"
    )

    df_etiquetas = df_inventario.copy()
    if texto_buscar:
      df_etiquetas = df_etiquetas[
          df_etiquetas.astype(str).apply(
              lambda row: row.str.contains(texto_buscar, case=False).any(), axis=1
          )
      ]

    if df_etiquetas.empty:
      st.warning("No se encontraron repuestos con ese criterio de búsqueda.")
    else:
      opciones_select = []
      for idx, row in df_etiquetas.iterrows():
        id_i = str(row.get("Id", ""))
        asiam_i = str(row.get("Código Asiam", ""))
        prod_i = str(row.get("PRODUCTO", ""))
        oem_i = str(row.get("Código Product", ""))
        etiqueta_txt = (
            f"ID: {id_i} | Asiam: {asiam_i} | {prod_i} | OEM: {oem_i}"
        )
        opciones_select.append((id_i, etiqueta_txt))

      seleccion_usuario = st.selectbox(
          "Selecciona el repuesto a etiquetar:",
          options=opciones_select,
          format_func=lambda x: x[1],
      )

      if seleccion_usuario:
        id_elegido = seleccion_usuario[0]
        item = df_inventario[
            df_inventario["Id"].astype(str) == str(id_elegido)
        ].iloc[0]

        id_val = str(item.get("Id", ""))
        marca_val = str(item.get("Marca", ""))
        prod_val = str(item.get("PRODUCTO", ""))
        oem_val = str(item.get("Código Product", ""))
        asiam_val = str(item.get("Código Asiam", ""))

        col1, col2 = st.columns(2)
        with col1:
          cantidad_copias = st.number_input(
              "Cantidad de copias a imprimir",
              min_value=1,
              max_value=50,
              value=1,
          )

        zpl_code = ""
        for _ in range(int(cantidad_copias)):
          zpl_code += f"""
^XA
^FO30,25^ADN,30,15^FDASIAM AUTOPART^FS
^FO30,65^A0N,22,22^FDID: {id_val} | Asiam: {asiam_val}^FS
^FO30,95^A0N,20,20^FDMarca: {marca_val[:28]}^FS
^FO30,120^A0N,20,20^FDDESC: {prod_val[:30]}^FS
^FO30,145^A0N,20,20^FDOEM: {oem_val}^FS
^FO30,175^BY2,2,50^BCN,50,Y,N,N^FD{id_val}^FS
^XZ
"""

        st.text_area(
            f"Código ZPL generado ({cantidad_copias} copias listas para Zebra):",
            zpl_code,
            height=200,
        )

        # --- BOTONES DE ACCIÓN PARA IMPRESIÓN ---
        col_btn1, col_btn2 = st.columns(2)

        with col_btn1:
          # Botón de descarga directa de archivo .zpl para la impresora
          st.download_button(
              label="📥 Descargar archivo ZPL para Zebra",
              data=zpl_code,
              file_name=f"etiqueta_{id_val}.zpl",
              mime="text/plain",
          )

        with col_btn2:
          if st.button("🖨️ Enviar a Impresión Rápida"):
            st.success(
                "✅ ¡Etiqueta lista! Descarga el archivo o envíalo por tu software"
                " Zebra Setup Utilities."
            )
  else:
    st.warning("No hay datos cargados para generar etiquetas.")