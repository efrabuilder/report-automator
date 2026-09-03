# app_streamlit.py
#
# Interfaz web (Streamlit) para el Inventory Automator - Manejo de Datos (EDA).
# Ejecutar con: streamlit run app_streamlit.py

import streamlit as st
import pandas as pd
import tempfile
import os

import core.analisis as analisis

st.set_page_config(page_title="Inventory Automator - Manejo de Datos (EDA)", layout="wide")



# Estado de sesion: un diccionario "estado" persistente entre
# interacciones. Se guarda en st.session_state (y no en una variable
# de modulo) porque Streamlit puede atender a varios usuarios distintos
# dentro del mismo proceso: cada uno necesita su propio df cargado.

if "estado" not in st.session_state:
    st.session_state.estado = analisis.nuevo_estado()
estado = st.session_state.estado



# Barra lateral: carga y navegacion (equivalente al menu principal)

st.sidebar.title("Menu principal")

archivo = st.sidebar.file_uploader("Cargar archivo CSV o Excel", type=["csv", "xlsx", "xls"])
if archivo is not None:
    ruta_temporal = os.path.join(tempfile.gettempdir(), archivo.name)
    with open(ruta_temporal, "wb") as f:
        f.write(archivo.getbuffer())

    extension = os.path.splitext(ruta_temporal)[1].lower()

    # Si es un Excel con varias hojas, se muestra un selector ANTES de
    # cargar los datos, para que el usuario elija cual hoja analizar
    # (Streamlit vuelve a ejecutar el script completo en cada seleccion,
    # asi que este bloque se repite hasta que "hoja_elegida" quede fija).
    hoja_elegida = None
    if extension in analisis.EXTENSIONES_EXCEL:
        try:
            hojas = analisis.listar_hojas_excel(ruta_temporal)
        except Exception as e:
            hojas = []
            st.sidebar.error(str(e))
        if len(hojas) > 1:
            hoja_elegida = st.sidebar.selectbox("Elegir hoja a analizar", hojas)
        elif hojas:
            hoja_elegida = hojas[0]

    archivo_nuevo = estado["ruta_actual"] != ruta_temporal
    hoja_distinta = estado["hoja_actual"] != hoja_elegida
    if archivo_nuevo or hoja_distinta:
        try:
            mensaje = analisis.cargar_archivo(estado, ruta_temporal, hoja=hoja_elegida)
            estado["hoja_actual"] = hoja_elegida
            st.sidebar.success(mensaje)
        except Exception as e:
            st.sidebar.error(str(e))

if estado["df"] is not None:
    hoja_info = f" (hoja '{estado['hoja_actual']}')" if estado["hoja_actual"] else ""
    st.sidebar.info(f"Archivo: {archivo.name if archivo else estado['ruta_actual']}{hoja_info}\n\n"
                     f"{estado['df'].shape[0]} filas, {estado['df'].shape[1]} columnas")
    if st.sidebar.button("Limpiar tabla (quitar espacios y duplicados)"):
        st.sidebar.success(analisis.limpiar_datos(estado))

    if st.sidebar.button("Normalizar categoria/vendedor"):
        msg1 = analisis.normalizar_categoria_texto(estado, "categoria")
        msg2 = analisis.normalizar_categoria_texto(estado, "vendedor")
        st.sidebar.success("Columnas 'categoria' y 'vendedor' normalizadas.")

    if st.sidebar.button("Normalizar cliente_frecuente (Si/No)"):
        msg3 = analisis.normalizar_categoria_texto(
            estado, "cliente_frecuente",
            mapa_valores={
                "Si": ["si", "Si", "SI", "sí", "Sí", "SÍ"],
                "No": ["no", "No", "NO"],
            },
        )
        st.sidebar.success(msg3)
else:
    st.sidebar.warning("Sin archivo cargado")

opcion = st.sidebar.radio("Ir a:", [
    "Informacion del conjunto de datos",
    "Primeras y ultimas filas",
    "Tipos de datos",
    "Valores nulos",
    "Datos duplicados",
    "Estadisticas descriptivas",
    "Filtrar o consultar datos",
    "Agrupaciones y operaciones",
    "Correlacion entre variables",
    "Representaciones graficas",
    "Analisis adicional (inconsistencias)",
    "Valores negativos",
    "Ingresar nuevos datos",
])

st.title("Inventory Automator - Manejo de Datos (EDA)")



# Cada seccion revisa primero si hay datos cargados

def requiere_datos():
    if estado["df"] is None:
        st.warning("Debe cargar un archivo CSV o Excel antes de realizar esta operacion.")
        return False
    return True


if opcion == "Informacion del conjunto de datos":
    if requiere_datos():
        st.text(analisis.info_general(estado)["texto"])

elif opcion == "Primeras y ultimas filas":
    if requiere_datos():
        n = st.slider("Cantidad de filas", 1, 20, 5)
        st.subheader("Primeras filas")
        st.dataframe(analisis.primeras_filas(estado, n))
        st.subheader("Ultimas filas")
        st.dataframe(analisis.ultimas_filas(estado, n))

elif opcion == "Tipos de datos":
    if requiere_datos():
        st.dataframe(analisis.tipos_datos(estado))

elif opcion == "Valores nulos":
    if requiere_datos():
        nulos = analisis.valores_nulos(estado)
        if nulos.empty:
            st.success("No hay valores nulos en el conjunto de datos.")
        else:
            st.dataframe(nulos)

elif opcion == "Datos duplicados":
    if requiere_datos():
        dup = analisis.datos_duplicados(estado)
        if dup.empty:
            st.success("No hay filas duplicadas.")
        else:
            st.warning(f"Se encontraron {dup.shape[0]} filas duplicadas.")
            st.dataframe(dup)

elif opcion == "Estadisticas descriptivas":
    if requiere_datos():
        st.dataframe(analisis.estadisticas_descriptivas(estado))

elif opcion == "Filtrar o consultar datos":
    if requiere_datos():
        col1, col2, col3 = st.columns(3)
        columna = col1.selectbox("Columna", estado["df"].columns)
        operador = col2.selectbox("Operador", ["==", "!=", ">", ">=", "<", "<=", "contiene"])
        valor = col3.text_input("Valor")
        if st.button("Filtrar") and valor != "":
            try:
                resultado = analisis.filtrar(estado, columna, operador, valor)
                st.write(f"{resultado.shape[0]} fila(s) encontradas")
                st.dataframe(resultado)
            except Exception as e:
                st.error(str(e))

elif opcion == "Agrupaciones y operaciones":
    if requiere_datos():
        col1, col2, col3 = st.columns(3)
        columna_grupo = col1.selectbox("Agrupar por", estado["df"].columns)
        columna_valor = col2.selectbox("Columna a operar", estado["df"].select_dtypes("number").columns)
        operacion = col3.selectbox("Operacion", ["sum", "mean", "count", "min", "max"])
        if st.button("Agrupar"):
            try:
                resultado = analisis.agrupar(estado, columna_grupo, columna_valor, operacion)
                st.bar_chart(resultado)
                st.dataframe(resultado)
            except Exception as e:
                st.error(str(e))

elif opcion == "Correlacion entre variables":
    if requiere_datos():
        columnas_num = analisis.columnas_numericas(estado)
        col1, col2 = st.columns(2)
        columna_x = col1.selectbox("Columna X", columnas_num)
        columna_y = col2.selectbox("Columna Y", columnas_num)
        if st.button("Calcular correlacion"):
            try:
                r = analisis.correlacion(estado, columna_x, columna_y)

                if abs(r) < 0.2:
                    interpretacion = "relacion practicamente nula."
                elif abs(r) < 0.5:
                    interpretacion = "relacion debil."
                elif abs(r) < 0.8:
                    interpretacion = "relacion moderada."
                else:
                    interpretacion = "relacion fuerte."

                st.metric(f"Correlacion de Pearson: {columna_x} vs {columna_y}", f"{r:.3f}")
                st.write(f"Interpretacion: {interpretacion}")
            except Exception as e:
                st.error(str(e))

elif opcion == "Representaciones graficas":
    if requiere_datos():
        tipo = st.selectbox("Tipo de grafico", list(analisis.TIPOS_GRAFICO.keys()))
        st.caption(analisis.TIPOS_GRAFICO[tipo]["descripcion"])

        col1, col2 = st.columns(2)
        opciones_x = analisis.columnas_validas_para(estado, tipo, "x")
        columna_x = col1.selectbox("Columna X", opciones_x)

        requisito_y = analisis.TIPOS_GRAFICO[tipo]["y"]
        columna_y = None
        if requisito_y != "ninguna":
            opciones_y = analisis.columnas_validas_para(estado, tipo, "y")
            etiqueta = "Columna Y (obligatoria)" if requisito_y == "numerica" else "Columna Y (opcional)"
            opciones_mostradas = opciones_y if requisito_y == "numerica" else [None] + opciones_y
            columna_y = col2.selectbox(etiqueta, opciones_mostradas)

        if st.button("Generar grafico"):
            try:
                fig = analisis.generar_grafico(estado, tipo, columna_x, columna_y)
                st.pyplot(fig)
            except Exception as e:
                st.error(str(e))

elif opcion == "Analisis adicional (inconsistencias)":
    if requiere_datos():
        inconsistencias = analisis.detectar_inconsistencias_texto(estado)
        if not inconsistencias:
            st.success("No se detectaron inconsistencias de mayusculas/espacios.")
        else:
            for columna, variantes in inconsistencias.items():
                st.write(f"**Columna '{columna}'**")
                for originales in variantes.values():
                    st.write(f"- {sorted(originales)}")

elif opcion == "Valores negativos":
    if requiere_datos():
        negativos = analisis.valores_negativos(estado)
        if negativos.empty:
            st.success("No se encontraron valores negativos.")
        else:
            st.write(f"{negativos.shape[0]} fila(s) con valores negativos:")
            st.dataframe(negativos)

elif opcion == "Ingresar nuevos datos":
    if requiere_datos():
        st.subheader("Submenu: Ingreso de datos")
        with st.expander("ℹ️ Reame - estructura del proyecto"):
            st.markdown("""
El submenu de ingreso de datos mantiene, en la medida de lo posible,
la misma estructura y tipos de datos del archivo CSV cargado.
Los registros ingresados quedan pendientes hasta presionar **Guardar**.
            """)

        estructura = analisis.estructura_columnas(estado)
        with st.form("form_nuevo_registro"):
            datos = {}
            for columna, tipo in estructura.items():
                datos[columna] = st.text_input(f"{columna} ({tipo})")
            enviado = st.form_submit_button("Ingresar registro")
        if enviado:
            try:
                st.success(analisis.ingresar_registro(estado, datos))
            except Exception as e:
                st.error(str(e))

        st.subheader("Registros ingresados (pendientes)")
        registros = analisis.mostrar_registros_ingresados(estado)
        if registros.empty:
            st.write("Sin registros pendientes.")
        else:
            st.dataframe(registros)
            if st.button("Guardar los nuevos datos"):
                st.success(analisis.guardar_registros_ingresados(estado))
