# app_streamlit.py
#
# Interfaz web (Streamlit) para el Inventory Automator - Manejo de Datos (EDA).
# Ejecutar con: streamlit run app_streamlit.py

import streamlit as st
import pandas as pd
import tempfile
import os
import io
import html
import matplotlib.pyplot as plt

import core.analisis as analisis

st.set_page_config(page_title="Inventory Automator - Manejo de Datos (EDA)", layout="wide")



# Estado de sesion: un diccionario "estado" persistente entre
# interacciones. Se guarda en st.session_state (y no en una variable
# de modulo) porque Streamlit puede atender a varios usuarios distintos
# dentro del mismo proceso: cada uno necesita su propio df cargado.

if "estado" not in st.session_state:
    st.session_state.estado = analisis.nuevo_estado()
estado = st.session_state.estado



# Colores del grafico

# Cada color se guarda en un diccionario "plano" persistente
# (st.session_state.estilo_plano) y los widgets usan claves "w_<campo>".
# Se hace asi porque Streamlit borra el valor de un widget cuando su
# seccion no se dibuja: el diccionario plano conserva la eleccion al
# navegar entre secciones.

CAMPOS_COLOR = [
    ("fondo", "Fondo del grafico"),
    ("fondo_grad_a", "Fondo del panel (arriba)"),
    ("fondo_grad_b", "Fondo del panel (abajo)"),
    ("texto", "Texto y numeros"),
    ("grilla", "Grilla y ejes"),
    ("acento", "Acento (titulo, tendencia)"),
]
ETIQUETAS_DESTINO = {etq: campo for campo, etq in CAMPOS_COLOR}
ETIQUETAS_DESTINO.update({f"Color {i + 1} de la paleta": f"p{i}" for i in range(6)})


def estilo_a_plano(e):
    plano = {k: v for k, v in e.items() if k != "paleta"}
    plano.update({f"p{i}": c for i, c in enumerate(e["paleta"])})
    return plano


def plano_a_estilo(plano):
    e = {k: v for k, v in plano.items() if not (k.startswith("p") and k[1:].isdigit())}
    e["paleta"] = [plano[f"p{i}"] for i in range(6)]
    return e


if "estilo_plano" not in st.session_state:
    st.session_state.estilo_plano = estilo_a_plano(analisis.resolver_estilo())


def _aplicar_tema():
    tema = analisis.TEMAS[st.session_state["w_tema"]]
    for k, v in estilo_a_plano(analisis.resolver_estilo(tema)).items():
        st.session_state[f"w_{k}"] = v


def _aplicar_color_nombrado():
    hexa = analisis.matplotlib.colors.CSS4_COLORS[st.session_state["w_css_nombre"]].upper()
    campo = ETIQUETAS_DESTINO[st.session_state["w_css_destino"]]
    st.session_state[f"w_{campo}"] = hexa


@st.cache_data
def _tabla_colores_html():
    filas = []
    for c in analisis.colores_css():
        r, g, b = c["rgb"]
        h, sat, l = c["hsl"]
        filas.append(
            f"<tr><td><span style='display:inline-block;width:38px;height:18px;"
            f"border:1px solid #8886;background:{c['hex']}'></span></td>"
            f"<td>{html.escape(c['nombre'])}</td><td>{c['hex']}</td>"
            f"<td>{r}, {g}, {b}</td><td>{h}, {sat}%, {l}%</td></tr>"
        )
    return (
        "<div style='max-height:320px;overflow-y:auto'><table style='width:100%;font-size:0.85rem'>"
        "<tr><th>Color</th><th>Nombre</th><th>Hex</th><th>RGB</th><th>HSL</th></tr>"
        + "".join(filas) + "</table></div>"
    )


def panel_colores():
    # Dibuja los controles de color y devuelve el diccionario de estilo
    # listo para pasarle a analisis.generar_grafico(estilo=...).
    plano = st.session_state.estilo_plano
    for k, v in plano.items():
        if f"w_{k}" not in st.session_state:
            st.session_state[f"w_{k}"] = v

    with st.expander("Colores y estilo"):
        st.selectbox("Tema", list(analisis.TEMAS), key="w_tema", on_change=_aplicar_tema)

        columnas = st.columns(3)
        for i, (campo, etiqueta) in enumerate(CAMPOS_COLOR):
            columnas[i % 3].color_picker(etiqueta, key=f"w_{campo}")
        st.checkbox("Fondo del panel con degradado", key="w_usar_degradado")

        st.caption("Paleta de colores (barras, lineas, pastel, cajas...)")
        columnas_paleta = st.columns(6)
        for i in range(6):
            columnas_paleta[i].color_picker(f"Color {i + 1}", key=f"w_p{i}")

        st.markdown("**Elegir un color por nombre**")
        c1, c2, c3 = st.columns([2, 2, 1])
        c1.selectbox("Color", sorted(analisis.matplotlib.colors.CSS4_COLORS), key="w_css_nombre")
        c2.selectbox("Aplicar a", list(ETIQUETAS_DESTINO), key="w_css_destino")
        c3.write("")
        c3.button("Aplicar", on_click=_aplicar_color_nombrado)
        with st.expander("Ver tabla de colores (nombre, HEX, RGB, HSL)"):
            st.markdown(_tabla_colores_html(), unsafe_allow_html=True)

    for k in plano:
        plano[k] = st.session_state[f"w_{k}"]
    return plano_a_estilo(plano)



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
    st.session_state["hojas_disponibles"] = []
    if extension in analisis.EXTENSIONES_EXCEL:
        try:
            hojas = analisis.listar_hojas_excel(ruta_temporal)
        except Exception as e:
            hojas = []
            st.sidebar.error(str(e))
        st.session_state["hojas_disponibles"] = hojas
        # La clave incluye el nombre del archivo: asi el selector de la barra
        # lateral y el de la seccion de graficos comparten el mismo valor.
        st.session_state["clave_hoja"] = f"hoja_sel_{archivo.name}"
        if len(hojas) > 1:
            hoja_elegida = st.sidebar.selectbox(
                "Elegir hoja a analizar", hojas, key=st.session_state["clave_hoja"])
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
    "Columnas calculadas (formula)",
    "KPIs",
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
        # Selector de hoja dentro de la seccion de graficos (solo si el
        # libro de Excel tiene mas de una hoja). Al cambiarlo se actualiza
        # el selector de la barra lateral y se recarga la hoja elegida.
        hojas = st.session_state.get("hojas_disponibles", [])
        if len(hojas) > 1:
            def _cambiar_hoja():
                st.session_state[st.session_state["clave_hoja"]] = st.session_state["hoja_graficos"]

            if estado["hoja_actual"] in hojas:
                st.session_state["hoja_graficos"] = estado["hoja_actual"]
            st.selectbox("Hoja del Excel a graficar", hojas, key="hoja_graficos", on_change=_cambiar_hoja)
            st.caption("Al cambiar de hoja se recargan los datos (se pierden limpiezas y columnas calculadas).")

        estilo = panel_colores()

        tipo = st.selectbox("Tipo de grafico", list(analisis.TIPOS_GRAFICO.keys()))
        st.caption(analisis.TIPOS_GRAFICO[tipo]["descripcion"])

        col1, col2 = st.columns(2)
        opciones_x = analisis.columnas_validas_para(estado, tipo, "x")
        # Para pastel/caja/violin se sugiere primero una columna con pocas
        # categorias (<= 12), porque con mas el grafico no es legible.
        indice_x = 0
        if tipo in ("pastel", "caja", "violin"):
            for i, nombre_col in enumerate(opciones_x):
                if estado["df"][nombre_col].nunique() <= 12:
                    indice_x = i
                    break
        columna_x = col1.selectbox("Columna X", opciones_x, index=indice_x, key=f"col_x_{tipo}")

        requisito_y = analisis.TIPOS_GRAFICO[tipo]["y"]
        columna_y = None
        if requisito_y != "ninguna":
            opciones_y = analisis.columnas_validas_para(estado, tipo, "y")
            etiqueta = "Columna Y (obligatoria)" if requisito_y == "numerica" else "Columna Y (opcional)"
            opciones_mostradas = opciones_y if requisito_y == "numerica" else [None] + opciones_y
            columna_y = col2.selectbox(etiqueta, opciones_mostradas)

        titulo_propio = st.text_input("Titulo del grafico (opcional)")

        # El grafico se redibuja solo al cambiar cualquier opcion o color
        # (ya no hace falta el boton "Generar").
        try:
            fig = analisis.generar_grafico(
                estado, tipo, columna_x, columna_y, estilo=estilo, titulo=titulo_propio or None
            )
            st.pyplot(fig)
            buffer = io.BytesIO()
            fig.savefig(buffer, format="png", dpi=150)
            st.download_button("Descargar PNG", buffer.getvalue(),
                               file_name=f"grafico_{tipo}.png", mime="image/png")
            plt.close(fig)
        except Exception as e:
            st.error(str(e))

elif opcion == "Columnas calculadas (formula)":
    if requiere_datos():
        st.write("Crea una columna nueva con una formula sobre las columnas existentes. "
                 "Despues aparece en los graficos y en los KPIs.")
        st.caption("Ejemplos: `cantidad * precio_unitario`, `total_venta * 0.13`, "
                   "`precio_unitario - 1000`. Para columnas con espacios use acentos graves: `precio unitario`.")
        st.write("Columnas numericas disponibles: " + ", ".join(f"`{c}`" for c in analisis.columnas_numericas(estado)))
        with st.form("form_formula"):
            c1, c2 = st.columns([1, 2])
            nombre_nueva = c1.text_input("Nombre de la columna nueva")
            formula = c2.text_input("Formula")
            crear = st.form_submit_button("Crear columna")
        if crear:
            try:
                st.success(analisis.agregar_columna_formula(estado, nombre_nueva, formula))
                st.dataframe(estado["df"][[nombre_nueva.strip()]].head(10))
            except Exception as e:
                st.error(str(e))

elif opcion == "KPIs":
    if requiere_datos():
        estilo = panel_colores()
        kpis = st.session_state.setdefault("kpis", [])

        with st.form("form_kpi"):
            c1, c2, c3 = st.columns(3)
            nombre_kpi = c1.text_input("Nombre del KPI")
            columna_kpi = c2.selectbox("Columna", list(estado["df"].columns))
            operacion_kpi = c3.selectbox("Operacion", list(analisis.OPERACIONES_KPI))
            filtro_kpi = st.text_input("Filtro opcional (formula)",
                                       placeholder="region == 'Heredia' and cantidad > 1")
            agregar = st.form_submit_button("Agregar KPI")
        if agregar:
            kpis.append({
                "nombre": nombre_kpi.strip() or f"{operacion_kpi} de {columna_kpi}",
                "columna": columna_kpi, "operacion": operacion_kpi, "filtro": filtro_kpi.strip(),
            })

        if not kpis:
            st.info("Todavia no hay KPIs. Agregue uno con el formulario.")
        else:
            por_fila = 3
            for inicio in range(0, len(kpis), por_fila):
                columnas = st.columns(por_fila)
                for j, kpi in enumerate(kpis[inicio:inicio + por_fila]):
                    indice = inicio + j
                    with columnas[j]:
                        try:
                            r = analisis.calcular_kpi(estado, kpi["columna"], kpi["operacion"], kpi["filtro"])
                            valor = analisis.formatear_kpi(r["valor"])
                            detalle = f"{kpi['operacion']} de {kpi['columna']} · {r['filas']} filas"
                            if kpi["filtro"]:
                                detalle += f" · {kpi['filtro']}"
                            color_borde = estilo["paleta"][indice % len(estilo["paleta"])]
                            st.markdown(
                                f"<div style='background:{estilo['fondo']};border:1px solid {estilo['grilla']};"
                                f"border-top:5px solid {color_borde};border-radius:10px;padding:14px 16px'>"
                                f"<div style='color:{estilo['texto']};font-size:0.9rem;opacity:.8'>{html.escape(kpi['nombre'])}</div>"
                                f"<div style='color:{estilo['acento']};font-size:2rem;font-weight:700'>{valor}</div>"
                                f"<div style='color:{estilo['texto']};font-size:0.75rem;opacity:.6'>{html.escape(detalle)}</div>"
                                f"</div>",
                                unsafe_allow_html=True,
                            )
                        except Exception as e:
                            st.error(f"{kpi['nombre']}: {e}")
                        if st.button("Quitar", key=f"quitar_kpi_{indice}"):
                            kpis.pop(indice)
                            st.rerun()

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
