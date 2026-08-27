# core/analisis.py
#
# Modulo central de logica para el Inventory Automator - Manejo de Datos (EDA).
# Aqui vive toda la carga, validacion, limpieza, analisis y graficos.
# Las tres interfaces (consola, Tkinter, Streamlit) llaman a estas mismas
# funciones para no duplicar logica ni arriesgar resultados distintos
# entre ellas.
#
# En vez de una clase, el "estado" del proyecto cargado (el DataFrame,
# la ruta del archivo, los registros ingresados a mano) se guarda en un
# diccionario comun que cada funcion recibe como primer parametro y
# modifica directamente, igual que en los notebooks se trabaja sobre un
# mismo df reutilizado entre celdas.

import os
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # se cambia a un backend interactivo desde main_tkinter si aplica
import matplotlib.pyplot as plt


def nuevo_estado():
    # Diccionario "en blanco" que representa un proyecto sin archivo
    # cargado todavia. Cada interfaz crea uno de estos al arrancar.
    return {
        "df": None,
        "ruta_actual": None,
        "registros_ingresados": [],
    }



# 1. Carga de datos

def cargar_csv(estado, ruta):
    # Carga un archivo CSV dentro de estado["df"]. Devuelve un mensaje de
    # exito. Lanza FileNotFoundError o ValueError con mensajes claros si
    # algo sale mal (requerimiento del informe: "archivos que no puedan
    # ser cargados correctamente").
    if not ruta:
        raise ValueError("Debe indicar una ruta de archivo.")
    if not os.path.isfile(ruta):
        raise FileNotFoundError(f"No se encontro el archivo: {ruta}")
    if not ruta.lower().endswith(".csv"):
        raise ValueError("El archivo debe tener extension .csv")

    try:
        df = pd.read_csv(ruta)
    except pd.errors.EmptyDataError:
        raise ValueError("El archivo CSV esta vacio.")
    except pd.errors.ParserError as e:
        raise ValueError(f"El archivo CSV tiene un formato invalido: {e}")

    if df.empty:
        raise ValueError("El archivo CSV no contiene registros.")

    if "cantidad" in df.columns and "precio_unitario" in df.columns:
        df["total_venta"] = df["cantidad"] * df["precio_unitario"]

    estado["df"] = df
    estado["ruta_actual"] = ruta
    estado["registros_ingresados"] = []
    return f"Archivo cargado correctamente: {os.path.basename(ruta)} ({df.shape[0]} filas, {df.shape[1]} columnas)"


def verificar_carga(estado):
    # Todas las funciones de analisis llaman esto primero. Se usa un
    # RuntimeError comun (no una clase de excepcion propia) para que las
    # interfaces lo distingan facilmente de un ValueError real y lo
    # muestren como "aviso" en vez de "error".
    if estado["df"] is None:
        raise RuntimeError("Debe cargar un archivo CSV antes de realizar esta operacion.")



# 2. Informacion general / estructura

def info_general(estado):
    verificar_carga(estado)
    df = estado["df"]
    buffer = []
    buffer.append(f"Archivo: {os.path.basename(estado['ruta_actual'])}")
    buffer.append(f"Filas: {df.shape[0]}")
    buffer.append(f"Columnas: {df.shape[1]}")
    buffer.append(f"Nombres de columnas: {list(df.columns)}")
    buffer.append(f"Memoria aproximada: {df.memory_usage(deep=True).sum() / 1024:.2f} KB")
    return {
        "texto": "\n".join(buffer),
        "filas": df.shape[0],
        "columnas": df.shape[1],
        "nombres_columnas": list(df.columns),
    }


def primeras_filas(estado, n=5):
    verificar_carga(estado)
    return estado["df"].head(n)


def ultimas_filas(estado, n=5):
    verificar_carga(estado)
    return estado["df"].tail(n)



# 3. Tipos de datos

def tipos_datos(estado):
    verificar_carga(estado)
    return estado["df"].dtypes.rename("tipo_dato").reset_index().rename(columns={"index": "columna"})



# 4. Valores nulos

def valores_nulos(estado):
    verificar_carga(estado)
    df = estado["df"]
    nulos = df.isnull().sum()
    porcentaje = (nulos / len(df) * 100).round(2)
    resultado = pd.DataFrame({"nulos": nulos, "porcentaje_%": porcentaje})
    return resultado[resultado["nulos"] > 0].sort_values("nulos", ascending=False)



# 5. Duplicados

def datos_duplicados(estado):
    verificar_carga(estado)
    df = estado["df"]
    return df[df.duplicated(keep=False)].sort_values(list(df.columns))



# 6. Inconsistencias en columnas de texto (mayusculas/espacios/acentos)

def detectar_inconsistencias_texto(estado):
    # Para cada columna de texto, arma un reporte con los valores unicos
    # que colapsarian a uno solo si se normalizan (trim + minusculas).
    # Ayuda a justificar el punto "datos inconsistentes" del informe.
    verificar_carga(estado)
    df = estado["df"]
    columnas_texto = df.select_dtypes(include="object").columns
    reporte = {}
    for col in columnas_texto:
        valores = df[col].dropna().astype(str)
        normalizados = valores.str.strip().str.lower()
        grupos = {}
        for original, norm in zip(valores, normalizados):
            grupos.setdefault(norm, set()).add(original)
        variantes = {k: v for k, v in grupos.items() if len(v) > 1}
        if variantes:
            reporte[col] = variantes
    return reporte


def valores_negativos(estado, columnas=None):
    # Filas con valores negativos en columnas donde no tiene sentido
    # (cantidad, precio_unitario por defecto). Ayuda a documentar
    # errores de captura como parte de la revision de calidad de datos.
    verificar_carga(estado)
    df = estado["df"]
    columnas = columnas or ["cantidad", "precio_unitario"]
    columnas = [c for c in columnas if c in df.columns]
    if not columnas:
        return df.iloc[0:0]
    mascara = (df[columnas] < 0).any(axis=1)
    return df[mascara]


def limpiar_datos(estado, eliminar_duplicados=True, normalizar_texto=True, columnas_texto=None):
    # Limpieza reproducible (opcion "limpiar tabla" del menu):
    # - quita espacios y unifica mayusculas/minusculas en columnas de texto
    # - elimina filas duplicadas
    # No imputa nulos automaticamente: eso se deja como decision del
    # analisis (se reporta pero no se inventa un valor).
    verificar_carga(estado)
    df = estado["df"].copy()
    acciones = []

    if normalizar_texto:
        cols = columnas_texto or list(df.select_dtypes(include="object").columns)
        for col in cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].where(df[col].str.lower() != "nan", other=pd.NA)
        acciones.append(f"Se quitaron espacios en blanco en columnas de texto: {cols}")

    if eliminar_duplicados:
        antes = len(df)
        df = df.drop_duplicates()
        acciones.append(f"Se eliminaron {antes - len(df)} filas duplicadas")

    estado["df"] = df.reset_index(drop=True)
    return "\n".join(acciones) if acciones else "No se aplico ninguna limpieza."


def normalizar_categoria_texto(estado, columna, mapa_valores=None):
    # Unifica variantes de una columna categorica (ej: 'Si'/'si'/'SI' -> 'Si').
    # mapa_valores es opcional: {valor_normalizado: [variantes]}. Si no se
    # da, normaliza a Title Case sobre el texto sin espacios.
    verificar_carga(estado)
    df = estado["df"]
    if columna not in df.columns:
        raise ValueError(f"La columna '{columna}' no existe.")

    # BUGFIX: con el dtype "str" que pandas usa por defecto desde 2.x/3.x,
    # astype(str) ya NO convierte los nulos en el texto "nan": los deja
    # como NaN real. Antes se asumia que se volvian texto, por eso mas
    # abajo sorted(serie.unique()) truena con "'<' not supported between
    # instances of 'float' and 'str'" apenas la columna tiene un nulo.
    # Se guarda la mascara de nulos ANTES de tocar nada, y se restauran
    # como nulos reales al final (no se inventa una categoria "Nan").
    es_nulo = df[columna].isna()
    serie = df[columna].astype(str).str.strip()

    if mapa_valores:
        inverso = {}
        for correcto, variantes in mapa_valores.items():
            for v in variantes:
                inverso[v.strip().lower()] = correcto
        serie = serie.apply(lambda x: inverso.get(x.lower(), x))
    else:
        serie = serie.str.title()

    serie = serie.mask(es_nulo, other=pd.NA)

    df[columna] = serie
    return f"Columna '{columna}' normalizada. Valores unicos ahora: {sorted(serie.dropna().unique())}"


# 7. Estadisticas descriptivas

def estadisticas_descriptivas(estado):
    verificar_carga(estado)
    return estado["df"].describe(include="all").transpose()



# 8. Filtrar / consultar

def filtrar(estado, columna, operador, valor):
    # operador: uno de '==', '!=', '>', '>=', '<', '<=', 'contiene'
    verificar_carga(estado)
    df = estado["df"]
    if columna not in df.columns:
        raise ValueError(f"La columna '{columna}' no existe.")

    serie = df[columna]

    if operador == "contiene":
        return df[serie.astype(str).str.contains(str(valor), case=False, na=False)]

    # intentar convertir el valor al tipo de la columna cuando aplica
    try:
        if pd.api.types.is_numeric_dtype(serie):
            valor = float(valor)
    except (TypeError, ValueError):
        pass

    operadores = {
        "==": serie == valor,
        "!=": serie != valor,
        ">": serie > valor,
        ">=": serie >= valor,
        "<": serie < valor,
        "<=": serie <= valor,
    }
    if operador not in operadores:
        raise ValueError(f"Operador no valido: {operador}")
    return df[operadores[operador]]



# 9. Agrupaciones y operaciones

def agrupar(estado, columna_grupo, columna_valor, operacion):
    # operacion: 'sum', 'mean', 'count', 'min', 'max'
    verificar_carga(estado)
    df = estado["df"]
    for c in (columna_grupo, columna_valor):
        if c not in df.columns:
            raise ValueError(f"La columna '{c}' no existe.")
    if operacion not in ("sum", "mean", "count", "min", "max"):
        raise ValueError(f"Operacion no valida: {operacion}")

    return getattr(df.groupby(columna_grupo)[columna_valor], operacion)().sort_values(ascending=False)


# 9b. Relaciones entre variables (correlacion)

def correlacion(estado, columna_x, columna_y):
    # Correlacion de Pearson entre dos columnas numericas (-1 a 1).
    # Cerca de 0 = sin relacion lineal; cerca de 1 o -1 = relacion fuerte.
    verificar_carga(estado)
    df = estado["df"]
    for c in (columna_x, columna_y):
        if c not in df.columns:
            raise ValueError(f"La columna '{c}' no existe.")
    numericas = set(columnas_numericas(estado))
    if columna_x not in numericas or columna_y not in numericas:
        raise ValueError(
            f"La correlacion requiere columnas numericas. "
            f"Columnas numericas disponibles: {sorted(numericas)}"
        )
    return df[columna_x].corr(df[columna_y])



# 10. Graficos

# Descriptor de cada tipo de grafico: que tipo de columna espera en X e Y,
# y una explicacion corta. Las tres interfaces usan esto para mostrar
# solo las combinaciones validas y evitar errores como "no numeric data
# to plot" (ej: pedir un histograma sobre una columna de texto). Al ser
# un diccionario a nivel de modulo (no un atributo de clase), se importa
# igual desde cualquiera de las tres interfaces.
TIPOS_GRAFICO = {
    "barras": {
        "x": "cualquiera",
        "y": "numerica_opcional",
        "descripcion": "Compara categorias. Y opcional: si se omite, cuenta filas por categoria.",
    },
    "linea": {
        "x": "cualquiera",
        "y": "numerica_opcional",
        "descripcion": "Muestra una tendencia (ideal con fechas en X). Y opcional: si se omite, cuenta filas.",
    },
    "histograma": {
        "x": "numerica",
        "y": "ninguna",
        "descripcion": "Muestra la distribucion de una columna numerica.",
    },
    "dispersion": {
        "x": "numerica",
        "y": "numerica",
        "descripcion": "Compara dos columnas numericas entre si.",
    },
    "pastel": {
        "x": "categorica",
        "y": "ninguna",
        "descripcion": "Muestra proporciones de una columna categorica (pocas categorias).",
    },
}


def columnas_numericas(estado):
    verificar_carga(estado)
    return list(estado["df"].select_dtypes(include="number").columns)


def columnas_categoricas(estado):
    verificar_carga(estado)
    return list(estado["df"].select_dtypes(exclude="number").columns)


def columnas_validas_para(estado, tipo, eje):
    # Devuelve la lista de columnas que tiene sentido ofrecer para el
    # eje 'x' o 'y' de un tipo de grafico dado. Uso pensado para llenar
    # los selectores/Combobox de las interfaces.
    if tipo not in TIPOS_GRAFICO:
        raise ValueError(f"Tipo de grafico no valido: {tipo}")
    requisito = TIPOS_GRAFICO[tipo][eje]
    if requisito == "ninguna":
        return []
    if requisito == "numerica" or requisito == "numerica_opcional":
        return columnas_numericas(estado)
    if requisito == "categorica":
        return columnas_categoricas(estado)
    return list(estado["df"].columns)  # 'cualquiera'


def generar_grafico(estado, tipo, columna_x, columna_y=None, ruta_salida=None):
    # tipo: 'barras', 'linea', 'histograma', 'dispersion', 'pastel'
    # Devuelve la figura de matplotlib (para Tkinter/Streamlit) y opcionalmente
    # la guarda en disco (para el informe en Word).
    # Valida con mensajes claros ANTES de graficar, para nunca dejar pasar
    # errores crudos de matplotlib como "no numeric data to plot".
    verificar_carga(estado)
    df = estado["df"]

    if tipo not in TIPOS_GRAFICO:
        raise ValueError(f"Tipo de grafico no valido: '{tipo}'. Opciones: {list(TIPOS_GRAFICO)}")
    if columna_x not in df.columns:
        raise ValueError(f"La columna X '{columna_x}' no existe.")
    if columna_y and columna_y not in df.columns:
        raise ValueError(f"La columna Y '{columna_y}' no existe.")

    requisito_x = TIPOS_GRAFICO[tipo]["x"]
    requisito_y = TIPOS_GRAFICO[tipo]["y"]
    numericas = set(columnas_numericas(estado))

    if requisito_x == "numerica" and columna_x not in numericas:
        raise ValueError(
            f"El grafico de '{tipo}' necesita que la columna X sea numerica. "
            f"'{columna_x}' es de texto/categoria. Columnas numericas disponibles: {sorted(numericas)}"
        )
    if requisito_x == "categorica" and columna_x in numericas:
        raise ValueError(
            f"El grafico de '{tipo}' funciona mejor con una columna categorica en X. "
            f"'{columna_x}' es numerica. Columnas categoricas disponibles: {columnas_categoricas(estado)}"
        )
    if requisito_y == "numerica" and not columna_y:
        raise ValueError(f"El grafico de '{tipo}' requiere una columna Y numerica.")
    if requisito_y == "numerica" and columna_y not in numericas:
        raise ValueError(
            f"La columna Y '{columna_y}' debe ser numerica para '{tipo}'. "
            f"Columnas numericas disponibles: {sorted(numericas)}"
        )
    if requisito_y == "ninguna" and columna_y:
        columna_y = None  # se ignora silenciosamente, no es un error

    fig, ax = plt.subplots(figsize=(8, 5))

    if tipo == "barras":
        datos = df.groupby(columna_x)[columna_y].sum().sort_values(ascending=False) if columna_y else df[columna_x].value_counts()
        datos.plot(kind="bar", ax=ax, color="#4C72B0")
        ax.set_ylabel(columna_y or "conteo")
    elif tipo == "linea":
        datos = df.groupby(columna_x)[columna_y].sum() if columna_y else df[columna_x].value_counts().sort_index()
        datos.plot(kind="line", ax=ax, marker="o", color="#55A868")
        ax.set_ylabel(columna_y or "conteo")
    elif tipo == "histograma":
        df[columna_x].plot(kind="hist", bins=20, ax=ax, color="#C44E52")
    elif tipo == "dispersion":
        ax.scatter(df[columna_x], df[columna_y], alpha=0.5, color="#8172B2")
        ax.set_xlabel(columna_x)
        ax.set_ylabel(columna_y)
    elif tipo == "pastel":
        datos = df[columna_x].value_counts()
        if len(datos) > 12:
            raise ValueError(
                f"'{columna_x}' tiene {len(datos)} categorias distintas, "
                "demasiadas para un grafico de pastel legible. Proba con 'barras'."
            )
        ax.pie(datos.values, labels=datos.index, autopct="%1.1f%%")

    ax.set_title(f"{tipo.capitalize()}: {columna_x}" + (f" vs {columna_y}" if columna_y else ""))
    fig.tight_layout()

    if ruta_salida:
        fig.savefig(ruta_salida, dpi=120)

    return fig



# 11. Submenu - ingreso manual de nuevos registros

def estructura_columnas(estado):
    # Devuelve columna -> tipo de dato, para guiar el ingreso manual.
    verificar_carga(estado)
    return {col: str(dtype) for col, dtype in estado["df"].dtypes.items()}


def ingresar_registro(estado, datos):
    # Valida y agrega un registro nuevo a estado["registros_ingresados"]
    # (aun no al df principal; se confirma con guardar_registros_ingresados).
    verificar_carga(estado)
    df = estado["df"]
    columnas_esperadas = set(df.columns)
    columnas_recibidas = set(datos.keys())
    faltantes = columnas_esperadas - columnas_recibidas
    if faltantes:
        raise ValueError(f"Faltan campos obligatorios: {sorted(faltantes)}")

    registro_validado = {}
    for col, valor in datos.items():
        dtype = df[col].dtype
        try:
            if pd.api.types.is_integer_dtype(dtype):
                registro_validado[col] = int(valor)
            elif pd.api.types.is_float_dtype(dtype):
                registro_validado[col] = float(valor)
            else:
                registro_validado[col] = str(valor)
        except (TypeError, ValueError):
            raise ValueError(f"El campo '{col}' espera un valor de tipo {dtype}, se recibio: '{valor}'")

    estado["registros_ingresados"].append(registro_validado)
    return f"Registro agregado ({len(estado['registros_ingresados'])} pendiente(s) de guardar)."


def mostrar_registros_ingresados(estado):
    return pd.DataFrame(estado["registros_ingresados"])


def guardar_registros_ingresados(estado):
    # Incorpora los registros ingresados manualmente al DataFrame principal.
    verificar_carga(estado)
    if not estado["registros_ingresados"]:
        return "No hay registros nuevos para guardar."
    nuevos = pd.DataFrame(estado["registros_ingresados"])
    estado["df"] = pd.concat([estado["df"], nuevos], ignore_index=True)
    cantidad = len(estado["registros_ingresados"])
    estado["registros_ingresados"] = []
    return f"Se guardaron {cantidad} registro(s) nuevo(s) en el dataset (total ahora: {len(estado['df'])} filas)."



# 12. Exportar

def exportar_csv(estado, ruta_salida):
    verificar_carga(estado)
    estado["df"].to_csv(ruta_salida, index=False)
    return f"Datos exportados a: {ruta_salida}"
