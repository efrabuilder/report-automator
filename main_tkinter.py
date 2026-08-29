# main_tkinter.py
#
# Interfaz grafica de escritorio (Tkinter) para el Inventory Automator -
# Manejo de Datos (EDA). Ejecutar con: python main_tkinter.py
#
# Igual que en main_consola.py, no se usa una clase para guardar el
# estado del programa: los widgets y el diccionario "estado" del
# proyecto (df, ruta_actual, registros_ingresados) se guardan en
# variables sueltas a nivel de modulo, y las funciones los leen y
# modifican directamente.

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import core.analisis as analisis

estado = analisis.nuevo_estado()

# Variables de los widgets principales, creadas en construir_ventana().
ventana = None
label_estado = None
texto_resultado = None
frame_grafico = None  # se crea dinamicamente al pedir un grafico



# Construccion de la ventana

def construir_ventana():
    global ventana
    ventana = tk.Tk()
    ventana.title("Inventory Automator - Manejo de Datos (EDA)")
    ventana.geometry("1000x650")

    crear_barra_superior()
    crear_menu_lateral()
    crear_area_resultados()

    actualizar_estado()


def crear_barra_superior():
    global label_estado
    barra = tk.Frame(ventana, bg="#2c3e50", height=50)
    barra.pack(side="top", fill="x")

    tk.Button(barra, text="Cargar archivo CSV/Excel", command=accion_cargar_csv,
              bg="#2980b9", fg="white", relief="flat", padx=12, pady=6
              ).pack(side="left", padx=10, pady=8)

    tk.Button(barra, text="Limpiar tabla", command=accion_limpiar_tabla,
              bg="#c0392b", fg="white", relief="flat", padx=12, pady=6
              ).pack(side="left", padx=5, pady=8)

    tk.Button(barra, text="Normalizar categoria/vendedor", command=accion_normalizar_categorias,
              bg="#8e44ad", fg="white", relief="flat", padx=12, pady=6
              ).pack(side="left", padx=5, pady=8)

    tk.Button(barra, text="Normalizar cliente frecuente", command=accion_normalizar_cliente_frecuente,
              bg="#8e44ad", fg="white", relief="flat", padx=12, pady=6
              ).pack(side="left", padx=5, pady=8)

    label_estado = tk.Label(barra, text="", bg="#2c3e50", fg="white")
    label_estado.pack(side="right", padx=15)


def crear_menu_lateral():
    lateral = tk.Frame(ventana, width=220, bg="#ecf0f1")
    lateral.pack(side="left", fill="y")

    opciones = [
        ("Info del dataset", accion_info),
        ("Primeras/ultimas filas", accion_primeras_ultimas),
        ("Tipos de datos", accion_tipos),
        ("Valores nulos", accion_nulos),
        ("Datos duplicados", accion_duplicados),
        ("Estadisticas descriptivas", accion_estadisticas),
        ("Filtrar / consultar", accion_filtrar),
        ("Agrupaciones", accion_agrupar),
        ("Correlacion entre variables", accion_correlacion),
        ("Generar grafico", accion_grafico),
        ("Inconsistencias de texto", accion_inconsistencias),
        ("Valores negativos", accion_valores_negativos),
        ("Ingresar nuevo registro", accion_ingresar_registro),
        ("Ver registros ingresados", accion_ver_ingresados),
        ("Guardar registros nuevos", accion_guardar_ingresados),
    ]
    for texto, comando in opciones:
        tk.Button(lateral, text=texto, command=comando, anchor="w",
                  relief="flat", bg="#ecf0f1", padx=10, pady=8
                  ).pack(fill="x")


def crear_area_resultados():
    global texto_resultado
    contenedor = tk.Frame(ventana)
    contenedor.pack(side="right", fill="both", expand=True)

    texto_resultado = tk.Text(contenedor, wrap="none", font=("Consolas", 10))
    texto_resultado.pack(fill="both", expand=True, padx=8, pady=8)



# Utilidades

def actualizar_estado():
    if estado["df"] is not None:
        label_estado.config(
            text=f"Archivo: {estado['ruta_actual'].split('/')[-1]} "
                 f"({estado['df'].shape[0]} filas)"
        )
    else:
        label_estado.config(text="Sin archivo cargado")


def mostrar_texto(texto):
    quitar_grafico()
    texto_resultado.pack(fill="both", expand=True, padx=8, pady=8)
    texto_resultado.delete("1.0", tk.END)
    texto_resultado.insert(tk.END, texto)


def mostrar_dataframe(df):
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    mostrar_texto(df.to_string())


def quitar_grafico():
    global frame_grafico
    if frame_grafico is not None:
        frame_grafico.destroy()
        frame_grafico = None


def mostrar_grafico(fig):
    global frame_grafico
    texto_resultado.pack_forget()
    quitar_grafico()
    frame_grafico = tk.Frame(texto_resultado.master)
    frame_grafico.pack(fill="both", expand=True, padx=8, pady=8)
    canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)


def manejar_error(e):
    # RuntimeError = "todavia no hay archivo cargado" (aviso).
    # Cualquier otra cosa se muestra como error.
    if isinstance(e, RuntimeError):
        messagebox.showwarning("Aviso", str(e))
    else:
        messagebox.showerror("Error", str(e))


def pedir_valor(titulo, etiquetas):
    # Ventana emergente simple para pedir varios valores de texto.
    ventana_pedido = tk.Toplevel(ventana)
    ventana_pedido.title(titulo)
    ventana_pedido.geometry("320x" + str(60 + 40 * len(etiquetas)))
    entradas = []
    for etiqueta in etiquetas:
        tk.Label(ventana_pedido, text=etiqueta).pack(anchor="w", padx=10, pady=(8, 0))
        entrada = tk.Entry(ventana_pedido, width=35)
        entrada.pack(padx=10)
        entradas.append(entrada)

    resultado = {"valores": None}

    def confirmar():
        resultado["valores"] = [e.get().strip() for e in entradas]
        ventana_pedido.destroy()

    tk.Button(ventana_pedido, text="Aceptar", command=confirmar).pack(pady=10)
    ventana_pedido.grab_set()
    ventana.wait_window(ventana_pedido)
    return resultado["valores"]



# Acciones - barra superior

def accion_cargar_csv():
    ruta = filedialog.askopenfilename(filetypes=[
        ("Archivos CSV y Excel", "*.csv *.xlsx *.xls"),
        ("Archivos CSV", "*.csv"),
        ("Archivos Excel", "*.xlsx *.xls"),
    ])
    if not ruta:
        return
    try:
        mensaje = analisis.cargar_archivo(estado, ruta)
        actualizar_estado()
        mostrar_texto(mensaje)
    except Exception as e:
        manejar_error(e)


def accion_limpiar_tabla():
    try:
        mensaje = analisis.limpiar_datos(estado)
        mostrar_texto(mensaje)
    except Exception as e:
        manejar_error(e)


def accion_normalizar_categorias():
    try:
        msg1 = analisis.normalizar_categoria_texto(estado, "categoria")
        msg2 = analisis.normalizar_categoria_texto(estado, "vendedor")
        mostrar_texto(msg1 + "\n\n" + msg2)
    except Exception as e:
        manejar_error(e)


def accion_normalizar_cliente_frecuente():
    try:
        mensaje = analisis.normalizar_categoria_texto(
            estado, "cliente_frecuente",
            mapa_valores={
                "Si": ["si", "Si", "SI", "sí", "Sí", "SÍ"],
                "No": ["no", "No", "NO"],
            },
        )
        mostrar_texto(mensaje)
    except Exception as e:
        manejar_error(e)


# Acciones - menu lateral

def accion_info():
    try:
        mostrar_texto(analisis.info_general(estado)["texto"])
    except Exception as e:
        manejar_error(e)


def accion_primeras_ultimas():
    try:
        texto = "--- Primeras 5 filas ---\n" + analisis.primeras_filas(estado, 5).to_string()
        texto += "\n\n--- Ultimas 5 filas ---\n" + analisis.ultimas_filas(estado, 5).to_string()
        mostrar_texto(texto)
    except Exception as e:
        manejar_error(e)


def accion_tipos():
    try:
        mostrar_dataframe(analisis.tipos_datos(estado))
    except Exception as e:
        manejar_error(e)


def accion_nulos():
    try:
        nulos = analisis.valores_nulos(estado)
        mostrar_texto("Sin valores nulos." if nulos.empty else nulos.to_string())
    except Exception as e:
        manejar_error(e)


def accion_duplicados():
    try:
        dup = analisis.datos_duplicados(estado)
        mostrar_texto("Sin filas duplicadas." if dup.empty else dup.to_string())
    except Exception as e:
        manejar_error(e)


def accion_estadisticas():
    try:
        mostrar_dataframe(analisis.estadisticas_descriptivas(estado))
    except Exception as e:
        manejar_error(e)


def accion_filtrar():
    try:
        analisis.verificar_carga(estado)
        valores = pedir_valor("Filtrar datos", [
            f"Columna {list(estado['df'].columns)}:",
            "Operador (== != > >= < <= contiene):",
            "Valor:",
        ])
        if not valores:
            return
        columna, operador, valor = valores
        resultado = analisis.filtrar(estado, columna, operador, valor)
        mostrar_texto(f"{resultado.shape[0]} fila(s) encontradas\n\n" + resultado.to_string())
    except Exception as e:
        manejar_error(e)


def accion_agrupar():
    try:
        analisis.verificar_carga(estado)
        valores = pedir_valor("Agrupaciones", [
            "Columna para agrupar:",
            "Columna a operar (numerica):",
            "Operacion (sum mean count min max):",
        ])
        if not valores:
            return
        columna_grupo, columna_valor, operacion = valores
        resultado = analisis.agrupar(estado, columna_grupo, columna_valor, operacion)
        mostrar_texto(resultado.to_string())
    except Exception as e:
        manejar_error(e)


def accion_correlacion():
    try:
        analisis.verificar_carga(estado)
        valores = pedir_valor("Correlacion entre variables", [
            f"Columnas numericas {analisis.columnas_numericas(estado)}",
            "Columna X:",
            "Columna Y:",
        ])
        if not valores:
            return
        _, columna_x, columna_y = valores
        r = analisis.correlacion(estado, columna_x, columna_y)

        if abs(r) < 0.2:
            interpretacion = "relacion practicamente nula."
        elif abs(r) < 0.5:
            interpretacion = "relacion debil."
        elif abs(r) < 0.8:
            interpretacion = "relacion moderada."
        else:
            interpretacion = "relacion fuerte."

        texto = (
            f"Correlacion de Pearson entre '{columna_x}' y '{columna_y}': {r:.3f}\n\n"
            f"Interpretacion: {interpretacion}"
        )
        mostrar_texto(texto)
    except Exception as e:
        manejar_error(e)


def accion_grafico():
    try:
        analisis.verificar_carga(estado)
        abrir_dialogo_grafico()
    except Exception as e:
        manejar_error(e)


def abrir_dialogo_grafico():
    # Ventana guiada para armar un grafico: se elige el tipo primero, y
    # las listas de columnas X/Y se actualizan solas para mostrar solo
    # las columnas que tienen sentido para ese tipo (evita pedir un
    # histograma sobre una columna de texto, etc.).
    ventana_grafico = tk.Toplevel(ventana)
    ventana_grafico.title("Generar grafico")
    ventana_grafico.geometry("380x320")
    ventana_grafico.grab_set()

    tipos = list(analisis.TIPOS_GRAFICO.keys())

    tk.Label(ventana_grafico, text="Tipo de grafico:").pack(anchor="w", padx=10, pady=(10, 0))
    combo_tipo = ttk.Combobox(ventana_grafico, values=tipos, state="readonly")
    combo_tipo.pack(fill="x", padx=10)
    combo_tipo.current(0)

    label_descripcion = tk.Label(ventana_grafico, text="", wraplength=340, fg="#555", justify="left")
    label_descripcion.pack(anchor="w", padx=10, pady=(4, 10))

    tk.Label(ventana_grafico, text="Columna X:").pack(anchor="w", padx=10)
    combo_x = ttk.Combobox(ventana_grafico, state="readonly")
    combo_x.pack(fill="x", padx=10)

    label_y = tk.Label(ventana_grafico, text="Columna Y:")
    combo_y = ttk.Combobox(ventana_grafico, state="readonly")

    def actualizar_campos(*_):
        tipo = combo_tipo.get()
        info = analisis.TIPOS_GRAFICO[tipo]
        label_descripcion.config(text=info["descripcion"])

        combo_x["values"] = analisis.columnas_validas_para(estado, tipo, "x")
        combo_x.set("")
        if combo_x["values"]:
            combo_x.current(0)

        if info["y"] == "ninguna":
            label_y.pack_forget()
            combo_y.pack_forget()
        else:
            etiqueta = "Columna Y (obligatoria):" if info["y"] == "numerica" else "Columna Y (opcional):"
            label_y.config(text=etiqueta)
            label_y.pack(anchor="w", padx=10)
            combo_y.pack(fill="x", padx=10)
            opciones_y = analisis.columnas_validas_para(estado, tipo, "y")
            combo_y["values"] = [""] + opciones_y if info["y"] == "numerica_opcional" else opciones_y
            combo_y.set("")

    combo_tipo.bind("<<ComboboxSelected>>", actualizar_campos)
    actualizar_campos()

    def generar():
        tipo = combo_tipo.get()
        columna_x = combo_x.get()
        columna_y = combo_y.get() or None
        if not columna_x:
            messagebox.showwarning("Aviso", "Debe seleccionar una columna X.")
            return
        try:
            fig = analisis.generar_grafico(estado, tipo, columna_x, columna_y)
            ventana_grafico.destroy()
            mostrar_grafico(fig)
        except Exception as e:
            manejar_error(e)

    tk.Button(ventana_grafico, text="Generar", command=generar, bg="#2980b9", fg="white"
              ).pack(pady=15)


def accion_inconsistencias():
    try:
        inconsistencias = analisis.detectar_inconsistencias_texto(estado)
        if not inconsistencias:
            mostrar_texto("No se detectaron inconsistencias de mayusculas/espacios.")
            return
        lineas = []
        for columna, variantes in inconsistencias.items():
            lineas.append(f"Columna '{columna}':")
            for originales in variantes.values():
                lineas.append(f"  -> {sorted(originales)}")
        mostrar_texto("\n".join(lineas))
    except Exception as e:
        manejar_error(e)


def accion_valores_negativos():
    try:
        negativos = analisis.valores_negativos(estado)
        if negativos.empty:
            mostrar_texto("No se encontraron valores negativos.")
        else:
            mostrar_texto(negativos.to_string())
    except Exception as e:
        manejar_error(e)


# Acciones - submenu ingreso de datos

def accion_ingresar_registro():
    try:
        analisis.verificar_carga(estado)
        estructura = analisis.estructura_columnas(estado)
        etiquetas = [f"{col} ({tipo}):" for col, tipo in estructura.items()]
        valores = pedir_valor("Ingresar nuevo registro", etiquetas)
        if not valores:
            return
        datos = dict(zip(estructura.keys(), valores))
        mensaje = analisis.ingresar_registro(estado, datos)
        mostrar_texto(mensaje)
    except Exception as e:
        manejar_error(e)


def accion_ver_ingresados():
    try:
        registros = analisis.mostrar_registros_ingresados(estado)
        mostrar_texto("Sin registros pendientes." if registros.empty else registros.to_string())
    except Exception as e:
        manejar_error(e)


def accion_guardar_ingresados():
    try:
        mensaje = analisis.guardar_registros_ingresados(estado)
        actualizar_estado()
        mostrar_texto(mensaje)
    except Exception as e:
        manejar_error(e)


if __name__ == "__main__":
    construir_ventana()
    ventana.mainloop()
