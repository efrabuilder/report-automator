# main_consola.py
# Interfaz de menu por consola para el Inventory Automator - Manejo de Datos (EDA).
# Ejecutar con: python main_consola.py

import os
import pandas as pd
import core.analisis as analisis

# Igual que en los notebooks se reutiliza un mismo df entre celdas, aqui
# se reutiliza un mismo diccionario "estado" entre todas las opciones del
# menu, para que "cargar" -> "analizar" -> "ingresar datos" trabajen
# siempre sobre el mismo conjunto de datos.
estado = analisis.nuevo_estado()


def pausa():
    input("\nPresione ENTER para continuar...")


def limpiar_pantalla():
    os.system("cls" if os.name == "nt" else "clear")


def pedir_entero(mensaje, minimo=None, maximo=None):
    # Pide un numero entero y reintenta si el formato es incorrecto.
    while True:
        valor = input(mensaje).strip()
        try:
            numero = int(valor)
            if minimo is not None and numero < minimo:
                print(f"Debe ser mayor o igual a {minimo}.")
                continue
            if maximo is not None and numero > maximo:
                print(f"Debe ser menor o igual a {maximo}.")
                continue
            return numero
        except ValueError:
            print("Entrada invalida, debe ser un numero entero.")



# Menu principal

def menu_principal():
    while True:
        print("\n" + "=" * 55)
        print(" INVENTORY AUTOMATOR - MANEJO DE DATOS (EDA)")
        print(" Estado: " + (
            f"'{os.path.basename(estado['ruta_actual'])}' cargado ({estado['df'].shape[0]} filas)"
            if estado["df"] is not None else "sin archivo cargado"
        ))
        print("=" * 55)
        print(" 1. Cargar archivo CSV o Excel")
        print(" 2. Mostrar informacion del conjunto de datos")
        print(" 3. Mostrar primeras y ultimas filas")
        print(" 4. Analizar tipos de datos")
        print(" 5. Analizar valores nulos")
        print(" 6. Analizar datos duplicados")
        print(" 7. Obtener estadisticas descriptivas")
        print(" 8. Filtrar o consultar datos")
        print(" 9. Realizar agrupaciones y operaciones")
        print("10. Generar representaciones graficas")
        print("11. Realizar analisis adicional (inconsistencias / limpieza)")
        print("12. Ingresar nuevos datos")
        print("13. Salir")
        print("=" * 55)

        opcion = input("Seleccione una opcion: ").strip()

        # RuntimeError = "todavia no hay archivo cargado" (aviso, no error).
        # ValueError/FileNotFoundError = datos o parametros invalidos (error).
        try:
            if opcion == "1":
                opcion_cargar_csv()
            elif opcion == "2":
                opcion_info_general()
            elif opcion == "3":
                opcion_primeras_ultimas()
            elif opcion == "4":
                opcion_tipos_datos()
            elif opcion == "5":
                opcion_valores_nulos()
            elif opcion == "6":
                opcion_duplicados()
            elif opcion == "7":
                opcion_estadisticas()
            elif opcion == "8":
                opcion_filtrar()
            elif opcion == "9":
                opcion_agrupar()
            elif opcion == "10":
                opcion_graficos()
            elif opcion == "11":
                opcion_analisis_adicional()
            elif opcion == "12":
                submenu_ingreso_datos()
            elif opcion == "13":
                print("Saliendo del programa. ¡Hasta luego!")
                break
            else:
                print("Opcion invalida, intente de nuevo (1-13).")
        except RuntimeError as e:
            print(f"\n[AVISO] {e}")
        except (ValueError, FileNotFoundError) as e:
            print(f"\n[ERROR] {e}")
        except Exception as e:
            print(f"\n[ERROR INESPERADO] {e}")

        if opcion != "13":
            pausa()



# Opciones del menu principal

def opcion_cargar_csv():
    # Abre un selector de archivos nativo (ventana de explorador) en vez de
    # pedir la ruta escrita. Si no hay entorno grafico disponible (por
    # ejemplo, una terminal remota sin pantalla), se cae de vuelta a pedir
    # la ruta por teclado para que la opcion nunca quede bloqueada.
    ruta = None
    try:
        import tkinter as tk
        from tkinter import filedialog
        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes("-topmost", True)
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo CSV o Excel",
            filetypes=[
                ("Archivos CSV y Excel", "*.csv *.xlsx *.xls"),
                ("Archivos CSV", "*.csv"),
                ("Archivos Excel", "*.xlsx *.xls"),
            ],
        )
        raiz.destroy()
    except Exception:
        print("(No se pudo abrir el selector grafico de archivos, "
              "se pedira la ruta manualmente)")

    if not ruta:
        ruta = input("Ingrese la ruta del archivo CSV o Excel: ").strip()

    print(analisis.cargar_archivo(estado, ruta))


def opcion_info_general():
    print(analisis.info_general(estado)["texto"])


def opcion_primeras_ultimas():
    n = pedir_entero("¿Cuantas filas desea ver? ", minimo=1)
    print("\n--- Primeras filas ---")
    print(analisis.primeras_filas(estado, n))
    print("\n--- Ultimas filas ---")
    print(analisis.ultimas_filas(estado, n))


def opcion_tipos_datos():
    print(analisis.tipos_datos(estado))


def opcion_valores_nulos():
    nulos = analisis.valores_nulos(estado)
    if nulos.empty:
        print("No hay valores nulos en el conjunto de datos.")
    else:
        print(nulos)


def opcion_duplicados():
    dup = analisis.datos_duplicados(estado)
    if dup.empty:
        print("No hay filas duplicadas.")
    else:
        print(f"Se encontraron {dup.shape[0]} filas duplicadas:")
        print(dup)


def opcion_estadisticas():
    pd.set_option("display.max_columns", None)
    print(analisis.estadisticas_descriptivas(estado))


def opcion_filtrar():
    print(f"Columnas disponibles: {list(estado['df'].columns)}")
    columna = input("Columna a filtrar: ").strip()
    print("Operadores: == != > >= < <= contiene")
    operador = input("Operador: ").strip()
    valor = input("Valor: ").strip()
    resultado = analisis.filtrar(estado, columna, operador, valor)
    print(f"\n{resultado.shape[0]} fila(s) encontradas:")
    print(resultado)


def opcion_agrupar():
    print(f"Columnas disponibles: {list(estado['df'].columns)}")
    columna_grupo = input("Columna para agrupar (categoria): ").strip()
    columna_valor = input("Columna a operar (numerica): ").strip()
    print("Operaciones: sum mean count min max")
    operacion = input("Operacion: ").strip()
    print(analisis.agrupar(estado, columna_grupo, columna_valor, operacion))


def opcion_graficos():
    print("\nTipos de grafico disponibles:")
    for nombre, info in analisis.TIPOS_GRAFICO.items():
        print(f"  {nombre:<11} -> {info['descripcion']}")

    tipo = input("\nTipo de grafico: ").strip().lower()
    if tipo not in analisis.TIPOS_GRAFICO:
        print(f"Tipo no valido. Opciones: {list(analisis.TIPOS_GRAFICO)}")
        return

    opciones_x = analisis.columnas_validas_para(estado, tipo, "x")
    print(f"Columnas validas para X: {opciones_x}")
    columna_x = input("Columna X: ").strip()

    requisito_y = analisis.TIPOS_GRAFICO[tipo]["y"]
    columna_y = None
    if requisito_y != "ninguna":
        opciones_y = analisis.columnas_validas_para(estado, tipo, "y")
        obligatoria = " (obligatoria)" if requisito_y == "numerica" else " (opcional, ENTER para omitir)"
        print(f"Columnas validas para Y{obligatoria}: {opciones_y}")
        columna_y = input("Columna Y: ").strip() or None

    ruta_salida = f"grafico_{tipo}_{columna_x}.png"
    analisis.generar_grafico(estado, tipo, columna_x, columna_y, ruta_salida=ruta_salida)
    print(f"Grafico guardado en: {ruta_salida}")


def opcion_analisis_adicional():
    print("\n--- Deteccion de inconsistencias en columnas de texto ---")
    inconsistencias = analisis.detectar_inconsistencias_texto(estado)
    if not inconsistencias:
        print("No se detectaron inconsistencias de mayusculas/espacios.")
    else:
        for columna, variantes in inconsistencias.items():
            print(f"\nColumna '{columna}':")
            for normalizado, originales in variantes.items():
                print(f"  -> {sorted(originales)}")

    respuesta = input("\n¿Desea limpiar el dataset ahora (quitar espacios y duplicados)? (s/n): ").strip().lower()
    if respuesta == "s":
        print(analisis.limpiar_datos(estado))

    respuesta_norm = input("¿Desea normalizar 'categoria' y 'vendedor' (unificar mayusculas/variantes)? (s/n): ").strip().lower()
    if respuesta_norm == "s":
        print(analisis.normalizar_categoria_texto(estado, "categoria"))
        print(analisis.normalizar_categoria_texto(estado, "vendedor"))

    respuesta_cf = input("¿Desea normalizar 'cliente_frecuente' (unificar Si/No)? (s/n): ").strip().lower()
    if respuesta_cf == "s":
        print(analisis.normalizar_categoria_texto(
            estado, "cliente_frecuente",
            mapa_valores={
                "Si": ["si", "Si", "SI", "sí", "Sí", "SÍ"],
                "No": ["no", "No", "NO"],
            },
        ))

    print("\n--- Relacion entre variables (correlacion) ---")
    respuesta_corr = input("¿Desea calcular la correlacion entre dos columnas numericas? (s/n): ").strip().lower()
    if respuesta_corr == "s":
        print(f"Columnas numericas disponibles: {analisis.columnas_numericas(estado)}")
        columna_x = input("Columna X: ").strip()
        columna_y = input("Columna Y: ").strip()
        r = analisis.correlacion(estado, columna_x, columna_y)
        print(f"\nCorrelacion de Pearson entre '{columna_x}' y '{columna_y}': {r:.3f}")
        if abs(r) < 0.2:
            print("Interpretacion: relacion practicamente nula.")
        elif abs(r) < 0.5:
            print("Interpretacion: relacion debil.")
        elif abs(r) < 0.8:
            print("Interpretacion: relacion moderada.")
        else:
            print("Interpretacion: relacion fuerte.")

    print("\n--- Valores negativos en columnas numericas ---")
    negativos = analisis.valores_negativos(estado)
    if negativos.empty:
        print("No se encontraron valores negativos.")
    else:
        print(f"Se encontraron {negativos.shape[0]} fila(s) con valores negativos:")
        print(negativos)


# Submenu - Ingreso de datos

def submenu_ingreso_datos():
    while True:
        print("\n--- Submenu: Ingreso de datos ---")
        print("1. Ingresar un nuevo registro")
        print("2. Mostrar registros ingresados")
        print("3. Guardar los nuevos datos")
        print("4. Regresar al menu principal")
        print("5. Reame (informacion de la estructura del proyecto)")

        opcion = input("Seleccione una opcion: ").strip()

        try:
            if opcion == "1":
                ingresar_registro_interactivo()
            elif opcion == "2":
                registros = analisis.mostrar_registros_ingresados(estado)
                print("Sin registros pendientes." if registros.empty else registros)
            elif opcion == "3":
                print(analisis.guardar_registros_ingresados(estado))
            elif opcion == "4":
                break
            elif opcion == "5":
                mostrar_readme()
            else:
                print("Opcion invalida, intente de nuevo (1-5).")
        except RuntimeError as e:
            print(f"\n[AVISO] {e}")
        except (ValueError, FileNotFoundError) as e:
            print(f"\n[ERROR] {e}")

        pausa()


def ingresar_registro_interactivo():
    estructura = analisis.estructura_columnas(estado)
    print("Ingrese los valores solicitados segun la estructura del dataset:")
    datos = {}
    for columna, tipo in estructura.items():
        valor = input(f"  {columna} ({tipo}): ").strip()
        datos[columna] = valor
    print(analisis.ingresar_registro(estado, datos))


def mostrar_readme():
    print("""
--- README del proyecto ---
Inventory Automator del modulo Manejo de Datos - EDA.
Estructura de archivos:
  core/analisis.py     -> logica compartida (carga, validacion, analisis, graficos)
  main_consola.py       -> esta interfaz de menu por consola
  main_tkinter.py        -> interfaz grafica de escritorio (Tkinter)
  app_streamlit.py       -> interfaz web (Streamlit)
  <archivo>.csv          -> dataset proporcionado por el docente

El submenu de ingreso de datos mantiene, en la medida de lo posible,
la misma estructura y tipos de datos del archivo CSV cargado.
""")


if __name__ == "__main__":
    menu_principal()
