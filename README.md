# Inventory Automator - Manejo de Datos (EDA)

## Estructura

- `core/analisis.py` -> logica compartida: carga, validaciones, limpieza, analisis, graficos.
- `main_consola.py` -> interfaz de menu por consola.
- `main_tkinter.py` -> interfaz grafica de escritorio (Tkinter).
- `app_streamlit.py` -> interfaz web (Streamlit).
- `ventas_tienda_tecnologia_ampliado.csv` -> dataset proporcionado por el docente.

## Novedades (interfaz Streamlit)

- **Hoja del Excel** elegible tambien desde la seccion de graficos.
- **Colores y estilo**: tema, fondo, degradado, texto, grilla, acento y paleta de 6 colores; ademas un selector de colores CSS por nombre con su tabla (HEX/RGB/HSL).
- **Nuevos graficos**: barras horizontales, area, violin y Pareto. Descarga en PNG.
- **Columnas calculadas**: formulas como `cantidad * precio_unitario * 0.13`.
- **KPIs**: tarjetas con suma, promedio, mediana, minimo, maximo, conteo o valores unicos, con filtro opcional.

## Como ejecutar cada interfaz

### Consola

```
python main_consola.py
```

### Tkinter (requiere entorno grafico / escritorio)

```
python main_tkinter.py
```

### Streamlit (abre en el navegador)

```
pip install streamlit
streamlit run app_streamlit.py
```

## Dependencias

```
pip install pandas matplotlib streamlit
```
