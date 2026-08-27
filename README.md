# Inventory Automator - Manejo de Datos (EDA)

## Estructura

- `core/analisis.py` -> logica compartida: carga, validaciones, limpieza, analisis, graficos.
- `main_consola.py` -> interfaz de menu por consola.
- `main_tkinter.py` -> interfaz grafica de escritorio (Tkinter).
- `app_streamlit.py` -> interfaz web (Streamlit).
- `ventas_tienda_tecnologia_ampliado.csv` -> dataset proporcionado por el docente.

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
