# Cuadros de Subejercicio — DGAPSGB

App de Streamlit que genera los "Cuadros de Subejercicio" (avance del gasto
ejercido por Programa Presupuestario) a partir del MAP crudo, con un cuadro
por cada UR + Pp.

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Uso

1. Sube el MAP (CSV o XLSX crudo, con columnas `UNIDAD`, `IDEN_PROY`,
   `PROYECTO`, `PARTIDA` y las columnas mensuales `ORI_*`, `MOD_*`, `EJE_*`).
2. Ajusta en la barra lateral, si hace falta:
   - **MDP mínimo de disponible por capítulo** (default 1.0 MDP): un capítulo
     solo se incluye si su disponible es mayor o igual a este monto.
   - **Capítulo máximo a considerar** (default 4999): excluye por default
     Inversión (5000/6000/7000), que se reporta aparte.
   - **Lista manual**: si se llena (formato `UR,Pp` por renglón, ej.
     `923,S318`), solo se generan esas combinaciones exactas.
3. Revisa el **resumen general** (UR, Programa, Modificado, Ejercido,
   Disponible y % disponible/modificado) antes de descargar.
4. Descarga el Excel final con un cuadro (hoja) por cada UR + Pp.

## Estructura

- `app.py` — interfaz de Streamlit.
- `cuadros_lib.py` — lógica de cálculo y generación de las hojas de Excel
  (catálogos de UR/Pp/Capítulo, lectura del MAP, selección, formato).
- `requirements.txt` — dependencias.

## Despliegue en Streamlit Community Cloud

1. Sube este repo a GitHub.
2. En [share.streamlit.io](https://share.streamlit.io), conecta el repo y
   selecciona `app.py` como archivo principal.
