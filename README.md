# Validador_Servitel
Validador Maestro para servitel 

- Autofix: mayúsculas, PAIS='VENEZUELA' si vacío, limpieza de espacios.
- Validación fechas: dd/mm/aaaa (dayfirst).
- Web: celdas con error en rojo, botón ERRORES.xlsx (bordes rojos, merge 'Unnamed:*' -> 'INFORMACION DEL ABONADO').
- Descarga validado bloqueada hasta 0 errores.

## Uso
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app_streamlit_reglas.py
