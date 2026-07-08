# 📦 counting-tool

Aplicación web para comparar Manifiestos de Embarque vs Conteos manuales de bultos. Diseñada para dispositivos móviles.

## Funcionalidades

- **Cargar Manifiesto** — Sube archivos `.xls` o `.xlsx` y visualiza metadatos (MBL, contenedor, total de envíos, peso) más la tabla completa con filtro de búsqueda
- **Cargar Conteos** — Sube 1 o 2 archivos de conteo manual y normaliza automáticamente los formatos de BL
- **Dashboard de comparación** — KPIs de cobertura (Total BL, Contados, %), detalle por BL con estado y listado de BLs extraños o faltantes

## Tecnologías

- [Streamlit](https://streamlit.io)
- pandas, openpyxl, xlrd
- Python 3.14+

## Deploy en Render

1. Sube este repo a GitHub
2. En [Render](https://render.com) → New → Web Service
3. Conecta tu repo → Branch `main` → Start Command: `streamlit run app.py`
4. Plan **Free**

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

```
counting-tool/
├── app.py
├── requirements.txt
├── README.md
└── .streamlit/
    └── config.toml
```
