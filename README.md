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

## Deploy en Streamlit Cloud (gratis, sin tarjeta)

### 1. Haz el repo público

```
GitHub → counting-tool repo → Settings → Danger Zone → Change visibility → Public
```

### 2. Despliega en Streamlit Cloud

- Ve a [Streamlit Community Cloud](https://streamlit.io/cloud)
- Inicia sesión con tu cuenta de **GitHub**
- Click **"New app"** → **"Deploy from GitHub"**
- Selecciona el repo `pedroj1985/counting-tool`
- Configura:

| Campo | Valor |
|---|---|
| **Branch** | `main` |
| **Main file** | `app.py` |
| **App URL** | `counting-tool.streamlit.app` |

- Click **"Deploy"**

Tu app estará disponible en **`https://counting-tool.streamlit.app`** en ~2 minutos. Sin tarjeta de crédito, sin tiempos de espera.

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
