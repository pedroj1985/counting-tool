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

## Deploy en Render (gratis)

### 1. Haz el repo público

Si el repo es privado, Render no podrá verlo:

```
GitHub → counting-tool repo → Settings → Danger Zone → Change visibility → Public
```

### 2. Conecta Render con GitHub

- Ve a [Render](https://render.com) y crea una cuenta (GitHub)
- Dashboard → **Account Settings** → **GitHub** → **Manage GitHub** → **Connect account**
- Dale acceso al repo `counting-tool`

### 3. Crea el Web Service

- Dashboard → **New +** → **Web Service**
- Selecciona tu repo `pedroj1985/counting-tool`
- Llena los campos:

| Campo | Valor |
|---|---|
| **Name** | `counting-tool` |
| **Branch** | `main` |
| **Runtime** | `Python 3` |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `streamlit run app.py --server.port $PORT` |
| **Plan** | **Free** |

### 4. Deploy

- Click **Create Web Service**
- Render instalará dependencias y desplegará automáticamente (~3 min)
- Tu app quedará en: `https://counting-tool.onrender.com`

> ⚠️ En el plan Free, Render duerme el servicio tras 15 min sin actividad. Al recargar la página tarda ~30s en despertar.

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
