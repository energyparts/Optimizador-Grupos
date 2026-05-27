# Groups Optimizer — Energy Parts

Aplicación web para generar combinaciones de grupos (G1–G7) a partir de existencias reales y notas de crédito de Johnson Controls.

## Funcionalidad

1. **Sube tu archivo de existencias** (.xlsx) — detecta automáticamente el stock por grupo
2. **Sube la Nota de Crédito** (.pdf de JCA) — extrae subtotal y puntos vía IA
3. **Genera combinaciones** que respetan el stock disponible, ordenadas por uniformidad
4. **Descarga Excel** con todas las combinaciones + información del documento

## Estructura del proyecto

```
groups-optimizer/
├── app.py              # Flask backend
├── requirements.txt
├── render.yaml         # Configuración Render
├── Procfile
├── .gitignore
├── static/
│   └── logo.png        # ← Coloca aquí el logo de Energy Parts
└── templates/
    └── index.html      # Frontend
```

## Deployment en Render

### 1. Subir a GitHub

```bash
git init
git add .
git commit -m "Initial commit — Groups Optimizer"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/groups-optimizer.git
git push -u origin main
```

### 2. Crear servicio en Render

1. Ve a [render.com](https://render.com) → **New** → **Web Service**
2. Conecta tu repositorio de GitHub
3. Render detecta el `render.yaml` automáticamente
4. En **Environment Variables** agrega:
   - `ANTHROPIC_API_KEY` = tu API key de Anthropic

### 3. Logo

Coloca el logo de Energy Parts como `static/logo.png` antes de hacer push.

## Variables de entorno

| Variable | Descripción |
|---|---|

| `PORT` | Puerto (Render lo asigna automáticamente) |

## Desarrollo local

```bash
pip install -r requirements.txt

python app.py
# Abre http://localhost:5000
```

## Fórmula de puntos

```
Monto sin IVA ÷ $294 = Puntos totales
Puntos × $294 = Costo total
```

Grupos: G1=1pt·$294, G2=1.3pt·$382.2, G3=1.7pt·$499.8, G4=2pt·$588, G5=2.5pt·$735, G6/G7=5pt·$1,470
