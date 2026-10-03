# CYSTEMS Data Lab

Portal educativo en Astro para aprender a conectar Power BI con una API de datos empresariales.

## Funciones

- Portal de estudio con una ruta guiada de 60 a 90 minutos.
- Dashboard web conectado a la API.
- Explorador de registros con filtros y paginación.
- API REST compatible con Power BI y CORS habilitado.
- Fuente en Google Sheets mediante una URL CSV publicada.
- Datos de demostración automáticos cuando la hoja no está configurada.

## Rutas

- `/` presentación del laboratorio.
- `/estudiar` secuencia didáctica.
- `/dashboard` indicadores en tiempo real.
- `/datos` explorador de la base.
- `/api` documentación para estudiantes.
- `/api/v1/ventas.json` registros para Power BI.
- `/api/v1/resumen.json` indicadores agregados.
- `/api/v1/schema.json` diccionario de campos.
- `/api/v1/health.json` estado del servicio y de la fuente.

## Desarrollo

```bash
npm install
npm run dev
```

Requiere Node.js 22.19 o superior. El proyecto también puede probarse con Docker:

```bash
docker build -t yuki-student .
docker run --rm -p 4321:4321 yuki-student
```

Copie `.env.example` como `.env` y configure `GOOGLE_SHEET_CSV_URL` con la URL CSV de la hoja `Ventas_En_Vivo`.

## Despliegue

El proyecto incluye `Dockerfile` y `compose.production.yaml` para ejecutarse detrás de Traefik en `api.cystems.ec`.
En el VPS de CYSTEMS, `.env` debe conservar `CYSTEMS_PROXY_NETWORK=edkudkydmldnkvfswxizh8fw`.

```bash
docker compose -f compose.production.yaml up -d --build
```

Después del despliegue, verifique:

```bash
curl https://api.cystems.ec/api/v1/health.json
curl "https://api.cystems.ec/api/v1/ventas.json?limit=5"
```

## Conexión desde Power BI

En Power BI use **Obtener datos > Web** e ingrese:

`https://api.cystems.ec/api/v1/ventas.json?refresh=1`

Seleccione la lista `data`, conviértala en tabla y expanda los registros.

## Materiales de clase

- `materiales/Practica_API_Cystems_Power_BI.xlsx`: fuente editable con 360 transacciones, hoja de control, metas y diccionario.
- `materiales/Guion_Practica_API_Cystems_90_Minutos.docx`: guion docente completo para una sesión de 90 minutos y ruta comprimida de 60 minutos.

## Publicar Google Sheets como fuente

1. Suba el Excel a Google Drive y ábralo con Google Sheets.
2. Conviértalo a una hoja nativa de Google.
3. Publique únicamente la hoja `Ventas_En_Vivo` como CSV.
4. Configure la URL resultante en `GOOGLE_SHEET_CSV_URL`.
5. Verifique que `/api/v1/health.json` muestre `source: google-sheets`.
