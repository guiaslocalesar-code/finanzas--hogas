# A Finanzas Hogar Worker

Migración de una app de finanzas personales desde Google Apps Script a Cloudflare Workers con Hono + TypeScript, manteniendo Google Sheets como base de datos.

## Qué incluye

- API REST sobre Cloudflare Workers
- Frontend estático servido por el mismo Worker
- Integración con Google Sheets API v4 por HTTP directo
- OAuth2 con Service Account y JWT firmado manualmente
- Deploy automático con GitHub Actions en cada push a `main`

## Estructura

```text
.
├── .github/workflows/deploy.yml
├── AGENT.md
├── README.md
├── package.json
├── tsconfig.json
├── wrangler.toml
└── src
    ├── index.ts
    ├── lib
    │   ├── google-auth.ts
    │   ├── sheets.ts
    │   └── types.ts
    └── static
        └── index.html
```

## Requisitos previos

- Node.js 20 o superior
- Cuenta de Cloudflare con Workers habilitado
- Un Google Sheet existente con una pestaña llamada `Transacciones`
- Una Google Service Account con acceso al Google Sheet

## Headers esperados en Google Sheets

La pestaña `Transacciones` usa estos headers:

```text
id,type,amount,category,description,date,createdAt,dueDate
```

Si la hoja está vacía, el Worker crea esos headers automáticamente en la primera fila.

## Instalación local

1. Instalar dependencias:

   ```bash
   npm install
   ```

2. Configurar variables del Worker.

   En `wrangler.toml` dejá:

   - `SPREADSHEET_ID`
   - `SERVICE_ACCOUNT_EMAIL`

3. Configurar el secret privado en Cloudflare:

   ```bash
   npx wrangler secret put PRIVATE_KEY
   ```

   Pegá la private key completa de la Service Account. Si la guardás con `\n`, el código la normaliza.

4. Levantar el entorno local:

   ```bash
   npm run dev
   ```

## Variables y secretos necesarios

### En Cloudflare Workers

- `SPREADSHEET_ID`
- `SERVICE_ACCOUNT_EMAIL`
- `PRIVATE_KEY`

### En GitHub Actions

Agregar estos secrets y variables en el repositorio:

- `CLOUDFLARE_ACCOUNT_ID`
- `CLOUDFLARE_API_TOKEN`
- `PRIVATE_KEY`
- `SPREADSHEET_ID` como Repository Variable
- `SERVICE_ACCOUNT_EMAIL` como Repository Variable

## Cómo obtener y compartir la Service Account

1. Crear una Service Account en Google Cloud.
2. Habilitar la Google Sheets API en el proyecto.
3. Generar una clave privada tipo JSON.
4. Copiar:
   - `client_email` como `SERVICE_ACCOUNT_EMAIL`
   - `private_key` como `PRIVATE_KEY`
5. Compartir el Google Sheet con el email de la Service Account con permisos de editor.

Importante:

- No subir el JSON de credenciales al repositorio.
- No hardcodear la `PRIVATE_KEY`.

## Cómo configurar GitHub Actions

1. Ir a `Settings > Secrets and variables > Actions`.
2. Crear los secrets:
   - `CLOUDFLARE_ACCOUNT_ID`
   - `CLOUDFLARE_API_TOKEN`
   - `PRIVATE_KEY`
3. Crear las variables:
   - `SPREADSHEET_ID`
   - `SERVICE_ACCOUNT_EMAIL`
4. Hacer push a la rama `main`.

El workflow ejecuta:

- `npm ci`
- `npm run typecheck`
- `wrangler deploy --var ...`
- sincronización del secret `PRIVATE_KEY` en Cloudflare

## Cómo configurar Cloudflare

1. Crear el Worker.
2. Actualizar `wrangler.toml` con el `name` final y `SPREADSHEET_ID`.
3. Definir `SERVICE_ACCOUNT_EMAIL` en `wrangler.toml` o como variable del entorno.
4. Cargar el secret:

   ```bash
   npx wrangler secret put PRIVATE_KEY
   ```

5. Desplegar:

   ```bash
   npm run deploy
   ```

Si desplegás por GitHub Actions, el workflow sube `PRIVATE_KEY` y pasa `SPREADSHEET_ID` y `SERVICE_ACCOUNT_EMAIL` desde GitHub automáticamente.

## Endpoints

- `GET /`
- `GET /health`
- `GET /api/transactions`
- `POST /api/transactions`
- `PATCH /api/transactions`
- `DELETE /api/transactions/:id`
- `GET /debug/sheets` temporal para diagnostico

## Contrato de datos

Cada transacción responde este formato:

```json
{
  "id": "1712500000000",
  "type": "expense",
  "amount": 12500,
  "category": "Supermercado",
  "description": "Compra semanal",
  "date": "2026-04-08",
  "createdAt": "2026-04-08T14:30:00.000Z",
  "dueDate": ""
}
```

Notas:

- `amount` sale siempre como número.
- `dueDate` se conserva aunque esté vacío.
- Las fechas vacías se devuelven como string vacío.
- Si existen columnas extra en el sheet, no se rompen las operaciones de update.

## Cómo funciona la autenticación Google

El Worker:

1. Construye un JWT manualmente.
2. Lo firma con `RS256` usando `crypto.subtle`.
3. Lo intercambia por un access token en `https://oauth2.googleapis.com/token`.
4. Usa ese token contra Google Sheets API v4.
5. Cachea el token en memoria del runtime cuando es posible.

## Notas de despliegue

- El primer despliegue puede requerir verificar que el nombre del Worker esté libre.
- El tab `Transacciones` debe existir dentro del spreadsheet.
- Si el sheet ya tiene datos, la primera fila se interpreta como headers.

## Referencias usadas

- [Cloudflare Workers configuration](https://developers.cloudflare.com/workers/wrangler/configuration/)
- [Cloudflare assets binding](https://developers.cloudflare.com/workers/static-assets/binding/)
- [Cloudflare GitHub Actions deploy](https://developers.cloudflare.com/workers/ci-cd/external-cicd/github-actions/)
- [Hono on Cloudflare Workers](https://hono.dev/getting-started/cloudflare-workers)
