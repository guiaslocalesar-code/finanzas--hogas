# A Finanzas Hogar Worker

Migracion de una app de finanzas personales desde Google Apps Script a Cloudflare Workers con Hono + TypeScript, manteniendo Google Sheets como base de datos.

## Que incluye

- API REST sobre Cloudflare Workers
- Frontend estatico servido por el mismo Worker
- Integracion con Google Sheets API v4 por HTTP directo
- OAuth2 con Service Account y JWT firmado manualmente
- Deploy automatico con GitHub Actions en cada push a `main`
- Modulo adicional para tarjetas, resumenes, cuotas y reintegros del negocio

## Estructura

```text
.
|-- .github/workflows/deploy.yml
|-- AGENT.md
|-- README.md
|-- package.json
|-- tsconfig.json
|-- wrangler.toml
`-- src
    |-- index.ts
    |-- lib
    |   |-- finance-modules.ts
    |   |-- google-auth.ts
    |   |-- sheets.ts
    |   `-- types.ts
    `-- static
        `-- index.html
```

## Requisitos previos

- Node.js 20 o superior
- Cuenta de Cloudflare con Workers habilitado
- Un Google Sheet existente con una pestana llamada `Transacciones`
- Una Google Service Account con acceso al Google Sheet

## Hojas y headers esperados

La pestana `Transacciones` usa estos headers:

```text
id,type,amount,category,description,date,createdAt,dueDate
```

Si la hoja esta vacia, el Worker crea esos headers automaticamente en la primera fila.

La ampliacion modular de tarjetas usa estas hojas adicionales:

- `Tarjetas`
  - `cardId,issuer,brand,bank,holder,last4,closeDay,dueDay,active,createdAt`
- `ResumenesTarjeta`
  - `summaryId,cardId,issuer,bank,holder,fileName,statementDate,closingDate,dueDate,nextDueDate,totalAmount,minimumPayment,currency,rawText,parseStatus,createdAt`
- `CuotasProyectadas`
  - `projectionId,summaryId,cardId,issuer,monthLabel,yearMonth,amount,sourceType,confirmed,createdAt`
- `CuotasDetalle`
  - `installmentId,summaryId,cardId,purchaseDate,merchant,installmentNumber,installmentTotal,amount,dueMonth,dueDate,ownerType,businessPercent,businessAmount,personalAmount,reimbursementStatus,notes,createdAt`
- `ReintegrosNegocio`
  - `reimbursementId,sourceType,sourceId,cardId,concept,totalPaid,businessAmount,personalAmount,reimbursementStatus,reimbursementDueDate,reimbursedAmount,reimbursedDate,notes,createdAt`

Estas hojas se crean o migran en forma aditiva con `POST /api/setup/modules` o al usar por primera vez los endpoints nuevos.

## Instalacion local

1. Instalar dependencias:

   ```bash
   npm install
   ```

2. Configurar variables del Worker.

   En Cloudflare defini:

   - `SPREADSHEET_ID`
   - `SERVICE_ACCOUNT_EMAIL`

3. Configurar el secret privado en Cloudflare:

   ```bash
   npx wrangler secret put PRIVATE_KEY
   ```

   Pega la private key completa de la Service Account. Si la guardas con `\n`, el codigo la normaliza.

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

Agregar estos repository secrets:

- `CLOUDFLARE_ACCOUNT_ID`
- `CLOUDFLARE_API_TOKEN`
- `PRIVATE_KEY`
- `SPREADSHEET_ID`
- `SERVICE_ACCOUNT_EMAIL`

## Como obtener y compartir la Service Account

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

## Como configurar GitHub Actions

1. Ir a `Settings > Secrets and variables > Actions`.
2. Crear los secrets:
   - `CLOUDFLARE_ACCOUNT_ID`
   - `CLOUDFLARE_API_TOKEN`
   - `PRIVATE_KEY`
   - `SPREADSHEET_ID`
   - `SERVICE_ACCOUNT_EMAIL`
3. Hacer push a la rama `main`.

El workflow ejecuta:

- `npm ci`
- `npm run typecheck`
- `wrangler deploy`
- sincronizacion del secret `PRIVATE_KEY` en Cloudflare

## Como configurar Cloudflare

1. Crear el Worker.
2. Actualizar `wrangler.toml` con el `name` final.
3. Definir `SPREADSHEET_ID` y `SERVICE_ACCOUNT_EMAIL` como variables del Worker.
4. Cargar el secret:

   ```bash
   npx wrangler secret put PRIVATE_KEY
   ```

5. Desplegar:

   ```bash
   npm run deploy
   ```

## Endpoints

- `GET /`
- `GET /health`
- `GET /api/transactions`
- `POST /api/transactions`
- `PATCH /api/transactions`
- `DELETE /api/transactions/:id`
- `GET /debug/sheets`
- `POST /api/setup/modules`
- `GET /api/cards`
- `POST /api/cards`
- `PATCH /api/cards`
- `GET /api/card-summaries`
- `POST /api/card-summaries`
- `POST /api/card-summaries/import`
- `GET /api/installments/projections`
- `GET /api/installments/detail`
- `PATCH /api/installments/detail`
- `GET /api/installments/outlook`
- `GET /api/reimbursements`
- `POST /api/reimbursements`
- `PATCH /api/reimbursements`

## Contrato de datos base

Cada transaccion responde este formato:

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

- `amount` sale siempre como numero.
- `dueDate` se conserva aunque este vacio.
- Las fechas vacias se devuelven como string vacio.
- Si existen columnas extra en el sheet, no se rompen las operaciones de update.

## Como funciona la autenticacion Google

El Worker:

1. Construye un JWT manualmente.
2. Lo firma con `RS256` usando `crypto.subtle`.
3. Lo intercambia por un access token en `https://oauth2.googleapis.com/token`.
4. Usa ese token contra Google Sheets API v4.
5. Cachea el token en memoria del runtime cuando es posible.

## Notas de despliegue

- El primer despliegue puede requerir verificar que el nombre del Worker este libre.
- El tab `Transacciones` debe existir dentro del spreadsheet.
- Si el sheet ya tiene datos, la primera fila se interpreta como headers.
- Las hojas nuevas del modulo de tarjetas se crean sin tocar la estructura de `Transacciones`.

## Referencias usadas

- [Cloudflare Workers configuration](https://developers.cloudflare.com/workers/wrangler/configuration/)
- [Cloudflare assets binding](https://developers.cloudflare.com/workers/static-assets/binding/)
- [Cloudflare GitHub Actions deploy](https://developers.cloudflare.com/workers/ci-cd/external-cicd/github-actions/)
- [Hono on Cloudflare Workers](https://hono.dev/getting-started/cloudflare-workers)
