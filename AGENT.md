# AGENT.md

## Objetivo del proyecto
Este repositorio contiene una app de finanzas personales migrada desde Google Apps Script a Cloudflare Workers, manteniendo Google Sheets como base de datos.

## Stack obligatorio
- Cloudflare Workers
- Hono
- TypeScript
- Google Sheets API v4 vía HTTP
- JWT firmado manualmente para Service Account
- Frontend estático servido desde el Worker

## Restricciones
- No usar `googleapis`
- No usar SDKs de Node incompatibles con Workers
- No romper compatibilidad con la hoja `Transacciones`
- No cambiar headers del sheet sin migración explícita
- No commitear secrets ni archivos `.json` privados
- No hardcodear private keys en el código

## Estructura esperada
- `src/index.ts`: rutas Hono y server principal
- `src/lib/google-auth.ts`: firma JWT y obtención de access token
- `src/lib/sheets.ts`: funciones de lectura/escritura Sheets API
- `src/lib/types.ts`: tipos compartidos
- `src/static/index.html`: frontend principal
- `.github/workflows/deploy.yml`: deploy automático
- `wrangler.toml`: config del Worker

## Convenciones de API
- `GET /health` devuelve `{ ok: true }`
- `GET /api/transactions` devuelve array de transacciones
- `POST /api/transactions` crea una transacción
- `PATCH /api/transactions` actualiza una transacción por `id`
- `DELETE /api/transactions/:id` elimina una transacción por `id`
- `GET /debug/sheets` se puede usar temporalmente para diagnostico de Google Sheets
- `POST /api/setup/modules` inicializa las hojas modulares nuevas si no existen
- `GET /api/cards`, `POST /api/cards`, `PATCH /api/cards` gestionan tarjetas
- `GET /api/card-summaries`, `POST /api/card-summaries`, `POST /api/card-summaries/import` gestionan resúmenes de tarjeta
- `GET /api/installments/projections`, `GET /api/installments/detail`, `PATCH /api/installments/detail`, `GET /api/installments/outlook` gestionan cuotas y proyección por mes
- `GET /api/reimbursements`, `POST /api/reimbursements`, `PATCH /api/reimbursements` gestionan reintegros del negocio

## Formato de datos
Headers esperados del Google Sheet:
`id,type,amount,category,description,date,createdAt,dueDate`

Cada transacción debe respetar:
- `id`: string
- `type`: `income` o `expense`
- `amount`: number
- `category`: string
- `description`: string
- `date`: `YYYY-MM-DD`
- `createdAt`: ISO string
- `dueDate`: `YYYY-MM-DD` o vacío

## Reglas de implementación
- Siempre leer la primera fila como headers
- Si la hoja está vacía, crear headers automáticamente
- Convertir `amount` a number en salida
- No perder columnas existentes
- Manejar strings vacíos en fechas
- Mantener lógica simple y predecible
- La hoja `Transacciones` se mantiene como fuente existente y no se reestructura
- Las hojas `Tarjetas`, `ResumenesTarjeta`, `CuotasProyectadas`, `CuotasDetalle` y `ReintegrosNegocio` deben crearse o migrarse de forma aditiva
- Si faltan headers nuevos en esas hojas, agregarlos al final sin borrar columnas previas

## Frontend
- No usar `google.script.run`
- Consumir la API del Worker con `fetch`
- Mantener UI responsive
- Priorizar legibilidad sobre complejidad visual

## Deploy
- Cada push a `main` debe quedar listo para deploy automático
- El workflow de GitHub no debe requerir pasos manuales extra
- El repo debe ser deployable con `npm run deploy`

## Cuando hagas cambios
1. Explicá primero qué archivos vas a tocar.
2. Hacé cambios mínimos y consistentes.
3. No reestructures todo si no hace falta.
4. Si cambiás el contrato de datos, actualizá README y tipos.
5. Si agregás endpoints, documentalos.

## Qué evitar
- dependencias innecesarias
- abstracciones excesivas
- frameworks extra
- secretos en texto plano
- romper compatibilidad con Cloudflare Workers
