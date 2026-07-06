# Lucciano's Europa — Reporte de Ventas Diario (automático por GitHub Actions)

Todos los días, en forma automática, el sistema baja los reportes de ventas que
Shares manda por mail (**uno por local**), arma la comparación contra el mismo
período de 2025 y envía el reporte por Gmail.

## Flujo (un job)

1. **Shares** manda por mail el cierre diario de cada local a una casilla de Gmail.
2. El workflow corre por horario y ejecuta:
   - `fetch_shares.py` → lee la casilla por IMAP, baja los adjuntos de cada local y
     escribe `data/ventas_dia.json` = `{fecha, ventas:{local: venta}}`.
   - `generar_acum2025.py` → suma del master 2025 el acumulado del mes hasta el día
     de corte, por local, y escribe `data/acum_interanual.json`.
   - `report.py` → toma el acumulado del mes de `data/acumulado.json`, le suma la
     venta del día, calcula la variación vs 2025 y genera el HTML.
   - `send_mail.py` → envía el mail por Gmail SMTP.

Diseño desacoplado: `report.py` y `generar_acum2025.py` sólo leen JSON, así que no
dependen de cómo llegue el mail. Todo lo específico del correo vive en `fetch_shares.py`.

## Locales

**Propias:** Roma, Barcelona 1, Barcelona 2, Madrid, Málaga 1
**Franquicias:** Málaga 3, Granada, Alicante, Valencia

Madrid y Alicante abrieron en 2026: no tienen histórico 2025, así que se muestran con
su venta y acumulado del mes pero **sin comparación** (variación "—"). La regla es
automática: cualquier local cuyo acumulado 2025 dé cero se marca "sin comp.".

## El master 2025 (comparativo)

`data/Ventas_Master_2025.xlsx`, hoja **"Ventas 2025"**: una fila por (local, día) con
la **Venta Bruta** (columna `Venta Bruta (EUR)`). Se arma una sola vez con
`unificar_ventas.py` a partir de los reportes anuales de Shares. Cubre 2025 completo,
así que sirve para todo 2026.

> **Métrica:** Venta Bruta (incluye IVA). El dato diario debe medir lo mismo (bruta).

## Protección anti doble-conteo y reinicio mensual

Igual que USA: si una fecha ya fue procesada (`last_date` en `data/acumulado.json`) no
se vuelve a sumar ni a enviar; y el 1° de cada mes el acumulado se reinicia solo.

## Secrets (Settings → Secrets and variables → Actions)

| Secret | Qué es |
|---|---|
| `IMAP_USER` / `IMAP_APP_PASS` | Casilla que recibe los mails de Shares. |
| `GMAIL_USER` / `GMAIL_APP_PASS` | Casilla que envía el reporte. |
| `MAIL_TO` | Destinatario(s), separados por coma. |

## Horario

`repository_dispatch` disparado por cron externo a las 7:00 AR + `schedule` nativo de
respaldo a las 7:30 AR (10:30 UTC). El cron nativo de GitHub suele atrasarse, por eso
el disparo real es externo.

## Pendiente de ajustar con el primer mail real

`fetch_shares.py` tiene marcados con `TODO` los tres puntos que dependen de ver un mail:
remitente de Shares, y cómo identificar el local de cada mail (remitente / asunto /
nombre interno del adjunto).
