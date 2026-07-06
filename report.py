"""
Lucciano's Europa - Reporte de Ventas Diario
Compara el mes en curso (2026) contra el mismo período del año anterior (2025).

Flujo (todo desacoplado por JSON, distinto de USA que usaba xlsx intermedios):
  1. data/ventas_dia.json     -> venta del día por local (lo arma fetch_shares.py)
  2. data/acum_interanual.json-> acumulado 2025 por local (lo arma generar_acum2025.py)
  3. data/acumulado.json      -> estado del acumulado del mes en curso (persistido)
  4. acum26 = acum previo del mes + venta del día
  5. Variación = (acum26 - acum25)/acum25   [si acum25==0 -> "sin comparación"]
  6. Genera preview.html

Locales sin historia 2025 (Madrid, Alicante, o Granada antes de abrir) tienen
acum25==0: se muestran con su venta y acumulado 2026, pero la variación va como "—".
"""
import json
import os
import sys
from pathlib import Path
from datetime import datetime

from locales import PROPIAS, FRANQUICIAS, BRANCH_ORDER

MESES_ES = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
            7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
DIAS_ES = {0: "LUNES", 1: "MARTES", 2: "MIÉRCOLES", 3: "JUEVES", 4: "VIERNES", 5: "SÁBADO", 6: "DOMINGO"}
MES_CORTO = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
             7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

BASE = Path(__file__).parent


def money(v):
    """Formato europeo: 1.234,56 €"""
    s = f"{v:,.2f}".replace(",", "§").replace(".", ",").replace("§", ".")
    return f"{s} €"


def pct_txt(pct):
    return f"{pct:.1f}".replace(".", ",") + "%"


def load_json(path, default=None):
    p = Path(path)
    if not p.exists():
        if default is not None:
            return default
        raise SystemExit(f"[ERROR] Falta {path}.")
    return json.loads(p.read_text(encoding="utf-8"))


def save_json(path, data):
    Path(path).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def build_report():
    dia = load_json(BASE / "data" / "ventas_dia.json")
    fecha = datetime.strptime(dia["fecha"], "%Y-%m-%d").date()
    venta_dia = {b: float(dia["ventas"].get(b, 0.0)) for b in BRANCH_ORDER}

    inter = load_json(BASE / "data" / "acum_interanual.json")
    acum25 = {b: float(inter["acum25"].get(b, 0.0)) for b in BRANCH_ORDER}
    a25_fin = datetime.strptime(inter["hasta"], "%Y-%m-%d").date()

    # Validación de período espejo: el acum25 debe terminar el mismo día/mes del año previo.
    errores = []
    if a25_fin.year != fecha.year - 1:
        errores.append(f"El acum. interanual termina en {a25_fin.year}, se esperaba {fecha.year - 1}.")
    if (a25_fin.month, a25_fin.day) != (fecha.month, fecha.day):
        errores.append(f"El día de corte no coincide: día {fecha.day}/{fecha.month} vs "
                       f"acum. hasta {a25_fin.day}/{a25_fin.month}.")
    if errores:
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a", encoding="utf-8") as f:
                f.write("send=false\n")
        print("[ERROR DE VALIDACIÓN] No se envía el mail:")
        for e in errores:
            print("  - " + e)
        sys.exit(1)

    state = load_json(BASE / "data" / "acumulado.json", default={"month": None, "last_date": None, "acumulado": {}})
    mes_actual = fecha.strftime("%Y-%m")

    # Anti doble-conteo
    if state.get("last_date") == fecha.isoformat():
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a", encoding="utf-8") as f:
                f.write("send=false\n")
        print(f"[SKIP] La fecha {fecha} ya fue procesada. No se suma de nuevo ni se envía mail.")
        sys.exit(0)

    # Reinicio mensual
    if state.get("month") != mes_actual:
        print(f"[MES NUEVO] {state.get('month')} -> {mes_actual}. Acumulado reiniciado a cero.")
        acum_prev = {b: 0.0 for b in BRANCH_ORDER}
    else:
        acum_prev = state.get("acumulado", {})

    acum26 = {b: round(acum_prev.get(b, 0.0) + venta_dia[b], 2) for b in BRANCH_ORDER}

    rows = []
    for b in BRANCH_ORDER:
        a26 = acum26[b]
        a25 = acum25[b]
        comparable = a25 > 0
        diff = a26 - a25 if comparable else 0.0
        pct = (diff / a25 * 100) if comparable else 0.0
        rows.append({"branch": b, "dia": venta_dia[b], "a26": a26, "a25": a25,
                     "diff": diff, "pct": pct, "cmp": comparable})

    def bloque_tot(grupo):
        # Los montos de venta (día y acum 2026) muestran TODOS los locales -> es la
        # venta real de la operación. Pero la VARIACIÓN se calcula manzana-manzana:
        # sólo los locales con histórico 2025, en ambos años. Así los locales nuevos
        # (Madrid, Alicante, Granada antes de abrir) no distorsionan el %.
        sel = [r for r in rows if r["branch"] in grupo]
        comp = [r for r in sel if r["cmp"]]
        dia = sum(r["dia"] for r in sel)          # todos
        a26 = sum(r["a26"] for r in sel)          # todos (acumulado real)
        a26_c = sum(r["a26"] for r in comp)       # sólo comparables, para el %
        a25 = sum(r["a25"] for r in comp)         # sólo comparables
        diff = a26_c - a25
        return {"dia": dia, "a26": a26, "a25": a25,
                "diff": diff, "pct": (diff / a25 * 100) if a25 else 0.0,
                "cmp": a25 > 0}

    totals = bloque_tot(BRANCH_ORDER)
    propias = bloque_tot(PROPIAS)
    franquicias = bloque_tot(FRANQUICIAS)

    new_state = {"month": mes_actual, "last_date": fecha.isoformat(), "acumulado": acum26}

    # Gráficos: sólo locales comparables (los nuevos no tienen barra 2025)
    from charts import build_charts
    mes = MES_CORTO[fecha.month]
    a26_lbl = f"Acum. {mes}/{str(fecha.year)[2:]}"
    a25_lbl = f"Acum. {mes}/{str(fecha.year - 1)[2:]}"
    comp_rows = [r for r in rows if r["cmp"]]
    chart_paths = build_charts(comp_rows, totals, mes, a26_lbl, a25_lbl, out_dir=str(BASE / "charts"))

    html = render_html(fecha, rows, totals, propias, franquicias)
    return html, new_state, fecha, totals, chart_paths


def render_html(fecha, rows, totals, propias, franquicias):
    mes = MES_CORTO[fecha.month]
    anio_corto = str(fecha.year)[2:]
    anio_ant = str(fecha.year - 1)[2:]
    fecha_larga = f"{DIAS_ES[fecha.weekday()]} {fecha.day} DE {MESES_ES[fecha.month]} DE {fecha.year}"
    a26_lbl = f"ACUM. {mes.upper()}/{anio_corto}"
    a25_lbl = f"ACUM. {mes.upper()}/{anio_ant}"

    def chip(pct, diff, cmp=True):
        if not cmp:
            return ('<span style="display:inline-block;background:#eeeeee;color:#9a9a9a;'
                    'font-weight:700;font-size:12px;padding:3px 9px;border-radius:20px;white-space:nowrap;">—</span>'
                    '<div style="color:#9a9a9a;font-size:11px;margin-top:3px;">sin comp.</div>')
        up = pct >= 0
        col = "#1a7d2e" if up else "#c62828"
        bg = "#eaf5ec" if up else "#fbecec"
        s = "+" if up else ""
        dd = f"(+{money(diff)})" if diff >= 0 else f"({money(diff)})"
        return (f'<span style="display:inline-block;background:{bg};color:{col};'
                f'font-weight:700;font-size:12px;padding:3px 9px;border-radius:20px;white-space:nowrap;">'
                f'{s}{pct_txt(pct)}</span>'
                f'<div style="color:{col};font-size:11px;margin-top:3px;">{dd}</div>')

    def fila(r, zebra):
        a25_cell = money(r['a25']) if r['cmp'] else '<span style="color:#c9c9c9;">—</span>'
        return f"""
        <tr style="background:{zebra};">
          <td style="padding:14px 18px;font-weight:700;color:#111111;font-size:14px;">{r['branch']}</td>
          <td style="padding:14px 12px;text-align:right;color:#111111;font-size:14px;">{money(r['dia'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#111111;font-weight:700;font-size:14px;">{money(r['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#9a9a9a;font-size:14px;">{a25_cell}</td>
          <td style="padding:14px 18px;text-align:right;">{chip(r['pct'], r['diff'], r['cmp'])}</td>
        </tr>"""

    def encabezado_grupo(nombre):
        return f"""
        <tr style="background:#eef1f6;">
          <td colspan="5" style="padding:9px 18px;color:#6B1F2A;font-size:10px;font-weight:800;letter-spacing:2px;">{nombre}</td>
        </tr>"""

    def fila_subtotal(nombre, s):
        return f"""
        <tr style="background:#f5f5f5;border-top:1px solid #e2e2e2;">
          <td style="padding:14px 18px;font-weight:800;color:#6B1F2A;font-size:13px;">{nombre}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#6B1F2A;font-size:13px;">{money(s['dia'])}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#6B1F2A;font-size:13px;">{money(s['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#9a7a80;font-size:13px;">{money(s['a25'])}</td>
          <td style="padding:14px 18px;text-align:right;">{chip(s['pct'], s['diff'], s['cmp'])}</td>
        </tr>"""

    by_name = {r["branch"]: r for r in rows}
    cuerpo = encabezado_grupo("PROPIAS")
    for i, b in enumerate(PROPIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Propias", propias)
    cuerpo += encabezado_grupo("FRANQUICIAS")
    for i, b in enumerate(FRANQUICIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Franquicias", franquicias)

    n = len(BRANCH_ORDER)
    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eeeeee;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eeeeee;">
<tr><td align="center" style="padding:24px 12px;">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:600px;max-width:600px;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,0.12);">
  <tr><td style="background:#000000;padding:38px 32px 32px 32px;text-align:center;">
    <img src="cid:logo" alt="Lucciano's" width="190" style="display:block;margin:0 auto;max-width:190px;height:auto;">
    <div style="color:#bdbdbd;font-size:12px;letter-spacing:4px;margin-top:18px;">REPORTE DE VENTAS DIARIO · EUROPA</div>
    <div style="color:#ffffff;font-size:13px;font-weight:700;letter-spacing:2px;margin-top:14px;">{fecha_larga}</div>
  </td></tr>
  <tr><td style="padding:30px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;">CONSOLIDADO · {n} SUCURSALES</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px;">
      <tr>
        <td width="33%" style="padding-right:7px;vertical-align:top;">
          <div style="background:#111111;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">VENTA DEL DÍA</div>
            <div style="color:#ffffff;font-size:20px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['dia'])}</div>
            <div style="color:#777777;font-size:11px;margin-top:6px;">{n} sucursales</div>
          </div>
        </td>
        <td width="34%" style="padding:0 7px;vertical-align:top;">
          <div style="background:#111111;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">{a26_lbl}</div>
            <div style="color:#ffffff;font-size:20px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a26'])}</div>
            <div style="color:#777777;font-size:11px;margin-top:6px;">mes en curso</div>
          </div>
        </td>
        <td width="33%" style="padding-left:7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">{a25_lbl}</div>
            <div style="color:#111111;font-size:20px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a25'])}</div>
            <div style="margin-top:8px;">{chip(totals['pct'], totals['diff'], totals['cmp'])}</div>
          </div>
        </td>
      </tr>
    </table>
  </td></tr>
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="background:#fafafa;border-radius:12px;padding:20px 22px;">
      <div style="color:#9a9a9a;font-size:11px;letter-spacing:2px;margin-bottom:6px;">AVANCE DEL MES vs AÑO ANTERIOR</div>
      <img src="cid:progreso" alt="Avance del mes" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
    </div>
  </td></tr>
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:12px;">COMPARATIVO POR SUCURSAL · {anio_corto} vs {anio_ant}</div>
    <img src="cid:comparativo" alt="Comparativo por sucursal" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
  </td></tr>
  <tr><td style="padding:24px 32px 36px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:14px;">DETALLE POR SUCURSAL</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-radius:12px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,0.06);">
      <thead>
        <tr style="background:#111111;">
          <th style="padding:13px 18px;text-align:left;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">SUCURSAL</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">DÍA</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a26_lbl}</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a25_lbl}</th>
          <th style="padding:13px 18px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">VARIACIÓN</th>
        </tr>
      </thead>
      <tbody>{cuerpo}
        <tr style="background:#111111;">
          <td style="padding:17px 18px;font-weight:800;color:#ffffff;font-size:14px;">TOTAL GENERAL</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['dia'])}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a26'])}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a25'])}</td>
          <td style="padding:17px 18px;text-align:right;">{chip(totals['pct'], totals['diff'], totals['cmp'])}</td>
        </tr>
      </tbody>
    </table>
  </td></tr>
  <tr><td style="background:#000000;padding:20px 32px;text-align:center;">
    <div style="color:#777777;font-size:11px;letter-spacing:1px;">LUCCIANO'S EUROPA · Reporte automático generado el {fecha.strftime('%d/%m/%Y')}</div>
    <div style="color:#5a5a5a;font-size:10px;margin-top:6px;">Variación calculada sobre sucursales con histórico 2025. Las de apertura 2026 se muestran sin comparación (—).</div>
  </td></tr>
</table>
</td></tr>
</table>
</body>
</html>"""
    return html


if __name__ == "__main__":
    html, new_state, fecha, totals, chart_paths = build_report()
    (BASE / "preview.html").write_text(html, encoding="utf-8")
    save_json(BASE / "data" / "acumulado.json", new_state)

    subject = f"Reporte Ventas {fecha.strftime('%d/%m/%Y')} - Lucciano's Europa"
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write("send=true\n")
            f.write(f"subject={subject}\n")
            f.write(f"date={fecha.isoformat()}\n")
    print(f"OK - fecha {fecha} | venta día {money(totals['dia'])} | acum26 {money(totals['a26'])} | var {pct_txt(totals['pct'])}")
