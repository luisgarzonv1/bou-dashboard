"""
BOU Sales Command Center — Streamlit + Plotly
================================================
Lee EN VIVO data/daily-log.json del repo luisgarzonv1/bou-dashboard (misma fuente
que ya actualizan las tareas programadas bou-dashboard-diario / bou-dashboard-cierre).
No modifica esa automatizacion -- esto es solo una capa de presentacion nueva.

Sistema de diseno tomado de bou-command-center-v2.jsx (mockup de referencia, datos
ficticios). Los datos aqui son REALES, leidos del JSON en vivo.

Correr localmente:
    pip install -r requirements.txt
    streamlit run bou_command_center.py

Publicar gratis (Streamlit Community Cloud): subir este archivo + requirements.txt
a un repo de GitHub, share.streamlit.io -> New app -> apuntar a este archivo.
"""

import requests
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from datetime import date, timedelta

DAILY_LOG_URL = "https://raw.githubusercontent.com/luisgarzonv1/bou-dashboard/main/data/daily-log.json"

# Categorías de BOU Logistica (proyecto aparte, ver Productos de BOU Logistica.xlsx) que
# Contífico mezcla dentro de inventario_por_artista porque comparten el mismo catálogo.
# Deben excluirse siempre de cualquier vista de "inventario por artista" de BOU Entertainment.
BOU_LOGISTICA_CATEGORIAS = {
    "NOVESS", "ECOM GLOBAL", "GRUPOESHOP", "FSA NATURAL", "GREEN LINE",
    "LSJ SAS", "Mindco", "SADE BUSINESS PERU S.A.C.", "SANATE FILTROS S A",
}

# ---------- Paleta (bou-command-center-v2.jsx, adaptada a fondo claro a pedido de Luis) ----------
# Los 6 acentos son EXACTAMENTE los del .jsx. Fondo/paneles/texto se invirtieron a modo claro;
# se agregan variantes "_INK" (mismo tono, más oscuro) solo para texto pequeño sobre blanco,
# donde el acento original no tiene suficiente contraste para ser legible.
BG = "#F6F6FB"
PANEL = "#FFFFFF"
BORDER = "#E4E4EE"
TEXT = "#15161C"
MUTED = "#6B6F80"
GOLD = "#FFB648"    # Ecuador / KPI principal
MINT = "#4FE3C1"    # Colombia / positivo
BLUE = "#7C9EFF"
CORAL = "#FF6B5C"   # alertas
LILAC = "#C792EA"
CYAN = "#67D4FF"
GOLD_INK = "#9A6414"    # para texto/eyebrow sobre blanco (mismo dorado, oscurecido)
CORAL_INK = "#C2392A"   # para texto de alerta sobre blanco

STATUS_CFG = {
    "ok":   {"icon": "✅", "color": MINT,  "text": "Conciliado"},
    "warn": {"icon": "⚠️", "color": GOLD, "text": "Pendiente"},
    "down": {"icon": "\U0001F534", "color": CORAL, "text": "Bloqueado"},
}

KOMMO_STAGE_ORDER_EC = ["Incoming leads", "Contacto inicial", "Nueva consulta", "ATENCION AGENTE", "leads frios (no responde)", "leads Tibios (interacción)", "leads Calientes (datos bancarios)", "Por registrar y FACTURAR", "Por entregar", "Cliente Registrado", "CAMBIOS Y DEVOLUCIONES", "Falta stock", "ERROR ENVIAR MENSAJE", "CONFIRMADOS EVENTOS", "BOU SESSION", "Sorteos y concursos", "Servicios Adicionales", "Reacciones y Comentarios", "R.Comentarios Revisados", "Colaboraciones", "SPAM", "Leads ganados", "Leads perdidos"]
KOMMO_STAGE_ORDER_COL = ["Incoming leads", "Contacto inicial", "Nueva consulta", "ATENCION AGENTE", "leads frios (no responde)", "leads Tibios (interacción)", "leads Calientes (datos bancarios)", "Por registrar ", "Por entregar", "CAMBIOS Y DEVOLUCIONES", "ERROR ENVIAR MENSAJE", "CONFIRMADOS", "Logrado con éxito", "Ventas Perdidos"]

COBRO_LABELS = {"TC": "Tarjeta de credito", "TRA": "Transferencia", "CH": "Cheque", "EF": "Efectivo", "SIN_DATO": "Sin dato"}
COBRO_COLORS = {"TC": GOLD, "TRA": MINT, "CH": BLUE, "EF": LILAC, "SIN_DATO": MUTED}

st.set_page_config(page_title="BOU Sales Command Center", layout="wide", page_icon="\U0001F4CA")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&display=swap');
html, body, [class*="css"] {{ font-family: 'Inter', ui-sans-serif, system-ui; }}
.stApp {{ background-color: {BG}; color: {TEXT}; }}
.block-container {{ padding-top: 2rem; max-width: 1300px; }}
.bcc-eyebrow {{
  display: inline-block; font-size: 11px; text-transform: uppercase; letter-spacing: 0.2em;
  color: #7A4E0E; background: #FFF1DC; border-radius: 6px; padding: 2px 8px; font-weight: 700; margin-bottom: 6px;
}}
.bcc-title {{ font-size: 32px; font-weight: 900; letter-spacing: -0.02em; margin: 0; line-height: 1.1; color: {TEXT}; }}
.bcc-subtitle {{ color: {MUTED}; font-size: 13px; margin-top: 6px; }}
.bcc-card {{
  position: relative; background: {PANEL}; border: 1px solid {BORDER}; border-radius: 16px;
  padding: 18px 18px 16px; overflow: hidden; box-shadow: 0 1px 3px rgba(20,20,40,0.06);
}}
.bcc-card::before {{ content: ""; position: absolute; top:0; left:0; right:0; height: 3px; background: var(--accent, {GOLD}); }}
.bcc-card .lbl {{ font-size: 11px; text-transform: uppercase; letter-spacing: 0.14em; color: {MUTED}; font-weight: 600; }}
.bcc-card .icon {{ float: right; font-size: 15px; }}
.bcc-card .val {{ font-size: 26px; font-weight: 900; color: {TEXT}; font-variant-numeric: tabular-nums; letter-spacing: -0.02em; margin-top: 10px; }}
.bcc-card .sub {{ font-size: 12px; color: {MUTED}; margin-top: 4px; }}
.bcc-badge {{
  display:flex; align-items:center; gap: 10px; background: {PANEL}; border: 1px solid {BORDER};
  border-radius: 12px; padding: 10px 14px; margin-bottom: 8px; box-shadow: 0 1px 2px rgba(20,20,40,0.04);
}}
.bcc-badge .ic {{ font-size: 15px; }}
.bcc-badge .lbl {{ font-size: 12px; font-weight: 700; color: {TEXT}; }}
.bcc-badge .sub {{ font-size: 10px; color: {MUTED}; }}
.bcc-panel-title {{ font-size: 14px; font-weight: 700; color: {TEXT}; margin-bottom: 4px; }}
.bcc-panel-note {{ font-size: 11px; color: {MUTED}; margin-bottom: 10px; }}
.bcc-note-box {{ background: {PANEL}; border: 1px solid {BORDER}; border-radius: 12px; padding: 12px 16px; font-size: 12px; color: {MUTED}; }}
.bcc-warn-box {{ background: #FFF6E9; border: 1px solid #F5DCA8; color: {GOLD_INK}; border-radius: 12px; padding: 12px 16px; font-size: 12px; }}
.bcc-dev-box {{ background: #FAFAFC; border: 1px dashed #C7C9D6; border-radius: 12px; padding: 14px 16px; font-size: 12px; color: {MUTED}; }}
</style>
""", unsafe_allow_html=True)


# ---------- Datos ----------
@st.cache_data(ttl=300)
def load_data():
    r = requests.get(DAILY_LOG_URL, timeout=20)
    r.raise_for_status()
    return r.json()


def kpi_card(label, value, sub=None, accent=GOLD, icon=""):
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    st.markdown(
        f'<div class="bcc-card" style="--accent:{accent};">'
        f'<span class="icon">{icon}</span><span class="lbl">{label}</span>'
        f'<div class="val">{value}</div>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def badge(status, label, sub):
    cfg = STATUS_CFG[status]
    st.markdown(
        f'<div class="bcc-badge"><span class="ic">{cfg["icon"]}</span>'
        f'<div><div class="lbl">{label}</div><div class="sub">{cfg["text"]} · {sub}</div></div></div>',
        unsafe_allow_html=True,
    )


def panel_header(title, note=None):
    st.markdown(f'<div class="bcc-panel-title">{title}</div>', unsafe_allow_html=True)
    if note:
        st.markdown(f'<div class="bcc-panel-note">{note}</div>', unsafe_allow_html=True)


def sin_dato(v):
    """Regla de veracidad: null/None nunca se grafica como si fuera un valor real."""
    return v is None


def fmt_or_missing(v, fmt="{:,.0f}"):
    return "sin dato" if sin_dato(v) else fmt.format(v)


def plotly_dark_layout(fig, height=280, **kwargs):
    """Nombre histórico de la función (se mantiene para no tocar cada call site) —
    desde el cambio a tema claro, en realidad aplica el layout de tema CLARO."""
    fig.update_layout(
        template="plotly_white", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
        font=dict(color=TEXT, size=11), height=height,
        margin=dict(t=20, b=30, l=40, r=20),
        legend=dict(orientation="h", y=1.15, font=dict(color=MUTED)),
        xaxis=dict(gridcolor=BORDER, zeroline=False),
        yaxis=dict(gridcolor=BORDER, zeroline=False),
        **kwargs,
    )
    return fig


def sum_field(rows, field):
    vals = [r.get(field) for r in rows if r.get(field) is not None]
    return sum(vals) if vals else None


def weighted_rate(numer_rows, denom_rows=None):
    """Suma numeradores/denominadores en vez de promediar tasas diarias (mas correcto)."""
    pass


log = load_data()
days_all = sorted(log.get("daily", []), key=lambda d: d["date"], reverse=True)
last_day = days_all[0]
today_str = last_day["date"]

# ---------- Header ----------
st.markdown(
    '<div class="bcc-warn-box">⚠️ Dashboard conectado a datos reales de '
    '<code>daily-log.json</code> (repo luisgarzonv1/bou-dashboard). Sigue las reglas de '
    'veracidad: campos en 0/null se muestran como "sin dato", nunca se suma USD+COP sin '
    'mostrar la TRM, y los campos nuevos sin metodologia documentada se marcan "en desarrollo".</div>',
    unsafe_allow_html=True,
)
st.write("")
c1, c2 = st.columns([3, 1])
with c1:
    st.markdown('<div class="bcc-eyebrow">BOU Entertainment</div>', unsafe_allow_html=True)
    st.markdown('<div class="bcc-title">Sales Command Center</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="bcc-subtitle">Ecuador + Colombia · dato conciliado contra la fuente de facturación</div>',
        unsafe_allow_html=True,
    )
with c2:
    trm = st.number_input("TRM ref. (COP por USD, manual)", min_value=1000, max_value=10000, value=4050, step=10)
    st.caption("Ingresado manualmente — no es un valor de mercado en vivo.")

st.write("")

tab_resumen, tab_ec, tab_col, tab_mkt, tab_conc = st.tabs(
    ["\U0001F6E1️ Resumen", "\U0001F310 Ecuador", "\U0001F310 Colombia", "\U0001F4E2 Marketing", "✅ Conciliación"]
)

# ---------- Selector de fecha (aplica a las primeras 4 pestanas) ----------
def date_range_selector(key_prefix):
    preset = st.radio(
        "Rango", ["Hoy", "7D", "30D", "Este mes", "Personalizado"],
        horizontal=True, key=f"{key_prefix}_preset", label_visibility="collapsed",
    )
    last_date = pd.to_datetime(today_str).date()
    if preset == "Hoy":
        start, end = last_date, last_date
    elif preset == "7D":
        start, end = last_date - timedelta(days=6), last_date
    elif preset == "30D":
        start, end = last_date - timedelta(days=29), last_date
    elif preset == "Este mes":
        start, end = last_date.replace(day=1), last_date
    else:
        c1, c2 = st.columns(2)
        start = c1.date_input("Desde", value=last_date - timedelta(days=13), key=f"{key_prefix}_start")
        end = c2.date_input("Hasta", value=last_date, key=f"{key_prefix}_end")
    return start.isoformat(), end.isoformat()


def rows_in_range(start, end):
    return [d for d in days_all if start <= d["date"] <= end]


def parcial_banner(rows, end):
    """Regla de veracidad: si el rango incluye el corte parcial de hoy, indicarlo visualmente."""
    today_row = next((r for r in rows if r["date"] == end), None)
    if today_row and today_row.get("parcial"):
        st.markdown(
            f'<div class="bcc-warn-box">⚠️ El día <b>{end}</b> aún está en '
            'corte parcial (10:30am) — los números de hoy pueden subir con el corte de '
            'cierre (6:00pm).</div>',
            unsafe_allow_html=True,
        )
        st.write("")


def mark_parcial_day(fig, rows, end):
    """Regla de veracidad #5: marcar visualmente el día en curso si sigue en corte parcial
    (borde/textura distinta) sobre un gráfico de tendencia diaria."""
    today_row = next((r for r in rows if r["date"] == end), None)
    if today_row and today_row.get("parcial"):
        fig.add_vline(x=end, line_width=2, line_dash="dot", line_color=CORAL)
        fig.add_annotation(x=end, y=1, yref="paper", yshift=10, showarrow=False,
                            text="hoy · parcial", font=dict(color=CORAL, size=10))
    return fig


# ============================================================
# RESUMEN
# ============================================================
with tab_resumen:
    start, end = date_range_selector("resumen")
    rows = rows_in_range(start, end)
    parcial_banner(rows, end)

    ec_total = sum_field(rows, "prestashop_total_ec_usd")
    col_total_cop = sum_field(rows, "prestashop_total_col_cop")
    col_total_usd = (col_total_cop / trm) if col_total_cop is not None else None
    consolidado = (ec_total or 0) + (col_total_usd or 0) if (ec_total is not None or col_total_usd is not None) else None

    ec_pedidos = sum_field(rows, "prestashop_pedidos_ec")
    col_pedidos = sum_field(rows, "prestashop_pedidos_col")
    pedidos_total = (ec_pedidos or 0) + (col_pedidos or 0) if (ec_pedidos is not None or col_pedidos is not None) else None

    fec = last_day.get("kommo_funnel_ec") or {}
    fcol = last_day.get("kommo_funnel_col") or {}
    won = (fec.get("won") or 0) + (fcol.get("won") or 0)
    lost = (fec.get("lost") or 0) + (fcol.get("lost") or 0)
    conv_rate = (won / (won + lost) * 100) if (won + lost) > 0 else None

    lider = None
    if ec_total is not None and col_total_usd is not None:
        lider = "Ecuador" if ec_total >= col_total_usd else "Colombia"

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card("Consolidado canal web (USD)", fmt_or_missing(consolidado, "${:,.0f}"), f"PrestaShop EC+COL · TRM ref. {trm}/USD", GOLD, "\U0001F4B5")
    with k2:
        kpi_card("Pedidos totales (web)", fmt_or_missing(pedidos_total, "{:,.0f}"), f"EC {ec_pedidos or 0} · COL {col_pedidos or 0}", MINT, "\U0001F6CD️")
    with k3:
        kpi_card("% Leads convertidos", fmt_or_missing(conv_rate, "{:.0f}%"), "ventana móvil 30 días (Kommo)", BLUE, "\U0001F4CA")
    with k4:
        kpi_card("Canal web líder", lider or "sin dato", "PrestaShop, USD equiv.", LILAC, "\U0001F4C8")

    st.caption(
        "Este resumen usa PrestaShop (canal web) porque es la única fuente con dato diario confiable en ambos "
        "países. Las fuentes primarias declaradas son Contífico para Ecuador y la Matriz de Ventas Sheets para "
        "Colombia — ver el detalle y por qué aún no están automatizadas en las pestañas Ecuador y Colombia."
    )

    st.write("")
    panel_header("Ingresos por país", "USD equivalente, referencial (TRM manual)")
    dfrows = pd.DataFrame(rows).sort_values("date")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=dfrows["date"], y=dfrows["prestashop_total_ec_usd"], name="Ecuador (USD)", line=dict(color=GOLD, width=2)))
    if "prestashop_total_col_cop" in dfrows:
        fig.add_trace(go.Scatter(x=dfrows["date"], y=dfrows["prestashop_total_col_cop"] / trm, name="Colombia (USD equiv.)", line=dict(color=MINT, width=2)))
    fig = mark_parcial_day(fig, rows, end)
    st.plotly_chart(plotly_dark_layout(fig), use_container_width=True)

    st.write("")
    st.markdown('<div class="bcc-panel-title">Estado de fuentes</div>', unsafe_allow_html=True)
    b1, b2, b3 = st.columns(3)
    with b1:
        ctf_caida = last_day.get("contifico_fuente_caida")
        badge("down" if ctf_caida else "ok", "Contífico — fuente primaria EC",
              "0 docs — bloqueo firma electrónica SRI" if ctf_caida else "recibiendo documentos")
        badge("ok", "PrestaShop EC — referencia/respaldo", "filtrado por current_state válido")
        badge("ok", "Kommo (EC + COL)", "conteo oficial vía API")
    with b2:
        badge("ok", "Matriz Sheets — fuente primaria COL", "conectada via CSV export, celdas combinadas ya no bloquean")
        badge("ok", "PrestaShop COL — complementario", "activo, se usa mientras se resuelve el Sheet")
    with b3:
        ga4_ok = last_day.get("ga4_usuarios_ec") is not None or last_day.get("ga4_usuarios_col") is not None
        badge("ok" if ga4_ok else "warn", "Windsor.ai (Meta + GA4)", "datos del último corte" if ga4_ok else "sin dato en el último corte")
        badge("ok", "Inventario por artista", f"corte de mañana · último dato {last_day.get('date') if last_day.get('inventario_total') else 'sin dato'}")

    nota = last_day.get("nota_verificacion")
    if nota:
        st.caption(f"Nota de verificación del corte más reciente ({last_day['date']}): {nota}")

# ============================================================
# ECUADOR
# ============================================================
with tab_ec:
    start, end = date_range_selector("ec")
    rows = rows_in_range(start, end)
    parcial_banner(rows, end)

    ec_total = sum_field(rows, "prestashop_total_ec_usd")
    ec_pedidos = sum_field(rows, "prestashop_pedidos_ec")
    ticket = (ec_total / ec_pedidos) if (ec_total and ec_pedidos) else None
    ctf_docs = sum_field(rows, "contifico_documentos")
    ctf_total = sum_field(rows, "contifico_total_usd")
    leads_ec = sum_field(rows, "kommo_leads_ec")

    # Jerarquía de fuente EC (definida por Luis en el chat "Dashboard de ventas con filtro de
    # fecha"): Contífico factura = fuente de verdad de venta consumada. Todo pedido pagado de
    # PrestaShop DEBE tener su factura correspondiente en Contífico -- el cruce real es por
    # número de pedido (referencia guardada en el documento de Contífico), algo que daily-log.json
    # todavía no trae a nivel de detalle (solo totales agregados por día). Mientras ese cruce
    # exacto no esté implementado, se muestra una brecha aproximada: pedidos PrestaShop vs.
    # documentos Contífico en el mismo rango. Un pedido sin factura no es una venta perdida, es
    # un pendiente de facturar. Kommo complementa con contexto de leads, no de ventas.
    ticket_base = ctf_total if (ctf_total and ctf_total > 0) else ec_total
    ticket_pedidos_base = ctf_docs if (ctf_total and ctf_total > 0) else ec_pedidos
    ticket = (ticket_base / ticket_pedidos_base) if (ticket_base and ticket_pedidos_base) else None
    brecha_docs = (ec_pedidos - ctf_docs) if (ec_pedidos is not None and ctf_docs is not None) else None

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card("Facturación Contífico (fuente primaria)", fmt_or_missing(ctf_total, "${:,.0f}"),
                  f"{fmt_or_missing(ctf_docs)} documentos" if (ctf_total or ctf_docs) else "bloqueado en el rango — ver PrestaShop al lado",
                  GOLD, "\U0001F9FE")
    with k2:
        kpi_card("PrestaShop EC (pedidos web)", fmt_or_missing(ec_total, "${:,.0f}"), f"{fmt_or_missing(ec_pedidos)} pedidos pagados · deben conciliar con Contífico", MINT, "\U0001F4B5")
    with k3:
        kpi_card("Ticket promedio", fmt_or_missing(ticket, "${:,.2f}"), "Contífico" if ticket_base is ctf_total and ctf_total else "PrestaShop (respaldo)", BLUE, "\U0001F4C8")
    with k4:
        kpi_card("Leads Kommo EC", fmt_or_missing(leads_ec, "{:,.0f}"), "contexto de leads, no de ventas", LILAC, "\U0001F465")

    st.markdown(
        '<div class="bcc-note-box">Contífico es la fuente primaria de ventas EC — todo pedido pagado de PrestaShop '
        'debe tener su factura correspondiente en Contífico. El cruce exacto es por número de pedido, guardado como '
        'referencia dentro del documento de Contífico; ese detalle aún no está en <code>daily-log.json</code> (solo '
        'trae totales agregados por día), así que hoy solo se puede comparar en agregado: '
        + (f'<b>{brecha_docs} pedido(s) PrestaShop de diferencia frente a documentos Contífico</b> en el rango — no es una venta perdida, es un posible pendiente de facturar (comparación aproximada, no cruce por número de pedido).'
           if brecha_docs is not None and brecha_docs != 0
           else 'los conteos coinciden en el rango seleccionado (comparación aproximada, no cruce por número de pedido).')
        + ' Kommo complementa con leads, no con ventas. No se suman Contífico + PrestaShop entre sí.</div>',
        unsafe_allow_html=True,
    )
    st.write("")

    col1, col2 = st.columns([2, 1])
    with col1:
        panel_header("Facturación diaria", "USD, Contífico (fuente primaria)")
        dfrows = pd.DataFrame(rows).sort_values("date")
        fig = go.Figure(go.Scatter(x=dfrows["date"], y=dfrows["contifico_total_usd"], fill="tozeroy",
                                    line=dict(color=GOLD, width=2), fillcolor="rgba(255,182,72,0.28)"))
        fig = mark_parcial_day(fig, rows, end)
        st.plotly_chart(plotly_dark_layout(fig, height=260), use_container_width=True)
    with col2:
        panel_header("Mensajería Kommo EC")
        entrantes = sum_field(rows, "kommo_msj_entrantes_ec")
        salientes = sum_field(rows, "kommo_msj_salientes_ec")
        lapso = last_day.get("kommo_lapso_medio_ec_s")
        st.markdown(f'<div class="bcc-card" style="--accent:{BLUE};"><span class="lbl">Entrantes / Salientes</span>'
                     f'<div class="val" style="font-size:20px;">{fmt_or_missing(entrantes)} / {fmt_or_missing(salientes)}</div>'
                     f'<div class="sub">Lapso resp. {(str(int(lapso//60))+"m") if lapso else "sin dato"}</div></div>', unsafe_allow_html=True)

    st.write("")
    panel_header("Ingresos web PrestaShop EC", "referencia/respaldo, no se suma a Contífico")
    dfrows_ec = pd.DataFrame(rows).sort_values("date")
    fig_ec2 = go.Figure(go.Scatter(x=dfrows_ec["date"], y=dfrows_ec["prestashop_total_ec_usd"], line=dict(color=MINT, width=1.5, dash="dot")))
    fig_ec2 = mark_parcial_day(fig_ec2, rows, end)
    st.plotly_chart(plotly_dark_layout(fig_ec2, height=180), use_container_width=True)

    st.write("")
    st.write("")
    panel_header("Ticket promedio (tendencia)", "Contífico si hay dato, si no PrestaShop EC")
    dfrows_tk = pd.DataFrame(rows).sort_values("date")
    dfrows_tk["ticket_dia"] = dfrows_tk.apply(
        lambda r: (r["contifico_total_usd"] / r["contifico_documentos"]) if r.get("contifico_total_usd") and r.get("contifico_documentos")
        else ((r["prestashop_total_ec_usd"] / r["prestashop_pedidos_ec"]) if r.get("prestashop_total_ec_usd") and r.get("prestashop_pedidos_ec") else None),
        axis=1,
    )
    fig_tk = go.Figure(go.Scatter(x=dfrows_tk["date"], y=dfrows_tk["ticket_dia"], mode="lines+markers", line=dict(color=CORAL, width=2)))
    fig_tk = mark_parcial_day(fig_tk, rows, end)
    st.plotly_chart(plotly_dark_layout(fig_tk, height=220), use_container_width=True)

    st.write("")
    panel_header("Método de cobro (BE, Contífico)", "cobros[].forma_cobro — solo documentos clasificados como BE, netea facturas con nota de crédito")
    cobro_totales = {}
    for r in rows:
        mc = r.get("contifico_metodo_cobro") or {}
        for k, v in mc.items():
            cobro_totales[k] = cobro_totales.get(k, 0) + (v or 0)
    if cobro_totales:
        labels = [COBRO_LABELS.get(k, k) for k in cobro_totales.keys()]
        colors = [COBRO_COLORS.get(k, MUTED) for k in cobro_totales.keys()]
        fig = go.Figure(go.Pie(labels=labels, values=list(cobro_totales.values()), hole=0.55,
                                marker=dict(colors=colors), textinfo="label+percent"))
        st.plotly_chart(plotly_dark_layout(fig, height=280, showlegend=False), use_container_width=True)
    else:
        st.markdown('<div class="bcc-note-box">sin dato de método de cobro en este rango</div>', unsafe_allow_html=True)

    panel_header("Funnel Kommo EC", "ventana móvil 30 días · orden real del pipeline")
    fec = last_day.get("kommo_funnel_ec") or {}
    by_status = fec.get("byStatus") or {}
    if by_status:
        total_funnel_ec = sum(by_status.values())
        items_ec = sorted(by_status.items(), key=lambda x: KOMMO_STAGE_ORDER_EC.index(x[0]) if x[0] in KOMMO_STAGE_ORDER_EC else 999)
        sdf = pd.DataFrame(items_ec, columns=["etapa", "leads"])
        sdf["pct"] = sdf["leads"].apply(lambda v: f"{(v/total_funnel_ec*100):.0f}%" if total_funnel_ec else "")
        fig = go.Figure(go.Bar(x=sdf["etapa"], y=sdf["leads"], text=sdf["pct"], textposition="outside", marker_color=BLUE))
        st.plotly_chart(plotly_dark_layout(fig, height=280), use_container_width=True)
    else:
        st.markdown('<div class="bcc-note-box">sin dato de funnel para Ecuador en este corte</div>', unsafe_allow_html=True)

    st.write("")
    panel_header("Inventario vendible por artista", "Bou Entertainment · solo corte de mañana")
    inv_day = next((d for d in days_all if d.get("inventario_total", {}).get("cantidad")), None)
    if inv_day:
        st.caption(f"Último dato: {inv_day['date']} (corte de mañana — no refleja ventas de la tarde de ese día)")
        art_raw = inv_day.get("inventario_por_artista") or {}
        # Excluir categorías de BOU Logistica (proyecto aparte, mismo catálogo Contífico)
        art = {k: v for k, v in art_raw.items() if k not in BOU_LOGISTICA_CATEGORIAS}
        excluidos_n = len(art_raw) - len(art)
        adf = pd.DataFrame([{"artista": k, "monto": v["monto_pvp"], "cantidad": v["cantidad"]} for k, v in art.items()]).sort_values("monto")
        fig = go.Figure(go.Bar(x=adf["monto"], y=adf["artista"], orientation="h", marker_color=CYAN,
                                customdata=adf["cantidad"], hovertemplate="%{y}<br>$%{x:,.2f} PVP<br>%{customdata} uds<extra></extra>"))
        st.plotly_chart(plotly_dark_layout(fig, height=max(360, 20 * len(adf))), use_container_width=True)
        if excluidos_n:
            st.caption(f"Se excluyeron {excluidos_n} categorías de BOU Logistica (proyecto aparte, comparten catálogo en Contífico).")
        # inventario_alertas_pvp es una lista de textos ya redactados por la tarea programada;
        # se descartan las líneas que hablan de categorías de BOU Logistica (no son artistas BE)
        alertas = [a for a in (inv_day.get("inventario_alertas_pvp") or [])
                   if not any(a.startswith(cat) for cat in BOU_LOGISTICA_CATEGORIAS)]
        if alertas:
            st.markdown(
                '<div class="bcc-warn-box">⚠️ PVP sin cargar en Contífico (afecta el valor mostrado, no la cantidad):<br>'
                + "<br>".join(f"· {a}" for a in alertas) + '</div>',
                unsafe_allow_html=True,
            )
    else:
        st.markdown('<div class="bcc-note-box">sin dato de inventario aún</div>', unsafe_allow_html=True)

    st.write("")
    panel_header("Detalle: campo ventas_consolidadas_ec", "distinto de la 'Facturación Contífico' de arriba — ver nota")
    vec = last_day.get("ventas_consolidadas_ec") or {}
    vec_total = vec.get("total_usd")
    st.markdown(
        '<div class="bcc-dev-box">\U0001F6A7 <b>En desarrollo, no confundir con el KPI de arriba.</b> '
        'La tarjeta "Facturación Contífico" usa <code>contifico_documentos</code>/<code>contifico_total_usd</code> — '
        'esos SÍ son datos reales documento por documento y son la fuente primaria de este tab. '
        f'<code>ventas_consolidadas_ec</code> es un campo distinto (desglose tienda/web/en_línea/evento) que '
        f'todavía viene en {fmt_or_missing(vec_total, "${:,.0f}")} porque su lógica de agregación aún no está '
        'enlazada a esos documentos — pendiente de confirmar con Luis.<br><br>'
        f'<code>{vec}</code></div>', unsafe_allow_html=True,
    )

# ============================================================
# COLOMBIA
# ============================================================
with tab_col:
    start, end = date_range_selector("col")
    rows = rows_in_range(start, end)
    parcial_banner(rows, end)

    col_total = sum_field(rows, "prestashop_total_col_cop")
    col_pedidos = sum_field(rows, "prestashop_pedidos_col")
    ticket = (col_total / col_pedidos) if (col_total and col_pedidos) else None
    col_usd_equiv = (col_total / trm) if col_total is not None else None
    leads_col = sum_field(rows, "kommo_leads_col")

    # Jerarquía de fuente COL (actualizado 13/ago): la Matriz de Ventas Colombia (Google
    # Sheets) es la fuente PRIMARIA declarada -- tiene columna de monto, Delivery cruzado con
    # PrestaShop e "Id orden Prestashop" para cruce exacto. La extracción vía export CSV YA
    # funciona (las celdas combinadas dejaron de ser un problema -- se lee la hoja completa,
    # 581 filas históricas desde may/2025). La única limitante real ahora: el equipo de
    # Colombia carga el Sheet de forma manual, así que un día sin filas registradas ahí
    # todavía se ve como 0 (no es un bug de lectura). PrestaShop COL sigue como complemento
    # funcional; Kommo aporta contexto de leads.
    vcol_preview = last_day.get("ventas_consolidadas_col") or {}
    vcol_total = vcol_preview.get("total_cop")

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card("Matriz Sheets (fuente primaria)", fmt_or_missing(vcol_total, "${:,.0f}"), "carga manual del equipo Colombia -- puede ir 1-2 días atrás del corte", MINT, "U0001F4CB")
    with k2:
        kpi_card("PrestaShop COL (complementario, activo)", fmt_or_missing(col_total, "${:,.0f}"), f"{fmt_or_missing(col_pedidos)} pedidos pagados", MINT, "\U0001F4B5")
    with k3:
        kpi_card("Ticket promedio", fmt_or_missing(ticket, "${:,.0f}"), "COP · base PrestaShop", BLUE, "\U0001F4C8")
    with k4:
        kpi_card("Leads Kommo COL", fmt_or_missing(leads_col, "{:,.0f}"), "contexto de leads, no de ventas", LILAC, "\U0001F465")

    st.markdown(
        '<div class="bcc-note-box">La Matriz de Ventas Colombia (Google Sheets) es la fuente primaria declarada '
        'para Colombia — tiene columna de monto, Delivery cruzado contra PrestaShop e "Id orden Prestashop" para '
        'cruce exacto. Hoy no se puede usar en el dashboard porque el equipo de Colombia combina celdas en el '
        'Sheet, lo que rompe la lectura automática (API/exportación solo lee la celda superior-izquierda de un '
        'rango combinado). Mientras se resuelve, PrestaShop COL es el número que se usa en el día a día. '
        f'Equivalente USD de PrestaShop: {fmt_or_missing(col_usd_equiv, "${:,.0f}")} · TRM ref. {trm} (manual).</div>',
        unsafe_allow_html=True,
    )
    st.write("")
    col1, col2 = st.columns([2, 1])
    with col1:
        panel_header("Ingresos web diarios", "COP, PrestaShop COL (complementario)")
        dfrows = pd.DataFrame(rows).sort_values("date")
        fig = go.Figure(go.Scatter(x=dfrows["date"], y=dfrows["prestashop_total_col_cop"], fill="tozeroy",
                                    line=dict(color=MINT, width=2), fillcolor="rgba(79,227,193,0.28)"))
        fig = mark_parcial_day(fig, rows, end)
        st.plotly_chart(plotly_dark_layout(fig, height=260), use_container_width=True)
    with col2:
        panel_header("Mensajería Kommo COL")
        entrantes = sum_field(rows, "kommo_msj_entrantes_col")
        salientes = sum_field(rows, "kommo_msj_salientes_col")
        lapso = last_day.get("kommo_lapso_medio_col_s")
        st.markdown(f'<div class="bcc-card" style="--accent:{BLUE};"><span class="lbl">Entrantes / Salientes</span>'
                     f'<div class="val" style="font-size:20px;">{fmt_or_missing(entrantes)} / {fmt_or_missing(salientes)}</div>'
                     f'<div class="sub">Lapso resp. {(str(int(lapso//60))+"m") if lapso else "sin dato"}</div></div>', unsafe_allow_html=True)

    st.write("")
    panel_header("Funnel Kommo COL", "ventana móvil 30 días · orden real del pipeline")
    fcol = last_day.get("kommo_funnel_col") or {}
    by_status = fcol.get("byStatus") or {}
    if by_status:
        total_funnel_col = sum(by_status.values())
        items_col = sorted(by_status.items(), key=lambda x: KOMMO_STAGE_ORDER_COL.index(x[0]) if x[0] in KOMMO_STAGE_ORDER_COL else 999)
        sdf = pd.DataFrame(items_col, columns=["etapa", "leads"])
        sdf["pct"] = sdf["leads"].apply(lambda v: f"{(v/total_funnel_col*100):.0f}%" if total_funnel_col else "")
        fig = go.Figure(go.Bar(x=sdf["etapa"], y=sdf["leads"], text=sdf["pct"], textposition="outside", marker_color=LILAC))
        st.plotly_chart(plotly_dark_layout(fig, height=280), use_container_width=True)
    else:
        st.markdown('<div class="bcc-note-box">sin dato de funnel para Colombia en este corte</div>', unsafe_allow_html=True)

    st.write("")
    panel_header("Detalle: Matriz Sheets Colombia (fuente primaria)", "metodología y estado real de ventas_consolidadas_col")
    st.markdown(
        '<div class="bcc-note-box">✅ <b>Extracción resuelta el 13/ago — ya no está bloqueada.</b> '
        'El Sheet "Resumen de ventas Bou Entertainment Colombia" tiene columna de monto (PVP), columna Delivery '
        'que cruza contra PrestaShop, y columna "Id orden Prestashop" para cruce exacto. Se registra por producto '
        '(una fila por producto, no por pedido). Columna "Tipo de Evento": Web = pedido web, EN LINEA = pedido '
        'por WhatsApp.<br><br>'
        'Las celdas combinadas ya no son un problema: se lee el Sheet como export CSV completo (no la API de '
        'lectura celda por celda), lo que trae las 581 filas reales sin perder datos. Se hizo un backfill '
        'histórico completo (may/2025 – ago/2026) sobre daily-log.json.<br><br>'
        '<b>Limitante real, no de metodología:</b> el equipo de Colombia carga el Sheet manualmente — si un día '
        'todavía no tiene filas ahí, este campo legítimamente muestra 0 (no significa que el dashboard esté roto). '
        'La última fecha con filas en el Sheet fuente puede ir 1-2 días atrás de hoy.<br><br>'
        f'<code>{vcol_preview}</code></div>', unsafe_allow_html=True,
    )# ============================================================
# MARKETING
# ============================================================
with tab_mkt:
    start, end = date_range_selector("mkt")
    rows = rows_in_range(start, end)
    parcial_banner(rows, end)

    spend = sum_field(rows, "meta_spend_usd")
    ingresos = sum_field(rows, "meta_ingresos_usd")
    roas = (ingresos / spend) if (ingresos is not None and spend) else None
    clics = sum_field(rows, "meta_clics")
    impresiones = sum_field(rows, "meta_impresiones")
    ctr = (clics / impresiones * 100) if (clics is not None and impresiones) else None
    compras_meta = sum_field(rows, "meta_compras")
    cpa = (spend / compras_meta) if (spend is not None and compras_meta) else None
    sesiones_ec = sum_field(rows, "ga4_sesiones_ec")
    sesiones_col = sum_field(rows, "ga4_sesiones_col")
    sesiones_total = (sesiones_ec or 0) + (sesiones_col or 0) if (sesiones_ec is not None or sesiones_col is not None) else None

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card("Spend Meta Ads", fmt_or_missing(spend, "${:,.0f}"), "periodo seleccionado, sin split por país", CORAL, "\U0001F4B8")
    with k2:
        kpi_card("ROAS", (f"{roas:.1f}x" if roas is not None else "sin dato"), "ingresos/spend, recalculado sobre la suma", MINT, "\U0001F4C8")
    with k3:
        kpi_card("CPA", (f"${cpa:,.2f}" if cpa is not None else "sin dato"),
                  f"spend/compras · {compras_meta or 0} compras en el periodo" if compras_meta else "sin compras registradas en el periodo",
                  LILAC, "\U0001F3AF")
    with k4:
        kpi_card("CTR", (f"{ctr:.2f}%" if ctr is not None else "sin dato"), "clics/impresiones del periodo", BLUE, "\U0001F5B1️")

    st.write("")
    panel_header("Sesiones por país (GA4)", f"por hostname, sin doble conteo · {sesiones_total or 0} sesiones totales (EC {sesiones_ec or 0} · COL {sesiones_col or 0})")
    dfrows = pd.DataFrame(rows).sort_values("date")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=dfrows["date"], y=dfrows.get("ga4_sesiones_ec"), name="Ecuador", marker_color=GOLD))
    fig.add_trace(go.Bar(x=dfrows["date"], y=dfrows.get("ga4_sesiones_col"), name="Colombia", marker_color=MINT))
    fig.update_layout(barmode="stack")
    st.plotly_chart(plotly_dark_layout(fig, height=260), use_container_width=True)

    st.write("")
    panel_header("Canales de adquisición (GA4)", "sesiones por canal, EC + COL agregado del periodo")
    canal_totales = {}
    for r in rows:
        ga4c = r.get("ga4_canales") or {}
        for pais_canales in (ga4c.values() if isinstance(ga4c, dict) else []):
            if isinstance(pais_canales, dict):
                for canal, val in pais_canales.items():
                    sesiones_val = val.get("sessions", 0) if isinstance(val, dict) else (val or 0)
                    canal_totales[canal] = canal_totales.get(canal, 0) + (sesiones_val or 0)
    if canal_totales:
        cdf = pd.DataFrame(sorted(canal_totales.items(), key=lambda x: -x[1]), columns=["canal", "sesiones"])
        fig_canal = go.Figure(go.Bar(x=cdf["sesiones"], y=cdf["canal"], orientation="h", marker_color=CYAN))
        st.plotly_chart(plotly_dark_layout(fig_canal, height=260), use_container_width=True)
    else:
        st.markdown('<div class="bcc-note-box">sin dato de canales GA4 en el periodo seleccionado</div>', unsafe_allow_html=True)

    st.write("")
    panel_header("Embudo GA4 (eventos clave)", "view_item → add_to_cart → begin_checkout → purchase, agregado del periodo")
    view_items = add_cart = add_checkout = purchases = 0
    has_events = False
    for r in rows:
        ev = r.get("ga4_eventos_clave") or {}
        for pais_events in ev.values() if isinstance(ev, dict) else []:
            if isinstance(pais_events, dict):
                has_events = True
                view_items += pais_events.get("view_item", 0) or 0
                add_cart += pais_events.get("add_to_cart", 0) or 0
                add_checkout += pais_events.get("begin_checkout", 0) or 0
                purchases += pais_events.get("purchase", 0) or 0
    if has_events:
        fig = go.Figure(go.Funnel(y=["View item", "Add to cart", "Begin checkout", "Purchase"], x=[view_items, add_cart, add_checkout, purchases],
                                   marker=dict(color=[CYAN, GOLD, BLUE, MINT])))
        st.plotly_chart(plotly_dark_layout(fig, height=280), use_container_width=True)
    else:
        st.markdown('<div class="bcc-note-box">sin dato de eventos GA4 en el periodo seleccionado</div>', unsafe_allow_html=True)

    st.write("")
    panel_header("Ad sets — frecuencia (fatiga)", "frequency > 5 = posible fatiga de audiencia (linea roja) · último valor observado en el periodo")
    adset_freq = {}
    for r in rows:
        for adset in (r.get("ads_ad_sets") or []):
            freq = adset.get("frequency")
            name = adset.get("adset")
            if name and freq:
                adset_freq[name] = freq
    if adset_freq:
        fdf = pd.DataFrame(sorted(adset_freq.items(), key=lambda x: -x[1]), columns=["ad set", "frequency"])
        bar_colors = [CORAL if f > 5 else BLUE for f in fdf["frequency"]]
        fig_freq = go.Figure(go.Bar(x=fdf["frequency"], y=fdf["ad set"], orientation="h", marker_color=bar_colors))
        fig_freq.add_vline(x=5, line_width=1, line_dash="dash", line_color=CORAL)
        st.plotly_chart(plotly_dark_layout(fig_freq, height=max(200, 30 * len(fdf))), use_container_width=True)
        n_fatigued = int((fdf["frequency"] > 5).sum())
        if n_fatigued:
            st.caption(f"⚠️ {n_fatigued} ad set(s) por encima de frequency 5 en el periodo.")
    else:
        st.markdown('<div class="bcc-note-box">sin dato de ad sets en el periodo seleccionado</div>', unsafe_allow_html=True)

    st.write("")
    st.markdown(
        '<div class="bcc-note-box">Nota: <code>meta_alcance</code> es la suma por campaña, no deduplicado — '
        'puede sobreestimar si hay audiencias superpuestas. Meta Ads no tiene split confiable por país, se muestra '
        'agregado. <code>nasa-histoires.myshopify.com</code> se excluye de GA4 EC/COL (canal Shopify separado).</div>',
        unsafe_allow_html=True,
    )

# ============================================================
# CONCILIACION (estatico, sin selector de fecha)
# ============================================================
with tab_conc:
    st.markdown(
        f'<div style="font-size:12px;color:{MUTED};margin-bottom:14px;">\U0001F550 Estado verificado 12/ago/2026 '
        '— este panel NO sigue el selector de fecha, muestra un estado actual, no una serie temporal.</div>',
        unsafe_allow_html=True,
    )
    # Contífico dinámico (no hardcodeado): la firma electrónica del SRI que lo bloqueó desde
    # 10/jul/2026 ya fue reactivada (confirmado por Luis, 12/ago) — este badge se calcula del
    # último dato real en vez de repetir un estado que ya cambió.
    ctf_caida_hoy = last_day.get("contifico_fuente_caida")
    b1, b2 = st.columns(2)
    with b1:
        badge("down" if ctf_caida_hoy else "ok", "Contífico",
              "0 documentos hoy — revisar bloqueo" if ctf_caida_hoy
              else f"activo, firma electrónica SRI reactivada — {fmt_or_missing(last_day.get('contifico_documentos'))} documentos el {last_day['date']}")
        badge("ok", "PrestaShop", "bug de current_state corregido 30/jul/2026 — histórico jun-jul no reauditado retroactivamente")
        badge("ok", "Kommo leads EC/COL", "fuente oficial: API con filtro de fecha, resuelto 30/jul/2026")
    with b2:
        badge("warn", "Sheet Colombia", "lecturas del conector de Drive pueden estar desactualizadas — no conciliado en montos, solo en conteo de líneas (19 vs 20, muestra puntual jul)")
        badge("ok", "Mensajería Kommo (EC + COL)", "endpoint /api/v4/events correcto, agregación diaria ya implementada")
        badge("warn", "ventas_consolidadas_ec / _col", "en el schema desde ago/2026, SIN metodología documentada aún — confirmar con Luis antes de usar")

    st.write("")
    panel_header("Pendientes que afectan la veracidad del dato", "tomados tal cual de la wiki de Notion, no se inventan nuevos")
    st.markdown(
        '<div class="bcc-card" style="--accent:' + CORAL + ';">'
        '<ul style="margin:0;padding-left:18px;font-size:13px;color:' + TEXT + ';line-height:1.9;">'
        '<li>Implementar cruce por número de pedido entre PrestaShop EC y Contífico (referencia guardada en el documento) — hoy solo se compara en agregado (pedidos vs. documentos), no pedido por pedido.</li>'
        '<li>Recalcular histórico de pedidos PrestaShop EC (1 jun–30 jul) con el filtro de <code>current_state</code> ya corregido.</li>'
        '<li>Cerrar conciliación de montos entre la Matriz de Ventas Colombia (Sheet) y PrestaShop COL (falta columna de monto estandarizada).</li>'
        '<li>Confirmar con Alexis (CEO) objetivos de venta y reglas de alerta antes de agregar alertas automáticas de meta.</li>'
        '<li>Confirmar con Luis el estado real de <code>ventas_consolidadas_ec</code> / <code>ventas_consolidadas_col</code> — existen en el schema pero no tienen metodología documentada.</li>'
        '<li>Habilitar en Contífico el permiso de los endpoints <code>/cliente/</code> y <code>/cobro/</code> si se quiere segmentar por cliente o forma de pago.</li>'
        '</ul></div>', unsafe_allow_html=True,
    )

st.write("")
st.markdown(f'<div style="text-align:center;font-size:11px;color:{MUTED};margin-top:24px;">'
            f'Datos reales · BOU Entertainment · última fila disponible: {today_str} '
            f'{"(parcial)" if last_day.get("parcial") else "(cierre)"}</div>', unsafe_allow_html=True)
