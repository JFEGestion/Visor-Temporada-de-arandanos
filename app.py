"""Dashboard Temporada Arándanos: Recepción, Control de estimados y Proceso.
Archivos esperados en la misma carpeta: Recepciones.xlsx, Procesos.xlsx y
CURVA TENTATIVA TEMP ARANDANOS 26-27 V4.xlsm"""
import io, os
import unicodedata
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Temporada Arándanos", page_icon="🫐", layout="wide")

# ───────────────────────── CONFIGURACIÓN ─────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
CAPACIDAD_LINEAS = {"SP1-L1": 36000, "SP2-L2": 18000, "SP2-L3": 14500, "SP3-L4": 14500}  # KN/día
TURNOS_POR_DIA = 2                       # capacidad por turno = capacidad día / 2
EXPORTADORAS_SIN_PROCESO = ["S&A"]       # formato bulk: no pasan por proceso
TOLERANCIA_RECLAS_KG = 1.0
TZ = ZoneInfo("America/Santiago")
# fila de la curva (en minúsculas, sin espacios extremos) -> (exportadora, condición)
MAPEO_CURVA = {
    "bluefield conv": ("Bluefield", "Convencional"),
    "bluefield orgánico": ("Bluefield", "Orgánico"),
    "rodeo orgánico": ("El Rodeo", "Orgánico"),
    "rodeo bulk orgánico": ("El Rodeo", "Orgánico"),
    "s&a bulk conv.": ("S&A", "Convencional"),
    "s&a bulk org.": ("S&A", "Orgánico"),
}
C_ACC, C_REJ = "#2E7D32", "#C62828"
COLORES_ESTADO = {"Aceptada": C_ACC, "Rechazada": C_REJ}


# ───────────────────────── UTILIDADES ─────────────────────────
def norm(txt):
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return t.strip().lower().replace(" ", "_").replace("/", "_")

def f0(x):
    return f"{x:,.0f}".replace(",", ".") if pd.notna(x) else "-"

def pct(x):
    return f"{x:.1%}".replace(".", ",") if pd.notna(x) else "-"

def mostrar(fig):
    try:
        st.plotly_chart(fig, width="stretch")
    except TypeError:
        st.plotly_chart(fig, use_container_width=True)

def tabla(df, key=None):
    try:
        st.dataframe(df, width="stretch", hide_index=True)
    except TypeError:
        st.dataframe(df, use_container_width=True, hide_index=True)

def boton_excel(df, nombre, key):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Datos")
    st.download_button("⬇️ Descargar Excel", buf.getvalue(), file_name=f"{nombre}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key)

def buscar_archivo(prefijo):
    for f in sorted(os.listdir(BASE)):
        if f.lower().startswith(prefijo.lower()) and f.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            return os.path.join(BASE, f)
    return None

def leer(path):
    return pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)

def corta_exportadora(x):
    t = str(x).lower()
    if "bluefield" in t: return "Bluefield"
    if "rodeo" in t: return "El Rodeo"
    if "s&a" in t: return "S&A"
    return str(x).strip()

def limpiar_texto(s):
    return s.fillna("").astype(str).str.strip()

def normalizar_variedad(s):
    s = limpiar_texto(s)
    clave = s.str.upper().str.replace(" ", "", regex=False)
    modo = s.groupby(clave).agg(lambda v: v.value_counts().index[0]).str.title()
    return clave.map(modo).fillna(s)

def lunes(fechas):
    return (fechas - pd.to_timedelta(fechas.dt.weekday, unit="D")).dt.normalize()

def lunes_iso(anio, semana):
    try:
        return pd.Timestamp(date.fromisocalendar(int(anio), int(semana), 1))
    except ValueError:  # semana 53 en año que no la tiene
        return pd.Timestamp(date.fromisocalendar(int(anio), 52, 1) + timedelta(days=7))

def lbl_semana(m):
    return f"S{m.isocalendar().week:02d} · {m:%d-%m}"


# ───────────────────────── CARGA DE DATOS ─────────────────────────
def preparar_recepcion(df):
    df = df.copy()
    df.columns = [norm(c) for c in df.columns]
    df["fecha_guia"] = pd.to_datetime(df["fecha_guia"], errors="coerce")
    df = df.dropna(subset=["fecha_guia"])
    df["kn"] = pd.to_numeric(df["kn"], errors="coerce").fillna(0)
    for c in ["exportadora", "productor", "huerto", "condicion", "tipo_de_producto", "formato_recip", "rechazo"]:
        df[c] = limpiar_texto(df[c])
    df["exportadora"] = df["exportadora"].map(corta_exportadora)
    df["variedad"] = normalizar_variedad(df["variedad"])
    df["huerto"] = df["huerto"].replace("", "(sin huerto)")
    df["rech"] = df["rechazo"].str.lower().str.contains("rechaz")      # vacío = a proceso
    df["condicion_est"] = np.where(df["condicion"].str.lower().str.startswith("conv"), "Convencional", "Orgánico")
    df["bulk"] = df["exportadora"].isin(EXPORTADORAS_SIN_PROCESO) | df["formato_recip"].str.lower().str.contains("bulk")
    df["fecha"] = df["fecha_guia"].dt.normalize()
    df["monday"] = lunes(df["fecha_guia"])
    df["camion"] = (df["exportadora"] + "|" + df["productor"] + "|" + df["guia"].astype(str) + "|"
                    + df["fecha"].dt.strftime("%Y-%m-%d"))
    return df.reset_index(drop=True)

def preparar_proceso(df):
    df = df.copy()
    df.columns = [norm(c) for c in df.columns]
    df["fecha_proceso"] = pd.to_datetime(df["fecha_proceso"], errors="coerce")
    df = df.dropna(subset=["fecha_proceso"])
    df["kn"] = pd.to_numeric(df["kn"], errors="coerce").fillna(0)
    for c in ["exportadora", "sala", "turno", "tipo_proceso", "tipo_producto", "condicion"]:
        df[c] = limpiar_texto(df[c])
    df["exportadora"] = df["exportadora"].map(corta_exportadora)
    df["sala"] = df["sala"].str.upper().str.replace(" ", "", regex=False)
    df["variedad"] = normalizar_variedad(df["variedad"])
    tp = df["tipo_producto"].str.lower()
    df["producto"] = np.select([tp.str.contains("merma"), tp.str.contains("jugo"), tp.str.startswith("iqf")],
                               ["Merma", "IQF Jugo", "IQF"], default="Exportación")
    df["fecha"] = df["fecha_proceso"].dt.normalize()
    df["trabajadores"] = pd.to_numeric(df["numero_de_trabajadores"], errors="coerce")
    return df.reset_index(drop=True)

def reclasificar(rec, proc):
    """Si un lote procesó más KN que sus KN aceptados, la diferencia es rechazo reclasificado."""
    rec = rec.copy()
    rec["reclasificado"] = False
    if rec.empty or proc.empty:
        return rec
    procesado = proc.groupby("lote").kn.sum()
    aceptado = rec[~rec.rech].groupby("lote").kn.sum().reindex(procesado.index).fillna(0)
    exceso = (procesado - aceptado).clip(lower=0)
    rj = rec[rec.rech].sort_values(["lote", "pallet"]).copy()
    rj["cum"] = rj.groupby("lote").kn.cumsum()
    rj["exc"] = rj["lote"].map(exceso).fillna(0)
    idx = rj.index[(rj.exc > 0) & (rj.cum <= rj.exc + TOLERANCIA_RECLAS_KG)]
    rec.loc[idx, "rech"] = False
    rec.loc[idx, "reclasificado"] = True
    return rec

@st.cache_data(show_spinner="Cargando datos…")
def cargar_todo(p_rec, m_rec, p_proc, m_proc):
    rec = preparar_recepcion(leer(p_rec))
    proc = preparar_proceso(leer(p_proc))
    rec = reclasificar(rec, proc)
    rec["estado"] = np.where(rec["rech"], "Rechazada", "Aceptada")
    return rec, proc

@st.cache_data(show_spinner="Leyendo curva de estimados…")
def cargar_estimados(path, mtime):
    ws = load_workbook(path, data_only=True).active
    filas = list(ws.iter_rows(values_only=True))
    hdr = next(i for i, r in enumerate(filas) if r[1] and str(r[1]).strip().lower() == "exportadora")
    cols = []
    for j in range(2, len(filas[hdr])):
        try:
            cols.append((j, int(float(filas[hdr][j]))))
        except (TypeError, ValueError):
            break  # termina la zona de semanas
    out = []
    for r in filas:
        clave = str(r[1]).strip().lower() if r[1] else ""
        if clave in MAPEO_CURVA:
            for j, sem in cols:
                kg = r[j] if isinstance(r[j], (int, float)) else 0
                out.append((*MAPEO_CURVA[clave], sem, float(kg)))
    return pd.DataFrame(out, columns=["exportadora", "condicion_est", "semana", "kg"])


# ───────────────────────── FILTROS ─────────────────────────
def msel(label, serie, key):
    ops = sorted(serie.dropna().unique().tolist())
    sel = st.sidebar.multiselect(label, ops, default=ops, key=key)
    return serie.isin(sel)

def rango_fechas(label, serie, key):
    if serie.empty:
        return pd.Series(True, index=serie.index)
    mn, mx = serie.min().date(), serie.max().date()
    r = st.sidebar.date_input(label, value=(mn, mx), min_value=mn, max_value=mx, key=key)
    if isinstance(r, (tuple, list)):
        ini, fin = (r[0], r[-1]) if len(r) else (mn, mx)
    else:
        ini = fin = r
    return (serie.dt.date >= ini) & (serie.dt.date <= fin)


# ───────────────────────── PÁGINA: RECEPCIÓN ─────────────────────────
def pagina_recepcion(rec):
    st.header("📥 Recepción de fruta")
    if rec.empty:
        st.info("Aún no hay recepciones cargadas.")
        return
    st.sidebar.subheader("Filtros de recepción")
    m = (msel("Exportadora", rec.exportadora, "r_exp") & rango_fechas("Fecha guía", rec.fecha, "r_fec")
         & msel("Variedad", rec.variedad, "r_var") & msel("Huerto", rec.huerto, "r_hue")
         & msel("Condición", rec.condicion, "r_con") & msel("Tipo de producto", rec.tipo_de_producto, "r_tip"))
    est = st.sidebar.radio("Estado de la fruta", ["Todas", "Aceptada", "Rechazada"], key="r_est")
    if est != "Todas":
        m &= rec.estado == est
    d = rec[m]
    if d.empty:
        st.warning("Sin datos con los filtros seleccionados.")
        return

    tot, rej = d.kn.sum(), d[d.rech].kn.sum()
    c = st.columns(6)
    c[0].metric("KN recibidos", f0(tot))
    c[1].metric("KN aceptados", f0(tot - rej))
    c[2].metric("KN rechazados", f0(rej))
    c[3].metric("% rechazo", pct(rej / tot if tot else np.nan))
    c[4].metric("Camiones", f0(d.camion.nunique()))
    c[5].metric("Pallets", f0(len(d)))

    dia = d.pivot_table(index="fecha", columns="estado", values="kn", aggfunc="sum", fill_value=0).reset_index()
    for e in ["Aceptada", "Rechazada"]:
        if e not in dia: dia[e] = 0.0
    dia["Total"] = dia.Aceptada + dia.Rechazada
    dia["Acumulado total"] = dia.Total.cumsum()
    dia["Acumulado rechazado"] = dia.Rechazada.cumsum()
    g1, g2 = st.columns(2)
    fig = px.line(dia, x="fecha", y=["Acumulado total", "Acumulado rechazado"], markers=True,
                  title="KN recibidos acumulados", color_discrete_sequence=["#1565C0", C_REJ])
    fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="")
    with g1: mostrar(fig)
    fig = px.bar(dia, x="fecha", y=["Aceptada", "Rechazada"], title="KN recibidos por día",
                 color_discrete_map=COLORES_ESTADO)
    fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="")
    with g2: mostrar(fig)

    cam = d.groupby("fecha").camion.nunique().reset_index(name="Camiones")
    g3, g4 = st.columns(2)
    fig = px.bar(cam, x="fecha", y="Camiones", title="Camiones recibidos por día", text="Camiones")
    fig.update_layout(xaxis_title="")
    with g3: mostrar(fig)
    por = d.groupby(["exportadora", "estado"]).kn.sum().reset_index()
    fig = px.bar(por, x="exportadora", y="kn", color="estado", title="KN por exportadora",
                 color_discrete_map=COLORES_ESTADO)
    fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="")
    with g4: mostrar(fig)

    g5, g6, g7 = st.columns(3)
    for col, campo, tit in [(g5, "tipo_de_producto", "Tipo de producto"), (g6, "condicion", "Condición"),
                            (g7, "variedad", "Variedad")]:
        x = d.groupby([campo, "estado"]).kn.sum().reset_index()
        if campo == "variedad":
            top = d.groupby(campo).kn.sum().nlargest(10).index
            x = x[x[campo].isin(top)]
        fig = px.bar(x, x=campo, y="kn", color="estado", title=f"KN por {tit.lower()}",
                     color_discrete_map=COLORES_ESTADO)
        fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="", showlegend=False)
        with col: mostrar(fig)

    h = d.groupby(["exportadora", "huerto"]).apply(lambda x: pd.Series({
        "KN recibidos": x.kn.sum(), "KN aceptados": x[~x.rech].kn.sum(), "KN rechazados": x[x.rech].kn.sum(),
        "Camiones": x.camion.nunique(), "Pallets": len(x)}), include_groups=False).reset_index()
    h["% rechazo"] = (h["KN rechazados"] / h["KN recibidos"]).round(3)
    h = h.sort_values("KN recibidos", ascending=False).round(0).assign(**{"% rechazo": lambda t: (h["% rechazo"] * 100).round(1)})
    st.subheader("Detalle por huerto")
    tabla(h)
    boton_excel(h, "recepcion_por_huerto", "x_r_h")
    with st.expander("Ver detalle por pallet"):
        det = d[["fecha_guia", "exportadora", "productor", "huerto", "guia", "lote", "pallet", "variedad", "condicion",
                 "tipo_de_producto", "estado", "reclasificado", "kn"]].sort_values("fecha_guia")
        tabla(det)
        boton_excel(det, "recepcion_detalle", "x_r_d")


# ───────────────────────── PÁGINA: CONTROL DE ESTIMADOS ─────────────────────────
def tabla_control(est, rec, corte):
    e = est.groupby(["exportadora", "condicion_est"]).apply(lambda x: pd.Series({
        "Estimado temporada": x.kg.sum(), "Estimado a la fecha": x[x.monday <= corte].kg.sum()}),
        include_groups=False).reset_index()
    r = rec.groupby(["exportadora", "condicion_est"]).apply(lambda x: pd.Series({
        "Recibido total": x.kn.sum(), "Recibido a la fecha": x[x.monday <= corte].kn.sum(),
        "Rechazado": x[x.rech].kn.sum()}), include_groups=False).reset_index()
    t = e.merge(r, on=["exportadora", "condicion_est"], how="outer").fillna(0)
    t["Desviación (KN)"] = t["Recibido a la fecha"] - t["Estimado a la fecha"]
    t["Desviación %"] = np.where(t["Estimado a la fecha"] > 0, t["Desviación (KN)"] / t["Estimado a la fecha"], np.nan)
    t["Por recibir (saldo)"] = t["Estimado temporada"] - t["Recibido total"]
    t["Avance %"] = np.where(t["Estimado temporada"] > 0, t["Recibido total"] / t["Estimado temporada"], np.nan)
    t["% rechazo"] = np.where(t["Recibido total"] > 0, t["Rechazado"] / t["Recibido total"], np.nan)
    return t.rename(columns={"exportadora": "Exportadora", "condicion_est": "Condición"})

def pagina_control(rec, est_raw, anio_ini):
    st.header("🎯 Control de recepción vs estimado (KN)")
    est = est_raw.copy()
    est["anio"] = np.where(est.semana >= 40, anio_ini, anio_ini + 1)
    est["monday"] = [lunes_iso(a, s) for a, s in zip(est.anio, est.semana)]

    st.sidebar.subheader("Filtros de control")
    exps = sorted(set(est.exportadora) | set(rec.exportadora))
    sel_e = st.sidebar.multiselect("Exportadora", exps, default=exps, key="c_exp")
    sel_c = st.sidebar.multiselect("Condición", ["Convencional", "Orgánico"], default=["Convencional", "Orgánico"], key="c_con")
    est = est[est.exportadora.isin(sel_e) & est.condicion_est.isin(sel_c)]
    rec = rec[rec.exportadora.isin(sel_e) & rec.condicion_est.isin(sel_c)]

    semanas = sorted(set(est.monday) | set(rec.monday))
    hoy = pd.Timestamp(lunes(pd.Series([pd.Timestamp(datetime.now(TZ).date())])).iloc[0])
    if hoy in semanas: ini = semanas.index(hoy)
    elif len(rec): ini = semanas.index(rec.monday.max())
    else: ini = 0
    corte = st.sidebar.selectbox("Semana de corte", semanas, index=ini, format_func=lbl_semana, key="c_cor")
    st.caption(f"Temporada {anio_ini}-{str(anio_ini + 1)[-2:]} · desviación calculada hasta la semana {lbl_semana(corte)} · "
               "el rechazo cuenta como recibido.")

    t = tabla_control(est, rec, corte)
    tot = t.sum(numeric_only=True)
    c = st.columns(6)
    c[0].metric("Estimado temporada", f0(tot["Estimado temporada"]))
    c[1].metric("Recibido total", f0(tot["Recibido total"]))
    c[2].metric("Por recibir (saldo)", f0(tot["Por recibir (saldo)"]))
    c[3].metric("Avance temporada", pct(tot["Recibido total"] / tot["Estimado temporada"] if tot["Estimado temporada"] else np.nan))
    dv = tot["Desviación (KN)"]
    c[4].metric("Desviación a la fecha", f0(dv), f"{pct(dv / tot['Estimado a la fecha']) if tot['Estimado a la fecha'] else '-'}")
    c[5].metric("Rechazado", f0(tot["Rechazado"]), pct(tot["Rechazado"] / tot["Recibido total"] if tot["Recibido total"] else np.nan),
                delta_color="off")

    w = pd.DataFrame({"monday": semanas})
    w["Estimado"] = w.monday.map(est.groupby("monday").kg.sum()).fillna(0)
    w["Aceptada"] = w.monday.map(rec[~rec.rech].groupby("monday").kn.sum()).fillna(0)
    w["Rechazada"] = w.monday.map(rec[rec.rech].groupby("monday").kn.sum()).fillna(0)
    w["Semana"] = w.monday.map(lbl_semana)
    w["Estimado acumulado"] = w.Estimado.cumsum()
    w["Recibido acumulado"] = (w.Aceptada + w.Rechazada).cumsum().where(w.monday <= corte)

    g1, g2 = st.columns(2)
    fig = go.Figure()
    fig.add_bar(x=w.Semana, y=w.Aceptada, name="Recibido aceptado", marker_color=C_ACC)
    fig.add_bar(x=w.Semana, y=w.Rechazada, name="Recibido rechazado", marker_color=C_REJ)
    fig.add_scatter(x=w.Semana, y=w.Estimado, name="Estimado", mode="lines+markers", line=dict(color="#1565C0", width=3))
    fig.update_layout(barmode="stack", title="KN por semana: estimado vs recibido", yaxis_title="KN", legend_title="")
    with g1: mostrar(fig)
    fig = go.Figure()
    fig.add_scatter(x=w.Semana, y=w["Estimado acumulado"], name="Estimado acumulado", mode="lines", line=dict(color="#1565C0", width=3))
    fig.add_scatter(x=w.Semana, y=w["Recibido acumulado"], name="Recibido acumulado", mode="lines+markers", line=dict(color=C_ACC, width=3))
    fig.update_layout(title="KN acumulados: estimado vs recibido", yaxis_title="KN", legend_title="")
    with g2: mostrar(fig)

    d = t.copy()
    d["Etiqueta"] = d.Exportadora + " · " + d["Condición"]
    fig = px.bar(d, x="Etiqueta", y="Desviación (KN)", title="Desviación a la fecha por exportadora y condición (KN)",
                 color=np.where(d["Desviación (KN)"] >= 0, "Sobre estimado", "Bajo estimado"),
                 color_discrete_map={"Sobre estimado": C_ACC, "Bajo estimado": C_REJ})
    fig.update_layout(legend_title="", xaxis_title="")
    mostrar(fig)

    st.subheader("Detalle por exportadora y condición")
    out = t.drop(columns=["Etiqueta"], errors="ignore").copy()
    for col in ["Desviación %", "Avance %", "% rechazo"]:
        out[col] = (out[col] * 100).round(1)
    out = out.round(0).assign(**{c: (t[c] * 100).round(1) for c in ["Desviación %", "Avance %", "% rechazo"]})
    tabla(out)
    boton_excel(out, "control_estimados", "x_c")
    sem = w[["Semana", "Estimado", "Aceptada", "Rechazada"]].round(0)
    sem["Recibido total"] = sem.Aceptada + sem.Rechazada
    sem["Desviación"] = sem["Recibido total"] - sem.Estimado
    with st.expander("Ver detalle semanal"):
        tabla(sem)
        boton_excel(sem, "control_semanal", "x_c_s")


# ───────────────────────── PÁGINA: PROCESO ─────────────────────────
def pagina_proceso(rec, proc):
    st.header("🏭 Proceso")
    if proc.empty:
        st.info("Aún no hay proceso cargado.")
        return
    st.sidebar.subheader("Filtros de proceso")
    base = (msel("Exportadora", proc.exportadora, "p_exp") & msel("Condición", proc.condicion, "p_con")
            & msel("Variedad", proc.variedad, "p_var"))
    m = (base & rango_fechas("Fecha proceso", proc.fecha_proceso, "p_fec") & msel("Línea (sala)", proc.sala, "p_sal")
         & msel("Turno", proc.turno, "p_tur") & msel("Tipo de proceso", proc.tipo_proceso, "p_tip"))
    d = proc[m]
    if d.empty:
        st.warning("Sin datos con los filtros seleccionados.")
        return

    tot = d.kn.sum()
    runs = d.drop_duplicates(["numero_proceso", "sala", "turno", "fecha"])
    c = st.columns(6)
    c[0].metric("KN procesados", f0(tot))
    for col, p in zip(c[1:5], ["Exportación", "IQF", "IQF Jugo", "Merma"]):
        col.metric(f"% {p}", pct(d[d.producto == p].kn.sum() / tot if tot else np.nan))
    c[5].metric("Trabajadores prom.", f"{runs.trabajadores.mean():.1f}".replace(".", ","))

    g1, g2 = st.columns(2)
    x = d.groupby(["sala", "producto"]).kn.sum().reset_index()
    fig = px.bar(x, x="sala", y="kn", color="producto", title="KN procesados por línea")
    fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="")
    with g1: mostrar(fig)
    x = d.groupby(["fecha", "sala"]).kn.sum().reset_index()
    fig = px.bar(x, x="fecha", y="kn", color="sala", title="KN procesados por día y línea")
    fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="")
    with g2: mostrar(fig)

    g3, g4 = st.columns(2)
    x = d.groupby(["turno", "sala"]).kn.sum().reset_index()
    fig = px.bar(x, x="turno", y="kn", color="sala", barmode="group", title="Comparativo por turno (KN)")
    fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="")
    with g3: mostrar(fig)
    x = d.groupby("producto").kn.sum().reset_index()
    fig = px.pie(x, names="producto", values="kn", hole=0.45, title="Rendimiento de proceso (KN por producto)")
    with g4: mostrar(fig)

    # Utilización por línea
    dl = d.groupby(["sala", "fecha"]).kn.sum().reset_index()
    u = dl.groupby("sala").agg(**{"Días operados": ("fecha", "nunique"), "KN total": ("kn", "sum"),
                                  "KN prom/día": ("kn", "mean")}).reset_index()
    u["Capacidad día"] = u.sala.map(CAPACIDAD_LINEAS)
    u["Capacidad turno"] = u["Capacidad día"] / TURNOS_POR_DIA
    u["% utilización"] = u["KN prom/día"] / u["Capacidad día"] * 100
    u["Trabajadores prom."] = u.sala.map(runs.groupby("sala").trabajadores.mean())
    u = u.rename(columns={"sala": "Línea"})
    st.subheader("Comparativo por línea y utilización de capacidad")
    tabla(u.round(1))
    boton_excel(u.round(1), "proceso_lineas", "x_p_l")
    if u["% utilización"].notna().any():
        fig = px.bar(u, x="Línea", y="% utilización", title="% utilización promedio diario vs capacidad", text_auto=".0f")
        fig.add_hline(y=100, line_dash="dash")
        mostrar(fig)
    tt = runs.groupby("turno").agg(**{"Trabajadores prom.": ("trabajadores", "mean")}).reset_index()
    tt["KN"] = tt.turno.map(d.groupby("turno").kn.sum())
    tabla(tt.round(1).rename(columns={"turno": "Turno"}))

    # Acumulado procesado vs recibido (no depende de fecha, línea, turno)
    st.subheader("Acumulado procesado vs recibido")
    st.caption("Recibido = aceptado (sin rechazo) y sin bulk. Procesado = KN totales incl. merma. "
               "Usa los filtros de exportadora, condición y variedad.")
    r = rec[~rec.bulk & rec.exportadora.isin(proc.exportadora[base].unique()) & rec.variedad.isin(proc.variedad[base].unique())
            & rec.condicion.isin(proc.condicion[base].unique())]
    pr = proc[base]
    dias = pd.DataFrame({"fecha": sorted(set(r.fecha) | set(pr.fecha))})
    dias["Recibido aceptado acum."] = dias.fecha.map(r[~r.rech].groupby("fecha").kn.sum()).fillna(0).cumsum()
    dias["Rechazado acum."] = dias.fecha.map(r[r.rech].groupby("fecha").kn.sum()).fillna(0).cumsum()
    dias["Procesado acum."] = dias.fecha.map(pr.groupby("fecha").kn.sum()).fillna(0).cumsum()
    dias["Pendiente de procesar"] = dias["Recibido aceptado acum."] - dias["Procesado acum."]
    fig = px.line(dias, x="fecha", y=["Recibido aceptado acum.", "Procesado acum.", "Rechazado acum."], markers=True,
                  color_discrete_sequence=["#1565C0", C_ACC, C_REJ])
    fig.update_layout(legend_title="", yaxis_title="KN", xaxis_title="")
    mostrar(fig)

    s = r.groupby("exportadora").apply(lambda x: pd.Series({
        "Recibido total": x.kn.sum(), "Rechazado": x[x.rech].kn.sum(), "Aceptado": x[~x.rech].kn.sum()}),
        include_groups=False).reset_index()
    s["Procesado"] = s.exportadora.map(pr.groupby("exportadora").kn.sum()).fillna(0)
    s["Pendiente por procesar"] = s.Aceptado - s.Procesado
    s["% procesado"] = (s.Procesado / s.Aceptado.replace(0, np.nan) * 100).round(1)
    s = s.rename(columns={"exportadora": "Exportadora"})
    st.subheader("Saldo pendiente de procesar por exportadora")
    tabla(s.round(0).assign(**{"% procesado": s["% procesado"]}))
    boton_excel(s.round(0), "saldo_proceso", "x_p_s")
    with st.expander("Ver acumulado diario"):
        tabla(dias.round(0))
        boton_excel(dias.round(0), "acumulado_proceso", "x_p_a")


# ───────────────────────── MAIN ─────────────────────────
def main():
    st.title("🫐 Temporada Arándanos – Recepción y Proceso")
    p_rec, p_proc, p_est = buscar_archivo("Recepciones"), buscar_archivo("Procesos"), buscar_archivo("CURVA")
    faltan = [n for n, p in [("Recepciones", p_rec), ("Procesos", p_proc), ("CURVA TENTATIVA…", p_est)] if not p]
    if faltan:
        st.error(f"No encuentro estos archivos en la carpeta: {', '.join(faltan)}")
        st.stop()
    try:
        rec, proc = cargar_todo(p_rec, os.path.getmtime(p_rec), p_proc, os.path.getmtime(p_proc))
        est = cargar_estimados(p_est, os.path.getmtime(p_est))
    except Exception as e:
        st.error(f"Error leyendo los archivos: {e}")
        st.stop()

    st.sidebar.title("Navegación")
    pag = st.sidebar.radio("Dashboard", ["Recepción", "Control de estimados", "Proceso"], label_visibility="collapsed")
    if st.sidebar.button("🔄 Recargar datos"):
        st.cache_data.clear()
        st.rerun()
    mt = max(os.path.getmtime(p_rec), os.path.getmtime(p_proc))
    st.sidebar.caption(f"Archivos actualizados: {datetime.fromtimestamp(mt, TZ):%d-%m-%Y %H:%M}")
    if int(rec.reclasificado.sum()):
        st.sidebar.caption(f"♻️ {int(rec.reclasificado.sum())} pallets rechazados reclasificados (procesados).")
    st.sidebar.divider()

    if pag == "Control de estimados":
        ref = rec.fecha.max() if len(rec) else pd.Timestamp(datetime.now(TZ).date())
        base = ref.year if ref.month >= 9 else ref.year - 1
        anio = st.sidebar.selectbox("Temporada (año de inicio)", [base - 1, base, base + 1], index=1,
                                    format_func=lambda a: f"{a}-{str(a + 1)[-2:]}")
        pagina_control(rec, est, anio)
    elif pag == "Recepción":
        pagina_recepcion(rec)
    else:
        pagina_proceso(rec, proc)


if __name__ == "__main__":
    main()
