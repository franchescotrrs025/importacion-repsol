import streamlit as st
import pandas as pd
import re, io
from datetime import date, datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Importación Guardias Repsol", page_icon="🛢️", layout="wide")

st.markdown("""
<style>
.main-header{background:linear-gradient(90deg,#1F4E79,#2E75B6);padding:1.2rem 2rem;border-radius:10px;color:white;margin-bottom:1.5rem}
.main-header h1{margin:0;font-size:1.8rem}
.main-header p{margin:0.3rem 0 0 0;opacity:.85;font-size:.95rem}
.metric-card{background:#f0f4f8;border-left:4px solid #2E75B6;padding:.8rem 1rem;border-radius:6px;margin-bottom:.5rem}
.metric-card .val{font-size:2rem;font-weight:bold;color:#1F4E79}
.metric-card .lbl{font-size:.8rem;color:#555}
.warn-box{background:#fff3cd;border-left:4px solid #ffc107;padding:.7rem 1rem;border-radius:6px;margin-top:.5rem}
.ok-box{background:#d4edda;border-left:4px solid #28a745;padding:.7rem 1rem;border-radius:6px;margin-top:.5rem}
.step-badge{background:#1F4E79;color:white;border-radius:50%;padding:.2rem .6rem;font-weight:bold;margin-right:.5rem;font-size:.9rem}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>🛢️ Importación de Guardias — Repsol Nuevo Lote 57</h1>
    <p>Sube los archivos de guardias y activos para generar el Excel de importación al sistema.</p>
</div>
""", unsafe_allow_html=True)

# ── CONSTANTES ────────────────────────────────────────
CODIGOS = {
    "CHEF":523,"COCINERO":523,"PANADERO PASTELERO":628,"PAN/PASTELERO":628,
    "PANADEROS":628,"AYUDANTE DE COCINA":523,"AYD.COCINA":523,
    "MOZO":629,"AZAFATA":629,"VAJILLERO":528,
    "AUXILIAR DE HOTELERIA":524,"HOTELERO":524,"LIDER DE HOTELERIA":524,
    "LIDER DE GRUPO":524,"LAVANDERO":525,
    "ALMACENERO":630,"AUXILIAR DE ALMACEN":630,"AUX. ALMACEN":630,
    "SUPERVISOR DE SEGURIDAD":528,"SUPERVISOR DE CALIDAD":528,
    "SUP.CALIDAD":528,"SUP.HSE":528,"JEFE CONTRATO":528,
    "JEFE DE CAMPAMENTO - TRAINING":528,"AUXILIAR ADMINISTRATIVO":528,
    "JEFE DE CONTRATO":528,"JEFE DE CAMPAMENTO - TRAINING":528,
    "AUXILIAR DE ADMINISTRACION":528,"NUTRICIONISTA":528,"INSTRUCTORA":528,
}
MESES = {'ene':1,'feb':2,'mar':3,'abr':4,'may':5,'jun':6,
         'jul':7,'ago':8,'sep':9,'oct':10,'nov':11,'dic':12}
ANIO  = datetime.now().year
STOP  = {"PERSONAL EN CAMPO","PERSONAL DESCANSO","TOTAL LOGISTICA",
         "TOTAL CATERING","TOTAL HOTELERIA","TOTALES"}

# ── HELPERS ───────────────────────────────────────────
def parsear_rango(texto):
    t = str(texto).strip().lower()
    m = re.match(r'(\d+)\s+al\s+(\d+)\s+(\w{3})', t)
    if not m: return []
    d_ini, d_fin, mes_str = int(m.group(1)), int(m.group(2)), m.group(3)
    mes = MESES.get(mes_str)
    if not mes: return []
    fecha_fin = date(ANIO, mes, d_fin)
    mes_ini = (mes-1) if d_fin < d_ini else mes
    anio_ini = ANIO if mes_ini > 0 else ANIO-1
    if mes_ini == 0: mes_ini = 12
    fecha_ini = date(anio_ini, mes_ini, d_ini)
    dias = []
    f = fecha_ini
    while f <= fecha_fin:
        dias.append(f); f += timedelta(days=1)
    return dias

def a_codigo(val, cargo):
    if pd.isna(val) or val is None: return -1
    s = str(val).strip().upper()
    if s in ["","NAN","0","0.0"]: return -1
    if "VACACION" in s or s.startswith("SSS"): return -1
    if s in ["1","1.0"]: return CODIGOS.get(cargo.upper(), 528)
    return -1

def cargar_activos(bytes_io):
    df = pd.read_excel(bytes_io)
    # Buscar columna de nombre
    col_nombre = next((c for c in df.columns if "NOMBRE" in c.upper()), None)
    col_dni    = next((c for c in df.columns if "DOCUMENTO" in c.upper() or "DNI" in c.upper()), None)
    if not col_nombre or not col_dni:
        raise ValueError(f"El archivo de Activos debe tener columnas de NOMBRE y DOCUMENTO. Columnas encontradas: {list(df.columns)}")
    df["_nc"] = (df[col_nombre].astype(str).str.upper().str.strip()
                 .str.replace(r"[,.\-]","",regex=True).str.replace(r"\s+"," ",regex=True))
    df["_dni"] = df[col_dni].astype(str).str.strip()
    return df

def buscar_dni(nombre, activos_df):
    n = re.sub(r"[,.\-]","", str(nombre).upper().strip())
    n = re.sub(r"\s+"," ", n).strip()
    if not n or n in ["NAN",""]: return "NO ENCONTRADO"
    m = activos_df[activos_df["_nc"]==n]
    if not m.empty: return m.iloc[0]["_dni"]
    palabras = set(n.split())
    for _, row in activos_df.iterrows():
        if len(palabras)>=2 and palabras==set(row["_nc"].split()):
            return row["_dni"]
    return "NO ENCONTRADO"

def es_nombre_valido(texto):
    """Filtra valores que no son nombres reales."""
    if not texto or texto in ["NAN","","0","0.0"]: return False
    if texto.startswith("SSS"): return False
    if texto == "HOTELERIA": return False
    try: float(texto); return False
    except: pass
    return True

def parsear_hoja_unica(df, activos_df):
    """Parsea un archivo donde todos los bloques están en una sola hoja."""
    resultados = []; i = 0
    while i < len(df):
        row = df.iloc[i]
        c0 = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip().upper() if pd.notna(row.iloc[1]) else ""

        if ("GUARDI" in c0 or "GUADI" in c0) and c1 == "CARGO":
            # Detectar columnas de fecha
            col_fechas = {}
            for ci, val in row.items():
                if isinstance(val, (datetime, pd.Timestamp)):
                    col_fechas[ci] = [pd.Timestamp(val).date()]
                elif pd.notna(val) and isinstance(val, str):
                    dias = parsear_rango(val)
                    if dias: col_fechas[ci] = dias
            if not col_fechas: i += 1; continue

            i += 1; guardia = None
            while i < len(df):
                fila = df.iloc[i]
                g     = str(fila.iloc[0]).strip() if pd.notna(fila.iloc[0]) else ""
                cargo = str(fila.iloc[1]).strip().upper() if pd.notna(fila.iloc[1]) else ""
                nom_h = str(fila.iloc[2]).strip().upper() if pd.notna(fila.iloc[2]) else ""

                if nom_h in STOP or cargo in STOP: break
                if ("GUARDI" in g.upper() or "GUADI" in g.upper()) and cargo == "CARGO": break

                if g and g.upper() not in ["NAN",""]: guardia = g
                if not cargo or cargo in ["NAN","APELLIDOS Y NOMBRES"]: i += 1; continue
                if not es_nombre_valido(nom_h): i += 1; continue

                dias = {}
                for ci, fechas_lista in col_fechas.items():
                    val = fila.iloc[ci] if ci < len(fila) else None
                    for f in fechas_lista: dias[f] = val

                resultados.append({
                    "guardia": guardia, "cargo": cargo, "nombre": nom_h,
                    "dni": buscar_dni(nom_h, activos_df),
                    "dias": dias,
                })
                i += 1
            continue
        i += 1
    return resultados

def procesar(guardias_bytes, activos_bytes):
    activos_df = cargar_activos(activos_bytes)
    xl = pd.ExcelFile(guardias_bytes)
    hojas = xl.sheet_names

    todas = []; fechas_set = set()
    for hoja in hojas:
        if hoja.upper() in ["ESTRUCTURA"]: continue
        df = pd.read_excel(guardias_bytes, sheet_name=hoja, header=None)
        filas = parsear_hoja_unica(df, activos_df)
        for f in filas: fechas_set.update(f["dias"].keys())
        todas.extend(filas)

    return todas, sorted(fechas_set)

def generar_excel(todas, fechas_ord):
    wb = Workbook(); ws = wb.active; ws.title = "Importacion"
    hf  = PatternFill("solid", start_color="1F4E79")
    sf  = PatternFill("solid", start_color="2E75B6")
    df2 = PatternFill("solid", start_color="D9D9D9")
    af  = PatternFill("solid", start_color="E2EFDA")
    ef  = PatternFill("solid", start_color="FFDCE0")
    hfont = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    dfont = Font(name="Arial", size=9)
    bfont = Font(bold=True, name="Arial", size=9)
    gfont = Font(bold=True, name="Arial", size=9, color="375623")
    rfont = Font(bold=True, name="Arial", size=9, color="C00000")
    cen = Alignment(horizontal="center", vertical="center")
    lft = Alignment(horizontal="left",   vertical="center")
    th  = Side(border_style="thin", color="BFBFBF")
    brd = Border(left=th, right=th, top=th, bottom=th)

    lc = get_column_letter(4+len(fechas_ord))
    ws.merge_cells(f"A1:{lc}1")
    ws["A1"] = "IMPORTACIÓN GUARDIAS REPSOL — NUEVO LOTE 57"
    ws["A1"].font = Font(bold=True, color="FFFFFF", name="Arial", size=12)
    ws["A1"].fill = hf; ws["A1"].alignment = cen; ws.row_dimensions[1].height = 22

    ws.merge_cells(f"A2:{lc}2")
    rng = f"Período: {fechas_ord[0].strftime('%d/%m/%Y')} al {fechas_ord[-1].strftime('%d/%m/%Y')}" if fechas_ord else ""
    ws["A2"] = rng; ws["A2"].font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    ws["A2"].fill = sf; ws["A2"].alignment = cen; ws.row_dimensions[2].height = 15

    hdrs = ["APELLIDOS Y NOMBRES","DNI","CARGO","DIRECCIÓN ID"] + [f.strftime("%d/%m/%Y") for f in fechas_ord]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font = hfont; c.fill = sf; c.alignment = cen; c.border = brd
    ws.row_dimensions[3].height = 18
    ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 26; ws.column_dimensions["D"].width = 14
    for i in range(len(fechas_ord)):
        ws.column_dimensions[get_column_letter(5+i)].width = 10

    re2 = 4
    for e in todas:
        cargo = e["cargo"]; dni = e["dni"]
        cod = CODIGOS.get(cargo.upper(), 528)
        vals = [e["nombre"], dni, cargo, cod] + [a_codigo(e["dias"].get(f), cargo) for f in fechas_ord]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=re2, column=ci, value=v); c.border = brd; c.font = dfont
            if ci <= 4:
                c.alignment = lft
                if ci == 1: c.font = bfont
                if ci == 2 and dni == "NO ENCONTRADO": c.fill = ef; c.font = rfont
            else:
                c.alignment = cen
                if v == -1: c.fill = df2
                else: c.fill = af; c.font = gfont
        re2 += 1

    ws.cell(row=re2, column=1, value=f"TOTAL: {re2-4} registros").font = bfont
    ws.freeze_panes = "E4"

    ws2 = wb.create_sheet("Códigos de Dirección")
    for ci, h in enumerate(["CARGO","CÓDIGO","ÁREA"],1):
        c = ws2.cell(row=1,column=ci,value=h); c.font=hfont; c.fill=sf; c.alignment=cen; c.border=brd
    ws2.column_dimensions["A"].width=35; ws2.column_dimensions["B"].width=12; ws2.column_dimensions["C"].width=30
    adesc = {523:"Cocina T.R.2",524:"Hotelería",525:"Lavandería",528:"Vajilla/Admin",629:"Mozo",630:"Almacén",628:"Panadería T.R.2"}
    for i,(k,v) in enumerate(CODIGOS.items(),2):
        ws2.cell(row=i,column=1,value=k).border=brd
        c=ws2.cell(row=i,column=2,value=v); c.border=brd; c.alignment=cen
        ws2.cell(row=i,column=3,value=adesc.get(v,"")).border=brd

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── HISTORIAL ─────────────────────────────────────────
if "historial" not in st.session_state:
    st.session_state.historial = []

# ── SIDEBAR ───────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Historial de importaciones")
    if st.session_state.historial:
        for h in reversed(st.session_state.historial[-10:]):
            st.markdown(f"""
            <div style='background:#f8f9fa;border-radius:6px;padding:.5rem .7rem;margin-bottom:.4rem;font-size:.82rem;border-left:3px solid #2E75B6'>
            🕐 <b>{h['fecha']}</b><br>
            👥 {h['registros']} registros · 📅 {h['dias']} días<br>
            ⚠️ {h['sin_dni']} sin DNI
            </div>""", unsafe_allow_html=True)
    else:
        st.info("Sin importaciones aún.")
    st.markdown("---")
    st.markdown("### 🗂️ Códigos de dirección")
    adesc = {523:"Cocina",524:"Hotelería",525:"Lavandería",528:"Vajilla/Admin",629:"Mozo",630:"Almacén",628:"Panadería"}
    codigos_vistos = set()
    for cargo, cod in CODIGOS.items():
        if cod not in codigos_vistos:
            st.markdown(f"<small>`{cod}` — {adesc.get(cod,'')} </small>", unsafe_allow_html=True)
            codigos_vistos.add(cod)

# ── MAIN ──────────────────────────────────────────────
st.markdown('<span class="step-badge">1</span> **Sube los archivos Excel**', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    f_guardias = st.file_uploader("📁 Archivo de Guardias Repsol", type=["xlsx"],
        help="El archivo con las guardias (puede tener una o varias hojas)")
with col2:
    f_activos = st.file_uploader("📁 Archivo de Activos (con DNIs)", type=["xlsx"],
        help="El Excel con la lista de trabajadores activos y sus DNIs")

st.markdown("---")

if f_guardias and f_activos:
    st.markdown('<span class="step-badge">2</span> **Procesando archivos...**', unsafe_allow_html=True)
    with st.spinner("Leyendo guardias y buscando DNIs..."):
        try:
            guardias_bytes = io.BytesIO(f_guardias.read())
            activos_bytes  = io.BytesIO(f_activos.read())
            todas, fechas_ord = procesar(guardias_bytes, activos_bytes)
        except Exception as ex:
            st.error(f"❌ Error: {ex}")
            try:
                guardias_bytes.seek(0)
                xl = pd.ExcelFile(guardias_bytes)
                st.warning(f"📋 Hojas en el archivo de guardias: {xl.sheet_names}")
            except: pass
            st.stop()

    if not todas:
        st.warning("No se encontraron registros. Verificá que el archivo tenga el formato correcto (columnas: GUARDIAS, CARGO, APELLIDOS Y NOMBRES, fechas).")
        st.stop()

    sin_dni = [e for e in todas if e["dni"] == "NO ENCONTRADO"]

    st.markdown('<span class="step-badge">3</span> **Resumen**', unsafe_allow_html=True)
    m1,m2,m3,m4 = st.columns(4)
    with m1:
        st.markdown(f'<div class="metric-card"><div class="val">{len(todas)}</div><div class="lbl">👥 Registros</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="metric-card"><div class="val">{len(fechas_ord)}</div><div class="lbl">📅 Días</div></div>', unsafe_allow_html=True)
    with m3:
        color = "#C00000" if sin_dni else "#28a745"
        st.markdown(f'<div class="metric-card"><div class="val" style="color:{color}">{len(sin_dni)}</div><div class="lbl">⚠️ Sin DNI</div></div>', unsafe_allow_html=True)
    with m4:
        rango = f"{fechas_ord[0].strftime('%d/%m')} → {fechas_ord[-1].strftime('%d/%m/%Y')}" if fechas_ord else "—"
        st.markdown(f'<div class="metric-card"><div class="val" style="font-size:1rem;padding-top:.5rem">{rango}</div><div class="lbl">📆 Período</div></div>', unsafe_allow_html=True)

    if sin_dni:
        nombres_sin = [e["nombre"] for e in sin_dni]
        st.markdown(f'<div class="warn-box">⚠️ <b>{len(nombres_sin)} sin DNI:</b><br>{"<br>".join(f"• {n}" for n in nombres_sin)}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="ok-box">✅ <b>Todos los DNIs encontrados.</b></div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<span class="step-badge">4</span> **Vista previa**', unsafe_allow_html=True)

    preview = []
    for e in todas:
        activos_c = sum(1 for f in fechas_ord if a_codigo(e["dias"].get(f), e["cargo"]) != -1)
        preview.append({
            "Nombre": e["nombre"], "DNI": e["dni"], "Cargo": e["cargo"],
            "Cód. Dirección": CODIGOS.get(e["cargo"].upper(), 528),
            "Días activo": activos_c, "Días descanso": len(fechas_ord)-activos_c,
        })

    df_prev = pd.DataFrame(preview)
    fc1, fc2 = st.columns([2,1])
    with fc2:
        solo_sin = st.checkbox("Solo sin DNI")
    df_show = df_prev[df_prev["DNI"]=="NO ENCONTRADO"] if solo_sin else df_prev

    def hl(row): return ["background-color:#FFDCE0"]*len(row) if row["DNI"]=="NO ENCONTRADO" else [""]*len(row)
    st.dataframe(df_show.style.apply(hl, axis=1), use_container_width=True, height=350)

    st.markdown("---")
    st.markdown('<span class="step-badge">5</span> **Descargar Excel**', unsafe_allow_html=True)

    guardias_bytes.seek(0)
    excel_buf = generar_excel(todas, fechas_ord)
    fecha_hoy = datetime.now().strftime("%Y%m%d_%H%M")

    st.download_button(
        label="⬇️ Descargar Excel de Importación",
        data=excel_buf,
        file_name=f"Importacion_Repsol_{fecha_hoy}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary"
    )

    entrada = {"fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
               "registros": len(todas), "dias": len(fechas_ord), "sin_dni": len(sin_dni)}
    if not st.session_state.historial or st.session_state.historial[-1]["fecha"] != entrada["fecha"]:
        st.session_state.historial.append(entrada)
else:
    st.info("👆 Sube ambos archivos para comenzar.")
    st.markdown("""
    **¿Qué hace esta app?**
    1. Lee el archivo de **guardias** (cualquier formato Repsol)
    2. Busca el **DNI** de cada trabajador en el archivo de Activos
    3. Reemplaza los `1` por el **código de dirección** del área
    4. Los vacíos y vacaciones → **-1** (descanso)
    5. Genera el **Excel listo para importar**
    """)
