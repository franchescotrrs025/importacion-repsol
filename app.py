import streamlit as st
import pandas as pd
import re, json, io
from datetime import date, datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── CONFIGURACIÓN DE PÁGINA ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Importación Guardias Repsol",
    page_icon="🛢️",
    layout="wide",
)

# ─── ESTILOS CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1F4E79, #2E75B6);
        padding: 1.2rem 2rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 1.5rem;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; }
    .main-header p  { margin: 0.3rem 0 0 0; opacity: 0.85; font-size: 0.95rem; }
    .metric-card {
        background: #f0f4f8;
        border-left: 4px solid #2E75B6;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 0.5rem;
    }
    .metric-card .val { font-size: 2rem; font-weight: bold; color: #1F4E79; }
    .metric-card .lbl { font-size: 0.8rem; color: #555; }
    .warn-box {
        background: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 0.7rem 1rem;
        border-radius: 6px;
        margin-top: 0.5rem;
    }
    .ok-box {
        background: #d4edda;
        border-left: 4px solid #28a745;
        padding: 0.7rem 1rem;
        border-radius: 6px;
        margin-top: 0.5rem;
    }
    .step-badge {
        background: #1F4E79;
        color: white;
        border-radius: 50%;
        padding: 0.2rem 0.6rem;
        font-weight: bold;
        margin-right: 0.5rem;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# ─── HEADER ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🛢️ Importación de Guardias — Repsol Nuevo Lote 57</h1>
    <p>Sube los archivos de guardias y activos para generar el Excel de importación al sistema.</p>
</div>
""", unsafe_allow_html=True)

# ─── CONSTANTES ───────────────────────────────────────────────────────────────
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
    "NUTRICIONISTA":528,"INSTRUCTORA":528,
}
MESES = {'ene':1,'feb':2,'mar':3,'abr':4,'may':5,'jun':6,
         'jul':7,'ago':8,'sep':9,'oct':10,'nov':11,'dic':12}
ANIO  = datetime.now().year
STOP  = {"PERSONAL EN CAMPO","PERSONAL DESCANSO","TOTAL LOGISTICA",
         "TOTAL CATERING","TOTAL HOTELERIA","TOTALES"}

# ─── FUNCIONES CORE ───────────────────────────────────────────────────────────
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

def cargar_activos(bytes_io):
    df = pd.read_excel(bytes_io)
    df["_nc"] = (df["APELLIDOS Y NOMBRES"].astype(str).str.upper().str.strip()
                 .str.replace(r"[,.\-]","",regex=True).str.replace(r"\s+"," ",regex=True))
    df["NRO_ DOCUMENTO"] = df["NRO_ DOCUMENTO"].astype(str).str.strip()
    return df

def buscar_dni(nombre, activos_df):
    n = re.sub(r"[,.\-]","", str(nombre).upper().strip())
    n = re.sub(r"\s+"," ", n).strip()
    if not n or n in ["NAN",""]: return "NO ENCONTRADO"
    m = activos_df[activos_df["_nc"]==n]
    if not m.empty: return m.iloc[0]["NRO_ DOCUMENTO"]
    palabras = set(n.split())
    for _, row in activos_df.iterrows():
        if len(palabras)>=2 and palabras==set(row["_nc"].split()):
            return row["NRO_ DOCUMENTO"]
    return "NO ENCONTRADO"

def detectar_hoja(guardias_bytes, keywords):
    """Busca la hoja cuyo nombre contenga alguna keyword (sin importar mayúsculas/espacios)."""
    try:
        xl = pd.ExcelFile(guardias_bytes)
        hojas = xl.sheet_names
    except:
        return None
    for hoja in hojas:
        if any(k.lower() in hoja.lower() for k in keywords):
            return hoja
    return None

def cargar_pax(guardias_bytes):
    hoja_pax = detectar_hoja(guardias_bytes, ["pax", "sdx", "personal"])
    if hoja_pax is None:
        xl = pd.ExcelFile(guardias_bytes)
        raise ValueError(f"No se encontró la hoja de personal (Pax SDX - NM).\nHojas disponibles: {xl.sheet_names}")
    pax_raw = pd.read_excel(guardias_bytes, sheet_name=hoja_pax, header=None)
    pax_map = {}; sec = None
    for _, row in pax_raw.iterrows():
        c1 = str(row[1]).strip().upper() if pd.notna(row[1]) else ""
        if c1 in ["CATERING","HOTELERIA"]: sec=c1; continue
        try: int(row[0])
        except: continue
        nombre = str(row[1]).strip().upper() if pd.notna(row[1]) else ""
        cargo  = str(row[2]).strip().upper() if pd.notna(row[2]) else ""
        if not nombre or nombre=="NAN": continue
        pax_map.setdefault((cargo,sec),[]).append(nombre)
    return pax_map

def parsear_hoja(guardias_bytes, hoja, sec_pax, pax_map, pax_ptr, activos_df):
    # Buscar hoja por nombre exacto o aproximado
    xl = pd.ExcelFile(guardias_bytes)
    hoja_real = hoja  # default
    for h in xl.sheet_names:
        if h.strip().upper() == hoja.strip().upper():
            hoja_real = h; break
    if hoja_real not in xl.sheet_names:
        return []  # Hoja no encontrada, saltar
    df = pd.read_excel(guardias_bytes, sheet_name=hoja_real, header=None)
    resultados=[]; i=0
    while i<len(df):
        row=df.iloc[i]
        c0=str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""
        c1=str(row.iloc[1]).strip().upper() if pd.notna(row.iloc[1]) else ""
        if ("GUARDI" in c0 or "GUADI" in c0) and c1=="CARGO":
            col_fechas={}
            for ci, val in row.items():
                if isinstance(val,(datetime,pd.Timestamp)):
                    col_fechas[ci]=[pd.Timestamp(val).date()]
                elif pd.notna(val) and isinstance(val,str):
                    dias=parsear_rango(val)
                    if dias: col_fechas[ci]=dias
            if not col_fechas: i+=1; continue
            i+=1; guardia=None
            while i<len(df):
                fila=df.iloc[i]
                g=str(fila.iloc[0]).strip() if pd.notna(fila.iloc[0]) else ""
                cargo=str(fila.iloc[1]).strip().upper() if pd.notna(fila.iloc[1]) else ""
                nom_h=str(fila.iloc[2]).strip().upper() if pd.notna(fila.iloc[2]) else ""
                if nom_h in STOP or cargo in STOP: break
                if ("GUARDI" in g.upper() or "GUADI" in g.upper()) and cargo=="CARGO": break
                if g and g.upper() not in ["NAN",""]: guardia=g
                if not cargo or cargo in ["NAN","APELLIDOS Y NOMBRES"]: i+=1; continue
                # Nombre
                if nom_h and nom_h not in ["NAN","APELLIDOS Y NOMBRES",""]:
                    nombre=nom_h
                else:
                    k=(cargo,sec_pax)
                    if k in pax_map and pax_ptr[k]<len(pax_map[k]):
                        nombre=pax_map[k][pax_ptr[k]]; pax_ptr[k]+=1
                    else:
                        nombre=f"SIN NOMBRE ({cargo})"
                dias={}
                for ci, fechas_lista in col_fechas.items():
                    val=fila.iloc[ci] if ci<len(fila) else None
                    for f in fechas_lista: dias[f]=val
                resultados.append({
                    "guardia":guardia,"cargo":cargo,"nombre":nombre,
                    "dni":buscar_dni(nombre,activos_df),"dias":dias,"hoja":hoja
                })
                i+=1
            continue
        i+=1
    return resultados

def a_codigo(val, cargo):
    if pd.isna(val) or val is None: return -1
    s=str(val).strip().upper()
    if s in ["","NAN","0","0.0"]: return -1
    if "VACACION" in s: return -1
    if s in ["1","1.0"]: return CODIGOS.get(cargo,528)
    return -1

def procesar(guardias_bytes, activos_bytes):
    activos_df = cargar_activos(activos_bytes)
    pax_map    = cargar_pax(guardias_bytes)
    pax_ptr    = {k:0 for k in pax_map}
    HOJAS      = {"CATERING":"CATERING","HOTELERIA":"HOTELERIA","ADM - LOG":"CATERING"}
    todas=[]; fechas_set=set()
    for hoja, sec in HOJAS.items():
        filas = parsear_hoja(guardias_bytes, hoja, sec, pax_map, pax_ptr, activos_df)
        todas.extend(filas)
        for r in filas: fechas_set.update(r["dias"].keys())
    return todas, sorted(fechas_set), activos_df

def generar_excel(todas, fechas_ord):
    wb=Workbook(); ws=wb.active; ws.title="Importacion"
    hf=PatternFill("solid",start_color="1F4E79")
    sf=PatternFill("solid",start_color="2E75B6")
    df2=PatternFill("solid",start_color="D9D9D9")
    af=PatternFill("solid",start_color="E2EFDA")
    ef=PatternFill("solid",start_color="FFDCE0")
    hfont=Font(bold=True,color="FFFFFF",name="Arial",size=10)
    dfont=Font(name="Arial",size=9)
    bfont=Font(bold=True,name="Arial",size=9)
    gfont=Font(bold=True,name="Arial",size=9,color="375623")
    rfont=Font(bold=True,name="Arial",size=9,color="C00000")
    cen=Alignment(horizontal="center",vertical="center")
    lft=Alignment(horizontal="left",vertical="center")
    th=Side(border_style="thin",color="BFBFBF")
    brd=Border(left=th,right=th,top=th,bottom=th)

    lc=get_column_letter(4+len(fechas_ord))
    ws.merge_cells(f"A1:{lc}1")
    ws["A1"]="IMPORTACIÓN GUARDIAS REPSOL — NUEVO LOTE 57"
    ws["A1"].font=Font(bold=True,color="FFFFFF",name="Arial",size=12)
    ws["A1"].fill=hf; ws["A1"].alignment=cen; ws.row_dimensions[1].height=22

    ws.merge_cells(f"A2:{lc}2")
    rng=f"Período: {fechas_ord[0].strftime('%d/%m/%Y')} al {fechas_ord[-1].strftime('%d/%m/%Y')}" if fechas_ord else ""
    ws["A2"]=rng; ws["A2"].font=Font(bold=True,color="FFFFFF",name="Arial",size=9)
    ws["A2"].fill=sf; ws["A2"].alignment=cen; ws.row_dimensions[2].height=15

    hdrs=["APELLIDOS Y NOMBRES","DNI","CARGO","ÁREA"]+[f.strftime("%d/%m/%Y") for f in fechas_ord]
    for ci,h in enumerate(hdrs,1):
        c=ws.cell(row=3,column=ci,value=h); c.font=hfont; c.fill=sf; c.alignment=cen; c.border=brd
    ws.row_dimensions[3].height=18
    ws.column_dimensions["A"].width=34; ws.column_dimensions["B"].width=13
    ws.column_dimensions["C"].width=26; ws.column_dimensions["D"].width=12
    for i in range(len(fechas_ord)): ws.column_dimensions[get_column_letter(5+i)].width=10

    re2=4
    for e in todas:
        nm=e["nombre"]; dni=e["dni"]; cargo=e["cargo"]; hj=e["hoja"]
        area={"CATERING":"CATERING","HOTELERIA":"HOTELERÍA"}.get(hj,"LOGÍSTICA")
        vals=[nm,dni,cargo,area]+[a_codigo(e["dias"].get(f),cargo) for f in fechas_ord]
        for ci,v in enumerate(vals,1):
            c=ws.cell(row=re2,column=ci,value=v); c.border=brd; c.font=dfont
            if ci<=4:
                c.alignment=lft
                if ci==1: c.font=bfont
                if ci==2 and dni=="NO ENCONTRADO": c.fill=ef; c.font=rfont
            else:
                c.alignment=cen
                if v==-1: c.fill=df2
                else: c.fill=af; c.font=gfont
        re2+=1
    ws.cell(row=re2,column=1,value=f"TOTAL: {re2-4} registros").font=bfont
    ws.freeze_panes="E4"

    # Hoja leyenda
    ws2=wb.create_sheet("Códigos de Dirección")
    for ci,h in enumerate(["CARGO","CÓDIGO","ÁREA SISTEMA"],1):
        c=ws2.cell(row=1,column=ci,value=h); c.font=hfont; c.fill=sf; c.alignment=cen; c.border=brd
    ws2.column_dimensions["A"].width=35; ws2.column_dimensions["B"].width=12; ws2.column_dimensions["C"].width=30
    adesc={523:"Cocina T.R.2",524:"Hotelería",525:"Lavandería",528:"Vajilla/Admin",629:"Mozo",630:"Almacén",628:"Panadería T.R.2"}
    for i,(k,v) in enumerate(CODIGOS.items(),2):
        ws2.cell(row=i,column=1,value=k).border=brd
        c=ws2.cell(row=i,column=2,value=v); c.border=brd; c.alignment=cen
        ws2.cell(row=i,column=3,value=adesc.get(v,"")).border=brd

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── HISTORIAL ────────────────────────────────────────────────────────────────
if "historial" not in st.session_state:
    st.session_state.historial = []

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Historial de importaciones")
    if st.session_state.historial:
        for h in reversed(st.session_state.historial[-10:]):
            st.markdown(f"""
            <div style='background:#f8f9fa;border-radius:6px;padding:0.5rem 0.7rem;margin-bottom:0.4rem;font-size:0.82rem;border-left:3px solid #2E75B6'>
            🕐 <b>{h['fecha']}</b><br>
            👥 {h['registros']} registros · 📅 {h['dias']} días<br>
            ⚠️ {h['sin_dni']} sin DNI
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Sin importaciones aún.")

    st.markdown("---")
    st.markdown("### 🗂️ Códigos de dirección")
    area_desc = {
        523:"Cocina T.R.2", 524:"Hotelería", 525:"Lavandería",
        528:"Vajilla/Admin", 629:"Mozo", 630:"Almacén", 628:"Panadería T.R.2"
    }
    for cargo, cod in list(CODIGOS.items())[:10]:
        st.markdown(f"<small>**{cargo}** → `{cod}`</small>", unsafe_allow_html=True)
    st.caption("Ver hoja 'Códigos' en el Excel generado para lista completa.")

# ─── MAIN ─────────────────────────────────────────────────────────────────────
st.markdown('<span class="step-badge">1</span> **Sube los archivos Excel**', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    f_guardias = st.file_uploader(
        "📁 Archivo de Guardias Repsol",
        type=["xlsx"],
        help="El archivo con las hojas: CATERING, HOTELERIA, ADM - LOG, Pax SDX - NM"
    )
with col2:
    f_activos = st.file_uploader(
        "📁 Archivo Activos (Activos al XX.XX)",
        type=["xlsx"],
        help="El Excel con la lista de trabajadores activos y sus DNIs"
    )

st.markdown("---")

if f_guardias and f_activos:
    st.markdown('<span class="step-badge">2</span> **Procesando archivos...**', unsafe_allow_html=True)

    with st.spinner("Leyendo guardias, asignando nombres y buscando DNIs..."):
        try:
            guardias_bytes = io.BytesIO(f_guardias.read())
            activos_bytes  = io.BytesIO(f_activos.read())
            todas, fechas_ord, activos_df = procesar(guardias_bytes, activos_bytes)
        except Exception as ex:
            st.error(f"❌ Error al procesar los archivos: {ex}")
            # Mostrar hojas disponibles para ayudar al diagnóstico
            try:
                guardias_bytes.seek(0)
                xl = pd.ExcelFile(guardias_bytes)
                st.warning(f"📋 Hojas encontradas en tu archivo de guardias: {xl.sheet_names}")
                st.info("Verificá que el archivo tenga las hojas: CATERING, HOTELERIA, ADM - LOG, Pax SDX - NM")
            except:
                pass
            st.stop()

    if not todas:
        st.warning("No se encontraron registros. Verificá que el archivo de guardias tenga el formato correcto.")
        st.stop()

    # Métricas
    sin_dni   = [e for e in todas if e["dni"]=="NO ENCONTRADO"]
    sin_nomb  = [e for e in todas if "SIN NOMBRE" in e["nombre"]]
    areas     = {"CATERING":0,"HOTELERIA":0,"ADM - LOG":0}
    for e in todas: areas[e["hoja"]] = areas.get(e["hoja"],0)+1

    st.markdown('<span class="step-badge">3</span> **Resumen del procesamiento**', unsafe_allow_html=True)
    m1,m2,m3,m4 = st.columns(4)
    with m1:
        st.markdown(f"""<div class="metric-card"><div class="val">{len(todas)}</div><div class="lbl">👥 Registros totales</div></div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""<div class="metric-card"><div class="val">{len(fechas_ord)}</div><div class="lbl">📅 Días procesados</div></div>""", unsafe_allow_html=True)
    with m3:
        color = "#C00000" if sin_dni else "#28a745"
        st.markdown(f"""<div class="metric-card"><div class="val" style="color:{color}">{len(sin_dni)}</div><div class="lbl">⚠️ Sin DNI encontrado</div></div>""", unsafe_allow_html=True)
    with m4:
        rango = f"{fechas_ord[0].strftime('%d/%m')} → {fechas_ord[-1].strftime('%d/%m/%Y')}" if fechas_ord else "—"
        st.markdown(f"""<div class="metric-card"><div class="val" style="font-size:1.1rem;padding-top:0.4rem">{rango}</div><div class="lbl">📆 Período cubierto</div></div>""", unsafe_allow_html=True)

    # Detalle por área
    ca, cb, cc = st.columns(3)
    with ca: st.metric("🍳 CATERING",  areas.get("CATERING",0),  "registros")
    with cb: st.metric("🛏️ HOTELERÍA", areas.get("HOTELERIA",0), "registros")
    with cc: st.metric("📦 LOGÍSTICA", areas.get("ADM - LOG",0), "registros")

    # Alertas
    if sin_dni:
        nombres_sin = [e["nombre"] for e in sin_dni if "SIN NOMBRE" not in e["nombre"]]
        if nombres_sin:
            st.markdown(f"""
            <div class="warn-box">
            ⚠️ <b>{len(nombres_sin)} trabajador(es) sin DNI encontrado</b> — aparecerán marcados en rojo en el Excel:<br>
            {'<br>'.join(f'• {n}' for n in nombres_sin)}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown('<div class="ok-box">✅ <b>Todos los DNIs fueron encontrados correctamente.</b></div>', unsafe_allow_html=True)

    if sin_nomb:
        st.markdown(f"""
        <div class="warn-box">
        ⚠️ <b>{len(sin_nomb)} puesto(s) sin nombre asignado</b> en la lista Pax SDX - NM.<br>
        Revisá que todos los puestos tengan su nombre en esa hoja.
        </div>
        """, unsafe_allow_html=True)

    # ── Vista previa ──
    st.markdown("---")
    st.markdown('<span class="step-badge">4</span> **Vista previa de datos**', unsafe_allow_html=True)

    preview_data = []
    for e in todas:
        activos_count = sum(1 for f in fechas_ord if a_codigo(e["dias"].get(f), e["cargo"]) != -1)
        descanso_count = len(fechas_ord) - activos_count
        preview_data.append({
            "Nombre": e["nombre"],
            "DNI": e["dni"],
            "Cargo": e["cargo"],
            "Área": {"CATERING":"🍳 Catering","HOTELERIA":"🛏️ Hotelería"}.get(e["hoja"],"📦 Logística"),
            "Días activo": activos_count,
            "Días descanso": descanso_count,
        })

    df_preview = pd.DataFrame(preview_data)

    # Filtros
    fc1, fc2 = st.columns([2,1])
    with fc1:
        filtro_area = st.multiselect(
            "Filtrar por área:",
            options=df_preview["Área"].unique().tolist(),
            default=df_preview["Área"].unique().tolist()
        )
    with fc2:
        solo_sin_dni = st.checkbox("Mostrar solo sin DNI")

    df_filtered = df_preview[df_preview["Área"].isin(filtro_area)]
    if solo_sin_dni:
        df_filtered = df_filtered[df_filtered["DNI"]=="NO ENCONTRADO"]

    def highlight_dni(row):
        if row["DNI"] == "NO ENCONTRADO":
            return ["background-color: #FFDCE0"]*len(row)
        return [""]*len(row)

    st.dataframe(
        df_filtered.style.apply(highlight_dni, axis=1),
        use_container_width=True,
        height=350
    )

    # ── Descarga ──
    st.markdown("---")
    st.markdown('<span class="step-badge">5</span> **Descargar Excel de importación**', unsafe_allow_html=True)

    guardias_bytes.seek(0)
    excel_buf = generar_excel(todas, fechas_ord)
    fecha_hoy = datetime.now().strftime("%Y%m%d_%H%M")
    nombre_archivo = f"Importacion_Repsol_{fecha_hoy}.xlsx"

    st.download_button(
        label="⬇️ Descargar Excel de Importación",
        data=excel_buf,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

    # Guardar en historial
    entrada_hist = {
        "fecha":     datetime.now().strftime("%d/%m/%Y %H:%M"),
        "registros": len(todas),
        "dias":      len(fechas_ord),
        "sin_dni":   len(sin_dni),
        "archivo":   nombre_archivo,
    }
    if not st.session_state.historial or st.session_state.historial[-1]["fecha"] != entrada_hist["fecha"]:
        st.session_state.historial.append(entrada_hist)

else:
    st.info("👆 Sube ambos archivos para comenzar el procesamiento.")
    st.markdown("""
    **¿Qué hace esta app?**
    1. Lee el archivo de **guardias Repsol** (hojas: CATERING, HOTELERIA, ADM-LOG, Pax SDX-NM)
    2. Busca el **DNI** de cada trabajador en el archivo de Activos
    3. Reemplaza los `1` por el **código de dirección** del área correspondiente
    4. Reemplaza los vacíos y vacaciones por **-1** (descanso)
    5. Genera un **Excel listo para importar** al sistema
    """)
