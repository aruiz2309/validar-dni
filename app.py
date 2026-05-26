import streamlit as st
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import unicodedata
import time
import random
import io

st.set_page_config(
    page_title="Validador DNI Perú",
    page_icon="🪪",
    layout="centered"
)

st.title("🪪 Validador de DNI - Perú")
st.markdown("Sube tu Excel con **NOMBRE COMPLETO** y **DNI**, valida contra eldni.com y dniperu.com, y descarga el resultado.")

# ─── Utilidades ───────────────────────────────────────────────

def normalizar(texto):
    if not texto:
        return ""
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto.upper()).strip()

def similitud(a, b):
    pa = set(normalizar(a).split())
    pb = set(normalizar(b).split())
    if not pa:
        return 0.0
    return len(pa & pb) / len(pa)

def make_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "es-PE,es;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
    })
    return s

# ─── Extraer nombre de HTML ────────────────────────────────────

def extraer_nombre_html(html):
    soup = BeautifulSoup(html, "html.parser")
    ap_pat, ap_mat, nombres = "", "", ""

    # Buscar en tabla
    for fila in soup.find_all("tr"):
        celdas = fila.find_all(["td", "th"])
        if len(celdas) >= 2:
            k = celdas[0].get_text(strip=True).lower()
            v = celdas[1].get_text(strip=True).upper()
            if "paterno" in k:
                ap_pat = v
            elif "materno" in k:
                ap_mat = v
            elif "nombre" in k and "apellido" not in k:
                nombres = v

    if ap_pat or nombres:
        return re.sub(r"\s+", " ", f"{ap_pat} {ap_mat} {nombres}").strip()

    # Buscar en divs/spans con clases típicas de resultado
    for clase in ["resultado", "result", "datos", "nombre", "data"]:
        tag = soup.find(class_=re.compile(clase, re.I))
        if tag:
            txt = tag.get_text(" ", strip=True).upper()
            m = re.search(r"([A-ZÁÉÍÓÚÑ]{2,}(?:\s+[A-ZÁÉÍÓÚÑ]{2,}){2,})", txt)
            if m:
                return m.group(1).strip()

    return ""

# ─── Consulta eldni.com ────────────────────────────────────────

def consultar_eldni(dni, session):
    try:
        url = "https://eldni.com/pe/buscar-datos-por-dni"
        session.headers.update({
            "Referer": "https://eldni.com/",
            "Origin":  "https://eldni.com",
        })
        session.get(url, timeout=15)
        time.sleep(random.uniform(0.5, 1.0))

        r = session.post(url, data={"dni": dni}, timeout=15)
        if r.status_code == 200:
            nombre = extraer_nombre_html(r.text)
            if nombre:
                return nombre

        # Intentar endpoint alternativo
        r2 = session.get(f"https://eldni.com/pe/buscar-datos-por-dni?dni={dni}", timeout=15)
        if r2.status_code == 200:
            return extraer_nombre_html(r2.text)

    except Exception:
        pass
    return ""

# ─── Consulta dniperu.com ──────────────────────────────────────

def consultar_dniperu(dni, session):
    try:
        url = "https://dniperu.com/buscar-dni-nombres-apellidos/"
        session.headers.update({
            "Referer": "https://dniperu.com/",
            "Origin":  "https://dniperu.com",
        })

        # GET para cookies y nonce
        r = session.get(url, timeout=15)
        time.sleep(random.uniform(0.5, 1.0))

        soup = BeautifulSoup(r.text, "html.parser")
        nonce = ""
        for inp in soup.find_all("input", {"type": "hidden"}):
            if "nonce" in (inp.get("name", "") + inp.get("id", "")).lower():
                nonce = inp.get("value", "")
                break

        # POST formulario
        data = {"dni": dni}
        if nonce:
            data["nonce"] = nonce

        r2 = session.post(url, data=data, timeout=15)
        if r2.status_code == 200:
            # Intentar JSON
            try:
                j = r2.json()
                d = j.get("data", j)
                np_ = d.get("apellidoPaterno", d.get("apellido_paterno", ""))
                nm_ = d.get("apellidoMaterno", d.get("apellido_materno", ""))
                nn_ = d.get("nombres", "")
                nc_ = d.get("nombre_completo", d.get("nombreCompleto", ""))
                if nc_:
                    return nc_.upper().strip()
                if np_ or nn_:
                    return re.sub(r"\s+", " ", f"{np_} {nm_} {nn_}").strip().upper()
            except Exception:
                pass

            nombre = extraer_nombre_html(r2.text)
            if nombre:
                return nombre

        # AJAX WordPress
        r3 = session.post(
            "https://dniperu.com/wp-admin/admin-ajax.php",
            data={"action": "buscar_dni", "dni": dni},
            headers={"X-Requested-With": "XMLHttpRequest"},
            timeout=15,
        )
        if r3.status_code == 200:
            try:
                j = r3.json()
                d = j.get("data", j)
                np_ = d.get("apellidoPaterno", d.get("apellido_paterno", ""))
                nm_ = d.get("apellidoMaterno", d.get("apellido_materno", ""))
                nn_ = d.get("nombres", "")
                nc_ = d.get("nombre_completo", d.get("nombreCompleto", ""))
                if nc_:
                    return nc_.upper().strip()
                if np_ or nn_:
                    return re.sub(r"\s+", " ", f"{np_} {nm_} {nn_}").strip().upper()
            except Exception:
                pass
            return extraer_nombre_html(r3.text)

    except Exception:
        pass
    return ""

# ─── Validar un DNI ───────────────────────────────────────────

def validar_dni(nombre_archivo, dni, session):
    # Web 1: eldni.com
    nombre_web = consultar_eldni(dni, session)

    # Web 2: dniperu.com (si web 1 no respondió)
    if not nombre_web:
        time.sleep(random.uniform(1.0, 2.0))
        nombre_web = consultar_dniperu(dni, session)

    if not nombre_web:
        return {"nombre_web": "NO ENCONTRADO", "estado": 0}

    sim = similitud(nombre_archivo, nombre_web)
    estado = 1 if sim >= 0.80 else 0
    return {"nombre_web": nombre_web, "estado": estado}

# ─── Leer Excel subido ────────────────────────────────────────

def leer_excel(archivo_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(archivo_bytes), read_only=True)
    ws = wb.active
    registros = []
    dnis_vistos = set()

    # Detectar fila de cabecera (primera fila con texto)
    fila_inicio = 2
    for i, row in enumerate(ws.iter_rows(max_row=5, values_only=True), start=1):
        vals = [str(v).lower() for v in row if v]
        if any("nombre" in v or "dni" in v for v in vals):
            fila_inicio = i + 1
            break

    for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
        nombre  = str(row[0]).strip() if row[0] else ""
        dni_raw = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        dni     = re.sub(r"\D", "", dni_raw).zfill(8)
        if not dni or not dni.replace("0", ""):
            continue
        if dni in dnis_vistos:
            continue
        dnis_vistos.add(dni)
        registros.append({"nombre": nombre, "dni": dni})

    return registros

# ─── Generar Excel de salida ──────────────────────────────────

def generar_excel(registros, resultados_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultado Validacion"

    thin  = Side(style="thin", color="BFBFBF")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    fill_cab   = PatternFill("solid", fgColor="1F4E79")
    fill_verde = PatternFill("solid", fgColor="E2EFDA")
    fill_rojo  = PatternFill("solid", fgColor="FCE4D6")
    fill_gris  = PatternFill("solid", fgColor="F2F2F2")
    fill_blanc = PatternFill("solid", fgColor="FFFFFF")

    font_cab   = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    font_datos = Font(name="Arial", size=10)
    align_c    = Alignment(horizontal="center", vertical="center")
    align_l    = Alignment(horizontal="left",   vertical="center")

    # Cabecera
    cabeceras = ["NOMBRE COMPLETO", "NRO DNI", "NOMBRE DE LA WEB", "ESTADO"]
    ws.row_dimensions[1].height = 28
    for col, titulo in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=col, value=titulo)
        c.font      = font_cab
        c.fill      = fill_cab
        c.border    = borde
        c.alignment = align_c

    # Datos
    for i, reg in enumerate(registros, start=2):
        ws.row_dimensions[i].height = 18
        dni        = reg["dni"]
        nombre_arc = reg["nombre"]
        r          = resultados_dict.get(dni, {"nombre_web": "NO ENCONTRADO", "estado": 0})
        nombre_web = r["nombre_web"]
        estado     = r["estado"]

        fill_fila = fill_gris if i % 2 == 0 else fill_blanc
        fill_est  = fill_verde if estado == 1 else fill_rojo

        valores = [nombre_arc, dni, nombre_web, estado]
        fills   = [fill_fila, fill_fila, fill_fila, fill_est]
        aligns  = [align_l, align_c, align_l, align_c]

        for col, (val, fill, aln) in enumerate(zip(valores, fills, aligns), 1):
            c = ws.cell(row=i, column=col, value=val)
            c.font      = font_datos
            c.fill      = fill
            c.border    = borde
            c.alignment = aln

    # Anchos
    for col, ancho in enumerate([42, 14, 42, 10], 1):
        ws.column_dimensions[get_column_letter(col)].width = ancho

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── INTERFAZ STREAMLIT ───────────────────────────────────────

archivo = st.file_uploader(
    "📂 Sube tu Excel con dos columnas: **NOMBRE COMPLETO** | **DNI**",
    type=["xlsx"]
)

if archivo:
    datos_bytes = archivo.read()
    registros   = leer_excel(datos_bytes)

    if not registros:
        st.error("❌ No se encontraron datos en el Excel. Verifica que tenga las columnas NOMBRE COMPLETO y DNI.")
    else:
        st.success(f"✅ Se detectaron **{len(registros)} DNIs únicos**")

        col1, col2 = st.columns(2)
        col1.metric("Total DNIs", len(registros))
        mins = (len(registros) * 5) // 60
        segs = (len(registros) * 5) % 60
        col2.metric("Tiempo estimado", f"~{mins}m {segs}s")

        with st.expander("👁️ Vista previa del Excel subido"):
            for r in registros[:10]:
                st.markdown(f"• `{r['dni']}` → {r['nombre']}")
            if len(registros) > 10:
                st.caption(f"... y {len(registros) - 10} más")

        st.warning("⚠️ No cierres esta pestaña mientras se procesa.")

        if st.button("🚀 Iniciar Validación", type="primary", use_container_width=True):

            resultados_dict = {}
            session = make_session()

            barra  = st.progress(0, text="Iniciando...")
            estado_txt = st.empty()

            col_a, col_b, col_c = st.columns(3)
            cnt_ok  = col_a.empty()
            cnt_inc = col_b.empty()
            cnt_no  = col_c.empty()
            ok, inc, no_enc = 0, 0, 0

            for i, reg in enumerate(registros):
                dni    = reg["dni"]
                nombre = reg["nombre"]

                estado_txt.markdown(
                    f"🔍 Procesando **{i+1} de {len(registros)}** — "
                    f"DNI: `{dni}` | `{nombre[:40]}`"
                )

                r = validar_dni(nombre, dni, session)
                resultados_dict[dni] = r

                if r["estado"] == 1:
                    ok += 1
                elif r["nombre_web"] == "NO ENCONTRADO":
                    no_enc += 1
                else:
                    inc += 1

                cnt_ok.metric("✅ Estado 1",  ok)
                cnt_inc.metric("❌ Estado 0 (diferente)", inc)
                cnt_no.metric("⚠️ No encontrado", no_enc)

                barra.progress(
                    (i + 1) / len(registros),
                    text=f"Procesando {i+1} de {len(registros)}..."
                )

                if i < len(registros) - 1:
                    time.sleep(random.uniform(2.5, 4.0))

            barra.progress(1.0, text="¡Completado!")
            estado_txt.success("✅ Validación finalizada")

            # Generar y ofrecer descarga
            excel_buf = generar_excel(registros, resultados_dict)

            st.download_button(
                label="📥 Descargar Excel con resultados",
                data=excel_buf,
                file_name="validacion_dni_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

            # Resumen
            st.markdown("---")
            st.subheader("📊 Resumen")
            total = len(registros)
            c1, c2, c3 = st.columns(3)
            c1.metric("✅ Correctos (1)",   ok,     f"{round(ok/total*100)}%")
            c2.metric("❌ Incorrectos (0)", inc,    f"{round(inc/total*100)}%")
            c3.metric("⚠️ No encontrados",  no_enc, f"{round(no_enc/total*100)}%")

            # Vista previa resultado
            st.markdown("---")
            st.subheader("👁️ Vista previa del resultado")
            st.markdown(
                "| NOMBRE COMPLETO | NRO DNI | NOMBRE DE LA WEB | ESTADO |"
                "\n|---|---|---|---|"
            )
            for reg in registros[:8]:
                dni = reg["dni"]
                r   = resultados_dict.get(dni, {})
                nw  = r.get("nombre_web", "—")
                est = r.get("estado", 0)
                ico = "✅ 1" if est == 1 else "❌ 0"
                st.markdown(f"| {reg['nombre']} | {dni} | {nw} | {ico} |")
            if len(registros) > 8:
                st.caption(f"... y {len(registros)-8} filas más en el Excel descargable.")

