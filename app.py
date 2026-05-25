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
    page_title="Validador de DNI - Perú",
    page_icon="🪪",
    layout="centered"
)

st.title("🪪 Validador de Nombres por DNI")
st.markdown("Sube tu Excel con DNIs, la app consulta las webs y descarga el resultado.")

# ─── Utilidades ───────────────────────────────────────────────────────────────

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

# ─── Consulta eldni.com por DNI ───────────────────────────────────────────────

def consultar_eldni(dni, session):
    resultado = {"nombre": "", "fuente": "eldni.com", "error": ""}
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept-Language": "es-PE,es;q=0.9",
            "Referer": "https://eldni.com/",
            "Origin": "https://eldni.com",
        }
        session.get("https://eldni.com/pe/buscar-datos-por-dni", headers=headers, timeout=15)
        resp = session.post(
            "https://eldni.com/pe/buscar-datos-por-dni",
            data={"dni": dni},
            headers=headers,
            timeout=15,
        )
        if resp.status_code != 200:
            resultado["error"] = f"HTTP {resp.status_code}"
            return resultado

        soup = BeautifulSoup(resp.text, "html.parser")
        ap_pat, ap_mat, nombres = "", "", ""
        for fila in soup.find_all("tr"):
            celdas = fila.find_all(["td", "th"])
            if len(celdas) >= 2:
                k = celdas[0].get_text(strip=True).lower()
                v = celdas[1].get_text(strip=True)
                if "paterno" in k:   ap_pat  = v
                elif "materno" in k: ap_mat  = v
                elif "nombre" in k:  nombres = v

        if ap_pat or nombres:
            resultado["nombre"] = f"{ap_pat} {ap_mat} {nombres}".strip()
        else:
            texto = soup.get_text(" ")
            m = re.search(r'(?:Nombres?|Apellidos?)\s*[:\-]\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ ]{3,})', texto)
            if m:
                resultado["nombre"] = m.group(1).strip()

        if not resultado["nombre"]:
            resultado["error"] = "Sin resultado"
    except Exception as e:
        resultado["error"] = str(e)[:60]
    return resultado

# ─── Consulta buscardniperu.com por nombre ────────────────────────────────────

def consultar_buscardni_nombre(nombre, session):
    resultado = {"nombre": "", "fuente": "buscardniperu.com", "error": ""}
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept-Language": "es-PE,es;q=0.9",
            "Referer": "https://buscardniperu.com/",
            "Origin": "https://buscardniperu.com",
        }
        session.get("https://buscardniperu.com/buscar-dni-por-nombres/", headers=headers, timeout=15)

        # Intentar POST al formulario
        resp = session.post(
            "https://buscardniperu.com/buscar-dni-por-nombres/",
            data={"nombre": nombre, "nombres": nombre},
            headers=headers,
            timeout=15,
        )

        if resp.status_code != 200:
            # Intentar AJAX
            resp = session.post(
                "https://buscardniperu.com/wp-admin/admin-ajax.php",
                data={"action": "buscar_nombre", "nombre": nombre},
                headers={**headers, "X-Requested-With": "XMLHttpRequest"},
                timeout=15,
            )

        soup = BeautifulSoup(resp.text, "html.parser")
        ap_pat, ap_mat, nombres_r = "", "", ""
        for fila in soup.find_all("tr"):
            celdas = fila.find_all(["td", "th"])
            if len(celdas) >= 2:
                k = celdas[0].get_text(strip=True).lower()
                v = celdas[1].get_text(strip=True)
                if "paterno" in k:   ap_pat   = v
                elif "materno" in k: ap_mat   = v
                elif "nombre" in k:  nombres_r = v

        if ap_pat or nombres_r:
            resultado["nombre"] = f"{ap_pat} {ap_mat} {nombres_r}".strip()

        # Intentar JSON
        if not resultado["nombre"]:
            try:
                j = resp.json()
                d = j.get("data", j)
                np = d.get("apellidoPaterno", d.get("apellido_paterno",""))
                nm = d.get("apellidoMaterno", d.get("apellido_materno",""))
                nn = d.get("nombres","")
                nc = d.get("nombre_completo", d.get("nombreCompleto",""))
                resultado["nombre"] = nc or f"{np} {nm} {nn}".strip()
            except Exception:
                pass

        if not resultado["nombre"]:
            resultado["error"] = "Sin resultado"
    except Exception as e:
        resultado["error"] = str(e)[:60]
    return resultado

# ─── Validar un registro ──────────────────────────────────────────────────────

def validar_registro(nombre_archivo, dni, session):
    # 1. Buscar por DNI en eldni.com
    r1 = consultar_eldni(dni, session)
    if r1["nombre"]:
        sim = similitud(nombre_archivo, r1["nombre"])
        return {
            "nombre_web": r1["nombre"],
            "fuente": "eldni.com",
            "similitud": round(sim * 100),
            "validacion": "CORRECTO" if sim >= 0.80 else "DIFERENTE",
            "error": "",
        }

    time.sleep(random.uniform(1.0, 2.0))

    # 2. Buscar por nombre en buscardniperu.com
    r2 = consultar_buscardni_nombre(nombre_archivo, session)
    sim = similitud(nombre_archivo, r2["nombre"]) if r2["nombre"] else 0
    return {
        "nombre_web": r2["nombre"],
        "fuente": "buscardniperu.com" if r2["nombre"] else "—",
        "similitud": round(sim * 100),
        "validacion": (
            "CORRECTO"      if sim >= 0.80 else
            "DIFERENTE"     if r2["nombre"] else
            "NO ENCONTRADO"
        ),
        "error": " | ".join(filter(None, [r1["error"], r2["error"]])),
    }

# ─── Leer Excel subido ────────────────────────────────────────────────────────

def leer_excel(archivo_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(archivo_bytes), read_only=True)
    ws = wb.active
    registros = []
    dnis_vistos = set()
    for row in ws.iter_rows(min_row=3, values_only=True):
        nombre = str(row[0]).strip() if row[0] else ""
        dni_raw = str(row[1]).strip() if row[1] else ""
        dni = re.sub(r"\D", "", dni_raw).zfill(8)
        if not dni or not dni.replace("0", ""):
            continue
        if dni in dnis_vistos:
            continue
        dnis_vistos.add(dni)
        registros.append({"nombre": nombre, "dni": dni})
    return registros

# ─── Generar Excel de resultado ───────────────────────────────────────────────

def generar_excel(resultados):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validacion DNI"

    AZUL   = "1F4E79"
    thin   = Side(style="thin", color="BFBFBF")
    borde  = Border(left=thin, right=thin, top=thin, bottom=thin)

    fill_cab   = PatternFill("solid", fgColor=AZUL)
    fill_verde = PatternFill("solid", fgColor="E2EFDA")
    fill_amar  = PatternFill("solid", fgColor="FFF2CC")
    fill_rojo  = PatternFill("solid", fgColor="FCE4D6")
    fill_par   = PatternFill("solid", fgColor="F2F2F2")
    fill_impar = PatternFill("solid", fgColor="FFFFFF")

    font_cab   = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    font_datos = Font(name="Arial", size=10)
    align_c    = Alignment(horizontal="center", vertical="center")
    align_l    = Alignment(horizontal="left",   vertical="center")

    cabeceras = ["N°", "DNI", "NOMBRE EN ARCHIVO", "NOMBRE CORRECTO (WEB)", "VALIDACIÓN", "SIMILITUD %", "FUENTE", "OBSERVACION"]
    ws.row_dimensions[1].height = 28
    for col, t in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=col, value=t)
        c.font = font_cab
        c.fill = fill_cab
        c.border = borde
        c.alignment = align_c

    for i, r in enumerate(resultados, start=2):
        ws.row_dimensions[i].height = 17
        val = r.get("validacion", "NO ENCONTRADO")
        fill_fila = fill_par if i % 2 == 0 else fill_impar
        fill_val  = fill_verde if val == "CORRECTO" else (fill_amar if val == "DIFERENTE" else fill_rojo)

        vals = [i-1, r["dni"], r["nombre_archivo"], r.get("nombre_web",""),
                val, r.get("similitud", 0), r.get("fuente",""), r.get("error","")]

        for col, val_celda in enumerate(vals, 1):
            c = ws.cell(row=i, column=col, value=val_celda)
            c.font   = font_datos
            c.border = borde
            c.alignment = align_c if col in (1, 2, 5, 6, 7) else align_l
            c.fill = fill_val if col == 5 else fill_fila

    # Resumen
    conteo = {}
    for r in resultados:
        v = r.get("validacion","NO ENCONTRADO")
        conteo[v] = conteo.get(v, 0) + 1

    fr = len(resultados) + 3
    ws.cell(row=fr,   column=1, value="RESUMEN").font = Font(name="Arial", bold=True, size=11)
    ws.cell(row=fr+1, column=1, value=f"✓ CORRECTO:       {conteo.get('CORRECTO',0)}")
    ws.cell(row=fr+2, column=1, value=f"≠ DIFERENTE:      {conteo.get('DIFERENTE',0)}")
    ws.cell(row=fr+3, column=1, value=f"✗ NO ENCONTRADO:  {conteo.get('NO ENCONTRADO',0)}")
    ws.cell(row=fr+4, column=1, value=f"   TOTAL:         {len(resultados)}")

    for col, ancho in enumerate([5, 12, 40, 40, 16, 13, 22, 40], 1):
        ws.column_dimensions[get_column_letter(col)].width = ancho

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── INTERFAZ STREAMLIT ───────────────────────────────────────────────────────

archivo = st.file_uploader(
    "📂 Sube tu Excel (.xlsx) con columnas: NOMBRE COMPLETO | DNI",
    type=["xlsx"]
)

if archivo:
    datos = archivo.read()
    registros = leer_excel(datos)

    st.success(f"✅ Se encontraron **{len(registros)} DNIs únicos** para validar.")

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total DNIs", len(registros))
    with col2:
        tiempo_est = len(registros) * 4
        minutos = tiempo_est // 60
        segundos = tiempo_est % 60
        st.metric("Tiempo estimado", f"~{minutos}m {segundos}s")

    st.info("💡 El proceso consulta eldni.com por DNI y buscardniperu.com por nombre como respaldo.")

    if st.button("🚀 Iniciar Validación", type="primary", use_container_width=True):

        resultados = []
        session = requests.Session()

        barra     = st.progress(0, text="Iniciando...")
        estado    = st.empty()
        col_a, col_b, col_c = st.columns(3)
        cnt_ok    = col_a.empty()
        cnt_dif   = col_b.empty()
        cnt_no    = col_c.empty()

        ok, dif, no_enc = 0, 0, 0

        for i, reg in enumerate(registros):
            dni    = reg["dni"]
            nombre = reg["nombre"]

            estado.markdown(f"🔍 Consultando **{i+1}/{len(registros)}** — DNI: `{dni}` | {nombre[:40]}")

            r = validar_registro(nombre, dni, session)
            r["dni"]             = dni
            r["nombre_archivo"]  = nombre
            resultados.append(r)

            v = r.get("validacion","")
            if v == "CORRECTO":      ok    += 1
            elif v == "DIFERENTE":   dif   += 1
            else:                    no_enc += 1

            cnt_ok.metric("✓ Correctos",      ok)
            cnt_dif.metric("≠ Diferentes",    dif)
            cnt_no.metric("✗ No encontrados", no_enc)

            barra.progress((i + 1) / len(registros),
                           text=f"Procesando {i+1} de {len(registros)}...")

            if i < len(registros) - 1:
                time.sleep(random.uniform(2.5, 4.0))

        barra.progress(1.0, text="¡Completado!")
        estado.success("✅ Validación finalizada")

        excel_buf = generar_excel(resultados)

        st.download_button(
            label="📥 Descargar Excel con resultados",
            data=excel_buf,
            file_name="resultados_validacion_dni.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        # Tabla resumen en pantalla
        st.markdown("---")
        st.subheader("📊 Resumen")
        col1, col2, col3 = st.columns(3)
        col1.metric("✓ Correctos",      ok,    delta=f"{round(ok/len(registros)*100)}%")
        col2.metric("≠ Diferentes",     dif,   delta=f"{round(dif/len(registros)*100)}%",  delta_color="inverse")
        col3.metric("✗ No encontrados", no_enc, delta=f"{round(no_enc/len(registros)*100)}%", delta_color="inverse")

        # Vista previa
        st.markdown("---")
        st.subheader("👁️ Vista previa de resultados")
        for r in resultados[:10]:
            icono = "✅" if r["validacion"] == "CORRECTO" else ("⚠️" if r["validacion"] == "DIFERENTE" else "❌")
            st.markdown(
                f"{icono} **{r['dni']}** | Archivo: `{r['nombre_archivo']}` "
                f"→ Web: `{r.get('nombre_web','—')}` ({r.get('similitud',0)}%)"
            )
        if len(resultados) > 10:
            st.caption(f"... y {len(resultados)-10} registros más en el Excel.")


