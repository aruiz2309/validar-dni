import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import unicodedata
import time
import random
import io

st.set_page_config(
    page_title="Validador DNI - Perú",
    page_icon="🪪",
    layout="centered"
)

st.title("🪪 Validador de Nombres por DNI")
st.caption("Consulta eldni.com y dniperu.com para verificar si el nombre coincide con el DNI")

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

# ─── Scraping con Playwright ──────────────────────────────────────────────────

def consultar_eldni(page, dni):
    """Busca por número de DNI en eldni.com"""
    try:
        page.goto("https://eldni.com/pe/buscar-datos-por-dni", timeout=20000)
        page.wait_for_load_state("networkidle", timeout=15000)

        # Llenar campo DNI
        campo = (
            page.query_selector("input[name='dni']") or
            page.query_selector("input[id='dni']") or
            page.query_selector("input[type='text']")
        )
        if not campo:
            return ""

        campo.fill("")
        campo.type(dni, delay=100)

        boton = (
            page.query_selector("button[type='submit']") or
            page.query_selector("input[type='submit']") or
            page.query_selector("button:has-text('Buscar')") or
            page.query_selector("button:has-text('Consultar')")
        )
        if boton:
            boton.click()
        else:
            campo.press("Enter")

        page.wait_for_load_state("networkidle", timeout=15000)
        time.sleep(2)

        # Extraer nombre de la tabla
        ap_pat, ap_mat, nombres = "", "", ""
        for fila in page.query_selector_all("table tr"):
            celdas = fila.query_selector_all("td, th")
            if len(celdas) >= 2:
                k = celdas[0].inner_text().strip().lower()
                v = celdas[1].inner_text().strip()
                if "paterno" in k:   ap_pat  = v
                elif "materno" in k: ap_mat  = v
                elif "nombre" in k:  nombres = v

        if ap_pat or nombres:
            return f"{ap_pat} {ap_mat} {nombres}".strip()

        # Fallback: buscar en texto de la pagina
        texto = page.inner_text("body")
        m = re.search(r'(?:Apellido Paterno|Apellido Materno|Nombres?)\s*[:\-]\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ ]{2,})', texto)
        if m:
            return m.group(1).strip()

    except Exception:
        pass
    return ""


def consultar_dniperu(page, dni):
    """Busca por número de DNI en dniperu.com"""
    try:
        page.goto("https://dniperu.com/buscar-dni-nombres-apellidos/", timeout=20000)
        page.wait_for_load_state("networkidle", timeout=15000)

        campo = (
            page.query_selector("input[name='dni']") or
            page.query_selector("input[id='dni']") or
            page.query_selector("input[type='text']") or
            page.query_selector("input[placeholder*='DNI']") or
            page.query_selector("input[placeholder*='dni']")
        )
        if not campo:
            return ""

        campo.fill("")
        campo.type(dni, delay=100)

        boton = (
            page.query_selector("button[type='submit']") or
            page.query_selector("input[type='submit']") or
            page.query_selector("button:has-text('Buscar')") or
            page.query_selector("button:has-text('Consultar')")
        )
        if boton:
            boton.click()
        else:
            campo.press("Enter")

        page.wait_for_load_state("networkidle", timeout=15000)
        time.sleep(2)

        # Extraer nombre
        ap_pat, ap_mat, nombres = "", "", ""
        for fila in page.query_selector_all("table tr"):
            celdas = fila.query_selector_all("td, th")
            if len(celdas) >= 2:
                k = celdas[0].inner_text().strip().lower()
                v = celdas[1].inner_text().strip()
                if "paterno" in k:   ap_pat  = v
                elif "materno" in k: ap_mat  = v
                elif "nombre" in k:  nombres = v

        if ap_pat or nombres:
            return f"{ap_pat} {ap_mat} {nombres}".strip()

        texto = page.inner_text("body")
        m = re.search(r'([A-ZÁÉÍÓÚÑ]{2,}\s+[A-ZÁÉÍÓÚÑ]{2,}(?:\s+[A-ZÁÉÍÓÚÑ]{2,})+)', texto)
        if m:
            return m.group(1).strip()

    except Exception:
        pass
    return ""


def validar_dni(page, nombre_archivo, dni):
    """
    1. Consulta eldni.com por DNI
    2. Si no encuentra, consulta dniperu.com por DNI
    3. Compara nombre encontrado con nombre del archivo
    """
    nombre_web = consultar_eldni(page, dni)

    if not nombre_web:
        time.sleep(random.uniform(1.5, 2.5))
        nombre_web = consultar_dniperu(page, dni)

    if not nombre_web:
        return {"nombre_web": "", "resultado": "NO ENCONTRADO"}

    sim = similitud(nombre_archivo, nombre_web)
    if sim >= 0.80:
        return {"nombre_web": nombre_web, "resultado": "CORRECTO"}
    else:
        return {"nombre_web": nombre_web, "resultado": "INCORRECTO"}


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

def generar_excel(archivo_original_bytes, resultados_dict):
    """
    Toma el Excel original y agrega 2 columnas al final:
    - NOMBRE ENCONTRADO (WEB)
    - CORRECTO / INCORRECTO / NO ENCONTRADO
    """
    wb = openpyxl.load_workbook(io.BytesIO(archivo_original_bytes))
    ws = wb.active

    # Detectar la ultima columna con datos en fila de cabecera (fila 2)
    ultima_col = ws.max_column + 1

    # Estilos
    thin   = Side(style="thin", color="BFBFBF")
    borde  = Border(left=thin, right=thin, top=thin, bottom=thin)
    AZUL   = "1F4E79"

    fill_cab   = PatternFill("solid", fgColor=AZUL)
    fill_verde = PatternFill("solid", fgColor="E2EFDA")
    fill_rojo  = PatternFill("solid", fgColor="FCE4D6")
    fill_amar  = PatternFill("solid", fgColor="FFF2CC")

    font_cab   = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    font_datos = Font(name="Arial", size=10)
    align_c    = Alignment(horizontal="center", vertical="center")
    align_l    = Alignment(horizontal="left",   vertical="center")

    # Cabeceras de las 2 columnas nuevas (en fila 2, donde estan los titulos)
    col_nombre = ultima_col
    col_result = ultima_col + 1

    cab1 = ws.cell(row=2, column=col_nombre, value="NOMBRE ENCONTRADO (WEB)")
    cab1.font = font_cab
    cab1.fill = fill_cab
    cab1.border = borde
    cab1.alignment = align_c

    cab2 = ws.cell(row=2, column=col_result, value="CORRECTO / INCORRECTO")
    cab2.font = font_cab
    cab2.fill = fill_cab
    cab2.border = borde
    cab2.alignment = align_c

    # Llenar datos desde fila 3
    for fila_num, (dni, datos) in enumerate(resultados_dict.items(), start=3):
        nombre_web = datos.get("nombre_web", "")
        resultado  = datos.get("resultado", "NO ENCONTRADO")

        fill_res = (
            fill_verde if resultado == "CORRECTO" else
            fill_rojo  if resultado == "INCORRECTO" else
            fill_amar
        )

        c1 = ws.cell(row=fila_num, column=col_nombre, value=nombre_web)
        c1.font      = font_datos
        c1.border    = borde
        c1.alignment = align_l

        c2 = ws.cell(row=fila_num, column=col_result, value=resultado)
        c2.font      = font_datos
        c2.border    = borde
        c2.alignment = align_c
        c2.fill      = fill_res

    # Ajustar anchos de columnas nuevas
    ws.column_dimensions[get_column_letter(col_nombre)].width = 40
    ws.column_dimensions[get_column_letter(col_result)].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── INTERFAZ ─────────────────────────────────────────────────────────────────

archivo = st.file_uploader(
    "📂 Sube tu Excel con las columnas: NOMBRE COMPLETO | DNI",
    type=["xlsx"]
)

if archivo:
    datos_bytes = archivo.read()
    registros   = leer_excel(datos_bytes)

    st.success(f"✅ **{len(registros)} DNIs únicos** listos para validar")

    col1, col2 = st.columns(2)
    col1.metric("Total DNIs", len(registros))
    mins = (len(registros) * 5) // 60
    segs = (len(registros) * 5) % 60
    col2.metric("Tiempo estimado", f"~{mins}m {segs}s")

    st.warning("⚠️ No cierres esta pestaña mientras se procesa.")

    if st.button("🚀 Iniciar Validación", type="primary", use_container_width=True):

        try:
            from playwright.sync_api import sync_playwright

            resultados_dict = {}   # dni -> {nombre_web, resultado}
            barra   = st.progress(0, text="Iniciando...")
            estado  = st.empty()

            col_a, col_b, col_c = st.columns(3)
            cnt_ok  = col_a.empty()
            cnt_inc = col_b.empty()
            cnt_no  = col_c.empty()
            ok, inc, no_enc = 0, 0, 0

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                context = browser.new_context(
                    user_agent=(
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/124.0.0.0 Safari/537.36"
                    ),
                    locale="es-PE",
                    viewport={"width": 1280, "height": 800},
                )
                page = context.new_page()

                for i, reg in enumerate(registros):
                    dni    = reg["dni"]
                    nombre = reg["nombre"]

                    estado.markdown(
                        f"🔍 **{i+1}/{len(registros)}** — DNI: `{dni}` | `{nombre[:45]}`"
                    )

                    r = validar_dni(page, nombre, dni)
                    resultados_dict[dni] = r

                    v = r["resultado"]
                    if v == "CORRECTO":      ok     += 1
                    elif v == "INCORRECTO":  inc    += 1
                    else:                    no_enc += 1

                    cnt_ok.metric("✅ Correctos",      ok)
                    cnt_inc.metric("❌ Incorrectos",   inc)
                    cnt_no.metric("⚠️ No encontrados", no_enc)

                    barra.progress(
                        (i + 1) / len(registros),
                        text=f"Procesando {i+1} de {len(registros)}..."
                    )

                    if i < len(registros) - 1:
                        time.sleep(random.uniform(2.5, 4.0))

                browser.close()

            barra.progress(1.0, text="¡Completado!")
            estado.success("✅ Validación finalizada")

            # Generar Excel con las 2 columnas nuevas sobre el original
            excel_buf = generar_excel(datos_bytes, resultados_dict)

            st.download_button(
                label="📥 Descargar Excel con resultados",
                data=excel_buf,
                file_name="validacion_dni_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

            # Resumen final
            st.markdown("---")
            st.subheader("📊 Resumen final")
            c1, c2, c3 = st.columns(3)
            c1.metric("✅ Correctos",      ok,     f"{round(ok/len(registros)*100)}%")
            c2.metric("❌ Incorrectos",    inc,    f"{round(inc/len(registros)*100)}%")
            c3.metric("⚠️ No encontrados", no_enc, f"{round(no_enc/len(registros)*100)}%")

            # Vista previa
            st.markdown("---")
            st.subheader("👁️ Vista previa")
            for dni, r in list(resultados_dict.items())[:10]:
                nombre_arch = next((x["nombre"] for x in registros if x["dni"] == dni), "")
                icono = "✅" if r["resultado"] == "CORRECTO" else ("❌" if r["resultado"] == "INCORRECTO" else "⚠️")
                nombre_web = r.get("nombre_web") or "—"
                st.markdown(
                    f"{icono} `{dni}` | Archivo: **{nombre_arch}** → Web: **{nombre_web}**"
                )
            if len(resultados_dict) > 10:
                st.caption(f"... y {len(resultados_dict)-10} registros más en el Excel descargable.")

        except ImportError:
            st.error("❌ Playwright no está instalado. Agrega 'playwright' al requirements.txt y vuelve a desplegar.")


