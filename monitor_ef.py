#!/usr/bin/env python3
"""
Monitor de publicaciones EF — CARM 2026
- Detecta nuevos documentos en tribunales
- Extrae nombres de las citaciones en PDF
- Genera dashboard HTML para GitHub Pages

USO:
  python monitor_ef.py --once       # Un solo barrido
  python monitor_ef.py              # Bucle cada 15 min
  python monitor_ef.py --discover   # Ver especialidades disponibles
  python monitor_ef.py --reset      # Empezar de cero

REQUISITOS:
  pip install requests beautifulsoup4 pdfplumber
"""

import argparse
import hashlib
import io
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("Faltan dependencias: pip install requests beautifulsoup4 pdfplumber")
    sys.exit(1)

# ─── Configuración ─────────────────────────────────────────────────────────────
BASE_URL   = "https://servicios.educarm.es/admin/index2.php"
ANYO       = "2026"
CONV       = "OPOPRI26"
COD_CUERPO = "0597"
FILTRO_ESP = ["FISICA"]

BASE_DIR   = Path(__file__).parent
STATE_FILE = BASE_DIR / "state_ef.json"
DOCS_DIR   = BASE_DIR / "docs"
LOG_FILE   = BASE_DIR / "monitor_ef.log"

IS_CI = os.getenv("GITHUB_ACTIONS") == "true"

KW_CITACION     = ["citaci", "llamamiento", "llamamient", "convoc"]
KW_CALIFICACION = ["calificaci", "nota", "puntuaci", "resultado", "definitiv", "provisi", "lista", "relacion"]


# ─── Utilidades ────────────────────────────────────────────────────────────────
def normalizar(texto):
    tr = str.maketrans("áéíóúÁÉÍÓÚàèìòùüÜñÑ", "aeiouAEIOUaeiouuUnN")
    return str(texto).upper().translate(tr)


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {level}: {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_state():
    if STATE_FILE.exists():
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"known": {}, "last_scan": None, "scans": 0}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def make_key(esp, trib, url):
    return hashlib.md5(f"{esp}|{trib}|{url}".encode()).hexdigest()


def get_tribunal_num(cod):
    m = re.search(r'(\d+)', str(cod))
    return m.group(1) if m else str(cod)


def classify_doc(text):
    t = normalizar(text)
    for kw in KW_CITACION:
        if normalizar(kw) in t:
            return "citacion"
    for kw in KW_CALIFICACION:
        if normalizar(kw) in t:
            return "calificacion"
    return "calificacion"


def fix_url(href):
    if href.startswith("//"):   return "https:" + href
    if href.startswith("/"):    return "https://servicios.educarm.es" + href
    if not href.startswith("http"): return "https://servicios.educarm.es/admin/" + href
    return href


# ─── Extracción de nombres desde PDF de citación ───────────────────────────────
def extract_names_from_pdf(url):
    """
    Descarga el PDF de una citación y extrae la lista de personas citadas.
    Devuelve lista de dicts: [{"orden": "1", "nombre": "GARCÍA LÓPEZ, MARÍA", "hora": "09:00"}]
    """
    try:
        import pdfplumber
    except ImportError:
        log("  ⚠️  pdfplumber no instalado: pip install pdfplumber", "WARN")
        return []

    try:
        r = SESSION.get(url, timeout=30)
        r.raise_for_status()
        pdf_bytes = io.BytesIO(r.content)
    except Exception as e:
        log(f"  ⚠️  Error descargando PDF para extracción: {e}", "WARN")
        return []

    try:
        with pdfplumber.open(pdf_bytes) as pdf:
            texto = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        log(f"  ⚠️  Error leyendo PDF: {e}", "WARN")
        return []

    return parse_names_from_text(texto)


def parse_names_from_text(texto):
    """
    Intenta extraer nombres de un texto de citación. Prueba varios patrones
    habituales en documentos de oposiciones de la CARM.
    """
    personas = []
    lineas   = [l.strip() for l in texto.splitlines() if l.strip()]

    # ── Patrón 1: líneas con número de orden + nombre + hora opcional ──
    # Ej: "1   GARCÍA LÓPEZ, MARÍA JESÚS   09:00"
    # Ej: "1.- GARCÍA LÓPEZ, MARÍA"
    p1 = re.compile(
        r'^(\d{1,3})[.\-\s]+([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\s,\-\'\.]{5,60}?)'
        r'(?:\s+(\d{1,2}[:\.\s]\d{2}\s*(?:h|horas?)?))?$',
        re.IGNORECASE
    )
    for linea in lineas:
        m = p1.match(linea)
        if m:
            nombre = limpiar_nombre(m.group(2))
            if nombre and es_nombre_valido(nombre):
                personas.append({
                    "orden":  m.group(1),
                    "nombre": nombre,
                    "hora":   normalizar_hora(m.group(3)) if m.group(3) else "",
                })

    if personas:
        return personas

    # ── Patrón 2: líneas en mayúsculas con coma (APELLIDOS, NOMBRE) ──
    # Ej: "GARCÍA LÓPEZ, MARÍA JESÚS"
    p2 = re.compile(r'^([A-ZÁÉÍÓÚÜÑ]{2,}(?:\s+[A-ZÁÉÍÓÚÜÑ]{2,})+,\s*[A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\s]{2,30})$')
    orden = 1
    for linea in lineas:
        m = p2.match(linea)
        if m:
            nombre = limpiar_nombre(m.group(1))
            if nombre and es_nombre_valido(nombre):
                personas.append({"orden": str(orden), "nombre": nombre, "hora": ""})
                orden += 1

    if personas:
        return personas

    # ── Patrón 3: tabla con columnas detectadas por posición ──
    # Busca bloque de texto que parezca una tabla con nombres
    bloque = []
    en_tabla = False
    for linea in lineas:
        n = normalizar(linea)
        if any(k in n for k in ["APELLIDO", "NOMBRE", "ASPIRANTE", "OPOSITOR", "ORDEN"]):
            en_tabla = True
            continue
        if en_tabla:
            if len(linea) > 5 and re.search(r'[A-ZÁÉÍÓÚÜÑ]{3,}', linea):
                bloque.append(linea)
            elif len(bloque) > 2 and len(linea) < 3:
                break

    orden = 1
    for linea in bloque:
        partes = re.split(r'\s{2,}|\t', linea)
        for parte in partes:
            parte = parte.strip()
            if re.match(r'^[A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\s,\-]{5,}$', parte):
                nombre = limpiar_nombre(parte)
                if nombre and es_nombre_valido(nombre):
                    personas.append({"orden": str(orden), "nombre": nombre, "hora": ""})
                    orden += 1
                    break

    return personas


def limpiar_nombre(nombre):
    nombre = re.sub(r'\s+', ' ', nombre).strip()
    nombre = nombre.strip(".,- ")
    return nombre


PALABRAS_EXCLUIDAS = {
    "TRIBUNAL", "EDUCACION", "FISICA", "CARM", "CONVOCATORIA", "OPOSICION",
    "PAGINA", "FECHA", "HORA", "LUGAR", "DIRECCION", "MURCIA", "REGION",
    "CONSEJERIA", "EDUCACION", "CULTURA", "DEPORTES", "SECRETARIA", "ASPIRANTE",
    "APELLIDOS", "NOMBRE", "DNI", "NIF", "ORDEN", "NUMERO", "LISTA",
}


def es_nombre_valido(nombre):
    n = normalizar(nombre)
    # Mínimo 2 palabras, al menos una con más de 2 letras
    palabras = [p for p in n.split() if p not in {",", "-"}]
    if len(palabras) < 2:
        return False
    # No debe ser una cabecera o palabra de formulario
    for excl in PALABRAS_EXCLUIDAS:
        if excl in n and len(palabras) <= 2:
            return False
    # Al menos una palabra de más de 3 letras en mayúsculas
    if not any(len(p) > 3 for p in palabras):
        return False
    return True


def normalizar_hora(hora_str):
    if not hora_str:
        return ""
    hora_str = hora_str.strip()
    m = re.search(r'(\d{1,2})[:\.\s](\d{2})', hora_str)
    if m:
        return f"{m.group(1).zfill(2)}:{m.group(2)}"
    return hora_str.strip()


# ─── HTTP / CARM ───────────────────────────────────────────────────────────────
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": BASE_URL,
    "Accept-Language": "es-ES,es;q=0.9",
})


def ajax_post(action, extra=None):
    data = {"anyo": ANYO, "convocatoria": CONV}
    if extra:
        data.update(extra)
    url = f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES&module=publicacionesTribunales&action={action}"
    r = SESSION.post(url, data=data, timeout=20)
    r.raise_for_status()
    return r.json()


def get_especialidades(todas=False):
    resp  = ajax_post("ajaxOpcionesEspecialidad", {"codCuerpo": COD_CUERPO})
    codes = resp["codEspecialidades"].split("#")
    names = resp["denEspecialidades"].split("#")
    return [
        {"code": c, "name": n} for c, n in zip(codes, names)
        if c and (todas or any(f in normalizar(n) for f in FILTRO_ESP))
    ]


def get_tribunales(cod_esp):
    resp = ajax_post("ajaxOpcionesTribunal", {"codCuerpo": COD_CUERPO, "codEspecialidad": cod_esp})
    return [c for c in resp["codTribunales"].split("#") if c]


def get_publicaciones(cod_esp, cod_trib):
    url = (f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
           f"&module=publicacionesTribunales&action=getPublicaciones&anyo={ANYO}&convocatoria={CONV}")
    r = SESSION.post(url, data={
        "cuerpos_lista": COD_CUERPO,
        "especialidades_lista": cod_esp,
        "tribunales_lista": cod_trib,
    }, timeout=20)
    r.raise_for_status()
    r.encoding = "iso-8859-1"
    soup = BeautifulSoup(r.text, "html.parser")
    pubs = []
    KW = (".pdf", "download", "fichero", "documento", "getfile", "descarga")
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        text = a.get_text(strip=True)
        if not text or len(text) < 3:
            continue
        if any(k in href.lower() for k in KW):
            pubs.append({"url": fix_url(href), "text": text})
    if not pubs:
        for a in soup.select("table a[href]"):
            href = a["href"].strip()
            text = a.get_text(strip=True)
            if href and text and len(text) > 2 and not href.startswith("#"):
                pubs.append({"url": fix_url(href), "text": text})
    return pubs


# ─── Barrido ───────────────────────────────────────────────────────────────────
def scan():
    log("🚀 Barrido EF CARM...")
    found = {}
    try:
        SESSION.get(
            f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
            f"&module=publicacionesTribunales&anyo={ANYO}&convocatoria={CONV}",
            timeout=15,
        )
        especialidades = get_especialidades()
        if not especialidades:
            log("  ⚠️  Sin especialidades EF. Prueba --discover.", "WARN")
            return found

        log(f"  Especialidades: {[e['name'] for e in especialidades]}")
        for esp in especialidades:
            tribs = get_tribunales(esp["code"])
            log(f"  {esp['name']}: {len(tribs)} tribunal(es)")
            for trib in tribs:
                pubs = get_publicaciones(esp["code"], trib)
                for pub in pubs:
                    key = make_key(esp["name"], trib, pub["url"])
                    found[key] = {
                        "especialidad": esp["name"],
                        "trib_code":    trib,
                        "tribunal":     f"T{get_tribunal_num(trib)}",
                        "url":          pub["url"],
                        "text":         pub["text"],
                        "tipo":         classify_doc(pub["text"]),
                        "timestamp":    datetime.now().isoformat(),
                        "personas":     [],   # se rellena para citaciones
                    }
    except Exception as e:
        log(f"❌ Error en barrido: {e}", "ERROR")
        import traceback; traceback.print_exc()

    log(f"  Total encontradas: {len(found)}")
    return found


# ─── Proceso ───────────────────────────────────────────────────────────────────
def process(found, state):
    known     = state.get("known", {})
    new_items = []

    for key, pub in found.items():
        existing = known.get(key)

        # ¿Es nuevo?
        if not existing:
            new_items.append(pub)

        # Extraer nombres de citaciones (solo si no se ha hecho ya)
        if pub["tipo"] == "citacion":
            personas_existentes = (existing or {}).get("personas", [])
            if not personas_existentes:
                log(f"  📄 Extrayendo nombres de {pub['tribunal']}: {pub['text'][:50]}")
                pub["personas"] = extract_names_from_pdf(pub["url"])
                if pub["personas"]:
                    log(f"     → {len(pub['personas'])} personas encontradas")
                else:
                    log(f"     → No se pudieron extraer nombres del PDF")
            else:
                pub["personas"] = personas_existentes

        known[key] = pub

    if new_items:
        log(f"\n🆕 {len(new_items)} novedad(es):")
        for it in new_items:
            log(f"  [{it['tipo'].upper()}] {it['tribunal']}: {it['text']}")

    state["known"]     = known
    state["last_scan"] = datetime.now().isoformat()
    state["scans"]     = state.get("scans", 0) + 1
    save_state(state)
    generate_dashboard(known)
    return len(new_items)


# ─── Dashboard ─────────────────────────────────────────────────────────────────
def generate_dashboard(known):
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    pubs = sorted(known.values(), key=lambda x: x.get("timestamp", ""), reverse=True)

    # ── Estadísticas ──
    n_cit  = sum(1 for p in pubs if p["tipo"] == "citacion")
    n_cal  = len(pubs) - n_cit
    tribs  = sorted({p["tribunal"] for p in pubs},
                    key=lambda x: int(re.sub(r'\D', '', x) or 0))

    # ── Resumen por tribunal ──
    trib_rows = ""
    for t in tribs:
        tp = [p for p in pubs if p["tribunal"] == t]
        nc = sum(1 for p in tp if p["tipo"] == "citacion")
        nk = sum(1 for p in tp if p["tipo"] == "calificacion")
        trib_rows += f"""
        <tr>
          <td><strong>{t}</strong></td>
          <td class="{'ok' if nc else 'empty'}">{nc or '—'}</td>
          <td class="{'ok' if nk else 'empty'}">{nk or '—'}</td>
        </tr>"""

    # ── Feed de publicaciones ──
    pub_rows = ""
    for p in pubs[:100]:
        badge = ('<span class="badge cit">Citación</span>'
                 if p["tipo"] == "citacion"
                 else '<span class="badge cal">Calificación</span>')
        fecha = p.get("timestamp", "")[:10]
        hora  = p.get("timestamp", "")[11:16]
        pub_rows += f"""
        <tr>
          <td>{p['tribunal']}</td>
          <td>{badge}</td>
          <td><a href="{p['url']}" target="_blank">{p['text'][:70]}</a></td>
          <td style="white-space:nowrap">{fecha}<br><small style="color:#aaa">{hora}</small></td>
        </tr>"""

    # ── Tablas de personas citadas por tribunal y día ──
    # Agrupar: tribunal → fecha → lista de personas
    citaciones_con_personas = [
        p for p in pubs
        if p["tipo"] == "citacion" and p.get("personas")
    ]

    # Organizar por tribunal y luego por fecha del documento
    from collections import defaultdict
    por_tribunal = defaultdict(lambda: defaultdict(list))
    for p in citaciones_con_personas:
        fecha = p.get("timestamp", "")[:10]
        por_tribunal[p["tribunal"]][fecha].append(p)

    bloques_citados = ""
    for trib in sorted(por_tribunal.keys(), key=lambda x: int(re.sub(r'\D', '', x) or 0)):
        por_fecha = por_tribunal[trib]
        for fecha in sorted(por_fecha.keys(), reverse=True):
            docs_dia = por_fecha[fecha]
            # Acumular todas las personas de todos los documentos de ese tribunal+día
            personas_dia = []
            for doc in docs_dia:
                for persona in doc.get("personas", []):
                    personas_dia.append({**persona, "_doc": doc["text"]})

            if not personas_dia:
                continue

            fecha_fmt  = datetime.fromisoformat(fecha).strftime("%d/%m/%Y") if fecha else fecha
            hay_horas  = any(p.get("hora") for p in personas_dia)
            hay_orden  = any(p.get("orden") for p in personas_dia)

            cab_extra  = "<th>Hora</th>" if hay_horas else ""
            filas_pers = ""
            for p in personas_dia:
                col_hora = f"<td style='white-space:nowrap'>{p.get('hora','')}</td>" if hay_horas else ""
                orden    = p.get("orden", "")
                filas_pers += f"""
                <tr>
                  {'<td class="ord">' + orden + '</td>' if hay_orden else ''}
                  <td>{p['nombre']}</td>
                  {col_hora}
                </tr>"""

            col_ord_th = "<th style='width:40px'>Nº</th>" if hay_orden else ""
            bloques_citados += f"""
            <div class="cit-block">
              <div class="cit-header">
                <span class="trib-tag">{trib}</span>
                <span class="fecha-tag">📅 {fecha_fmt}</span>
                <span class="count-tag">{len(personas_dia)} personas</span>
              </div>
              <table class="pers-table">
                <thead><tr>{col_ord_th}<th>Nombre</th>{cab_extra}</tr></thead>
                <tbody>{filas_pers}</tbody>
              </table>
            </div>"""

    sin_personas_msg = ""
    if citaciones_con_personas == [] and any(p["tipo"] == "citacion" for p in pubs):
        sin_personas_msg = '<p class="info-msg">⏳ Extrayendo nombres de las citaciones en el próximo barrido...</p>'

    last_scan = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="900">
<title>EF CARM 2026</title>
<style>
* {{ box-sizing:border-box; margin:0; padding:0 }}
body {{ font-family:'Segoe UI',Arial,sans-serif; background:#f0f2f5; color:#333; font-size:14px }}
header {{ background:linear-gradient(135deg,#1e3a5f,#2563eb); color:#fff; padding:18px 24px }}
header h1 {{ font-size:19px; font-weight:700 }}
header p  {{ font-size:12px; opacity:.75; margin-top:3px }}
.wrap {{ padding:16px; max-width:1100px; margin:0 auto }}

/* Stats */
.stats {{ display:flex; gap:12px; flex-wrap:wrap; margin:16px 0 }}
.stat {{ background:#fff; border-radius:10px; padding:14px 20px; flex:1; min-width:110px;
         box-shadow:0 1px 4px rgba(0,0,0,.08); text-align:center }}
.stat .num {{ font-size:28px; font-weight:700; color:#2563eb }}
.stat .lbl {{ font-size:11px; color:#888; margin-top:2px }}

/* Cards */
.card {{ background:#fff; border-radius:10px; padding:16px; margin-bottom:14px;
         box-shadow:0 1px 4px rgba(0,0,0,.08) }}
.card h2 {{ font-size:14px; font-weight:700; color:#1e3a5f; margin-bottom:12px;
            padding-bottom:8px; border-bottom:2px solid #e5e7eb }}

/* Tablas generales */
table {{ width:100%; border-collapse:collapse; font-size:13px }}
th {{ background:#f8fafc; text-align:left; padding:8px 10px; font-weight:600;
      color:#555; border-bottom:2px solid #e5e7eb }}
td {{ padding:7px 10px; border-bottom:1px solid #f0f0f0; vertical-align:middle }}
tr:last-child td {{ border-bottom:none }}
tr:hover td {{ background:#fafbff }}
td.ok    {{ color:#166534; font-weight:600 }}
td.empty {{ color:#bbb }}

/* Badges */
.badge {{ display:inline-block; border-radius:12px; padding:2px 8px;
          font-size:11px; font-weight:600 }}
.badge.cit {{ background:#dbeafe; color:#1d4ed8 }}
.badge.cal {{ background:#dcfce7; color:#166534 }}

/* Bloques de personas citadas */
.cit-block {{
  border:1px solid #e5e7eb; border-radius:8px;
  margin-bottom:14px; overflow:hidden;
}}
.cit-header {{
  background:#f1f5f9; padding:10px 14px;
  display:flex; align-items:center; gap:10px; flex-wrap:wrap;
}}
.trib-tag {{
  background:#2563eb; color:#fff; border-radius:20px;
  padding:2px 10px; font-size:12px; font-weight:700;
}}
.fecha-tag {{ font-size:13px; font-weight:600; color:#374151 }}
.count-tag {{ font-size:12px; color:#6b7280; margin-left:auto }}
.pers-table {{ width:100%; border-collapse:collapse; font-size:13px }}
.pers-table th {{ background:#f8fafc; padding:7px 12px; text-align:left;
                  font-weight:600; color:#555; border-bottom:1px solid #e5e7eb }}
.pers-table td {{ padding:6px 12px; border-bottom:1px solid #f5f5f5 }}
.pers-table td.ord {{ color:#9ca3af; font-size:12px; width:36px }}
.pers-table tr:last-child td {{ border-bottom:none }}
.pers-table tr:hover td {{ background:#fafbff }}
.pers-table tr:nth-child(even) td {{ background:#fafeff }}

.info-msg {{ color:#6b7280; font-style:italic; padding:8px 0 }}
a {{ color:#2563eb; text-decoration:none }}
a:hover {{ text-decoration:underline }}
.ts {{ font-size:11px; color:#aaa; text-align:right; margin-top:8px }}
</style>
</head>
<body>
<header>
  <h1>📋 Oposiciones EF — CARM 2026</h1>
  <p>Publicaciones de tribunales · Educación Física · Se recarga cada 15 min</p>
</header>
<div class="wrap">

  <div class="stats">
    <div class="stat"><div class="num">{len(tribs)}</div><div class="lbl">Tribunales</div></div>
    <div class="stat"><div class="num" style="color:#1d4ed8">{n_cit}</div><div class="lbl">Citaciones</div></div>
    <div class="stat"><div class="num" style="color:#166534">{n_cal}</div><div class="lbl">Calificaciones</div></div>
    <div class="stat"><div class="num" style="color:#6b7280">{len(pubs)}</div><div class="lbl">Total docs</div></div>
  </div>

  <div class="card">
    <h2>📊 Estado por tribunal</h2>
    <table>
      <thead><tr><th>Tribunal</th><th>Citaciones</th><th>Calificaciones</th></tr></thead>
      <tbody>{trib_rows}</tbody>
    </table>
  </div>

  <div class="card">
    <h2>👥 Personas citadas por tribunal y día</h2>
    {sin_personas_msg}
    {bloques_citados if bloques_citados else '<p class="info-msg">Sin citaciones publicadas aún.</p>'}
  </div>

  <div class="card">
    <h2>🕐 Todas las publicaciones</h2>
    <table>
      <thead><tr>
        <th style="width:70px">Tribunal</th>
        <th style="width:110px">Tipo</th>
        <th>Documento</th>
        <th style="width:90px">Fecha</th>
      </tr></thead>
      <tbody>{pub_rows}</tbody>
    </table>
  </div>

  <div class="ts">Última actualización: {last_scan}</div>
</div>
</body>
</html>"""

    out = DOCS_DIR / "index.html"
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"  🌐 Dashboard generado → {out}")


# ─── Descubrimiento ────────────────────────────────────────────────────────────
def discover():
    log("Consultando especialidades...")
    SESSION.get(
        f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
        f"&module=publicacionesTribunales&anyo={ANYO}&convocatoria={CONV}", timeout=15
    )
    for e in get_especialidades(todas=True):
        marca = " ← MONITORIZADA" if any(f in normalizar(e["name"]) for f in FILTRO_ESP) else ""
        log(f"  [{e['code']}] {e['name']}{marca}")


# ─── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--once",     action="store_true")
    parser.add_argument("--discover", action="store_true")
    parser.add_argument("--reset",    action="store_true")
    parser.add_argument("--interval", type=int, default=15)
    args = parser.parse_args()

    if IS_CI:
        log("🤖 GitHub Actions.")

    if args.reset:
        STATE_FILE.unlink(missing_ok=True)
        log("Estado reiniciado.")
        return

    if args.discover:
        discover()
        return

    state = load_state()
    log(f"Publicaciones conocidas: {len(state.get('known', {}))}")

    count = 0
    while True:
        count += 1
        log(f"\n── Barrido #{count} ({datetime.now().strftime('%H:%M:%S')}) ──")
        found = scan()
        if found is not None:
            process(found, state)
            state = load_state()

        if args.once:
            log("✅ Completado.")
            break

        log(f"  ⏳ Próximo en {args.interval} min...")
        try:
            time.sleep(args.interval * 60)
        except KeyboardInterrupt:
            log("👋 Detenido.")
            break


if __name__ == "__main__":
    main()
