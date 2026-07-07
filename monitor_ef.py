#!/usr/bin/env python3
"""
Monitor de publicaciones EF — CARM 2026
Detecta novedades, extrae nombres de citaciones, genera dashboard y manda email.

USO:
  python monitor_ef.py --once       # Un solo barrido
  python monitor_ef.py              # Bucle cada 15 min
  python monitor_ef.py --discover   # Ver especialidades
  python monitor_ef.py --reset      # Empezar de cero

REQUISITOS:
  pip install requests beautifulsoup4 pdfplumber

CONFIGURACIÓN EMAIL (local): edita config_ef.json
CONFIGURACIÓN EMAIL (GitHub Actions): secrets GMAIL_USUARIO, GMAIL_CONTRASENA_APP, EMAIL_DESTINO
"""

import argparse
import hashlib
import io
import json
import os
import re
import smtplib
import sys
import time
from collections import defaultdict
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
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
CONFIG_FILE = BASE_DIR / "config_ef.json"

IS_CI = os.getenv("GITHUB_ACTIONS") == "true"

KW_CITACION     = ["citaci", "llamamiento", "llamamient", "convoc"]
KW_CALIFICACION = ["calificaci", "nota", "puntuaci", "resultado", "definitiv",
                   "provisi", "lista", "relacion"]

MESES = {
    "enero":"01","febrero":"02","marzo":"03","abril":"04","mayo":"05","junio":"06",
    "julio":"07","agosto":"08","septiembre":"09","octubre":"10","noviembre":"11","diciembre":"12"
}


# ─── Config (local o env vars) ────────────────────────────────────────────────
def load_config():
    cfg = {"gmail_usuario": "", "gmail_contrasena_app": "", "email_destino": ""}
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    # Variables de entorno tienen prioridad (GitHub Actions Secrets)
    for key, env in [("gmail_usuario","GMAIL_USUARIO"),
                     ("gmail_contrasena_app","GMAIL_CONTRASENA_APP"),
                     ("email_destino","EMAIL_DESTINO")]:
        if os.getenv(env):
            cfg[key] = os.getenv(env)
    return cfg


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

def es_citacion_aspirantes(text):
    """Solo las citaciones a aspirantes determinados contienen listas de nombres."""
    t = normalizar(text)
    return "ASPIRANTE" in t or "DETERMINAD" in t

def fix_url(href):
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"): return "https://servicios.educarm.es" + href
    if not href.startswith("http"): return "https://servicios.educarm.es/admin/" + href
    return href

def fmt_fecha(iso):
    """'2026-07-15' → '15/07/2026'"""
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m/%Y")
    except Exception:
        return iso


# ─── Email ─────────────────────────────────────────────────────────────────────
def enviar_email(config, new_items):
    usuario  = config.get("gmail_usuario", "")
    password = config.get("gmail_contrasena_app", "")
    destino  = config.get("email_destino", "")

    if not usuario or not password or not destino:
        log("  ℹ️  Email no configurado — se omite notificación.", "WARN")
        return

    n = len(new_items)
    asunto = f"[EF CARM 2026] {n} publicación{'es' if n>1 else ''} nueva{'s' if n>1 else ''}"

    filas = ""
    for it in new_items:
        tipo = "🔔 Citación" if it["tipo"] == "citacion" else "📋 Calificación"
        filas += f"""
        <tr>
          <td style="padding:6px 12px;border-bottom:1px solid #eee">{it['tribunal']}</td>
          <td style="padding:6px 12px;border-bottom:1px solid #eee">{tipo}</td>
          <td style="padding:6px 12px;border-bottom:1px solid #eee">
            <a href="{it['url']}" style="color:#1a73e8">{it['text']}</a>
          </td>
        </tr>"""

    html = f"""<html><body style="font-family:Arial,sans-serif;color:#333">
      <h2 style="color:#1e3a5f">Nuevas publicaciones EF — CARM 2026</h2>
      <p>Detectadas el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}</p>
      <table style="border-collapse:collapse;width:100%;margin-top:12px">
        <thead><tr style="background:#f5f5f5">
          <th style="padding:8px 12px;text-align:left">Tribunal</th>
          <th style="padding:8px 12px;text-align:left">Tipo</th>
          <th style="padding:8px 12px;text-align:left">Documento</th>
        </tr></thead>
        <tbody>{filas}</tbody>
      </table>
      <p style="margin-top:20px;color:#888;font-size:12px">
        Monitor automático EF CARM 2026 · GitHub Actions
      </p>
    </body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"]    = usuario
    msg["To"]      = destino
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as smtp:
            smtp.starttls()
            smtp.login(usuario, password)
            smtp.sendmail(usuario, destino, msg.as_bytes())
        log(f"  📧 Email enviado → {destino}")
    except Exception as e:
        log(f"  ❌ Error enviando email: {e}", "ERROR")


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
    url = (f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
           f"&module=publicacionesTribunales&action={action}")
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
    resp = ajax_post("ajaxOpcionesTribunal",
                     {"codCuerpo": COD_CUERPO, "codEspecialidad": cod_esp})
    return [c for c in resp["codTribunales"].split("#") if c]

def get_publicaciones(cod_esp, cod_trib):
    url = (f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
           f"&module=publicacionesTribunales&action=getPublicaciones"
           f"&anyo={ANYO}&convocatoria={CONV}")
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


# ─── Extracción de PDF ─────────────────────────────────────────────────────────

def extraer_datos_pdf(url):
    """
    Descarga el PDF y devuelve:
      {"personas": [...], "fecha_citacion": "2026-07-15"}
    personas: [{"orden":"1", "nombre":"GARCÍA LÓPEZ, MARÍA", "hora":"09:00"}]
    """
    try:
        import pdfplumber
    except ImportError:
        log("  ⚠️  pdfplumber no instalado: pip install pdfplumber", "WARN")
        return {"personas": [], "fecha_citacion": None}

    try:
        r = SESSION.get(url, timeout=30)
        r.raise_for_status()
        pdf_bytes = io.BytesIO(r.content)
    except Exception as e:
        log(f"  ⚠️  Error descargando PDF: {e}", "WARN")
        return {"personas": [], "fecha_citacion": None}

    try:
        with pdfplumber.open(pdf_bytes) as pdf:
            texto = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        log(f"  ⚠️  Error leyendo PDF: {e}", "WARN")
        return {"personas": [], "fecha_citacion": None}

    return {
        "personas": parse_nombres(texto),
        "fecha_citacion": extraer_fecha_citacion(texto),
    }


def extraer_fecha_citacion(texto):
    """
    Busca la fecha del acto (el día para el que se cita, no la publicación).
    Devuelve 'YYYY-MM-DD' o None.
    """
    # "Fecha: 15 de julio de 2026" / "el día 15 de julio de 2026"
    p = re.compile(
        r'(?:fecha\s*(?:de\s*(?:actuaci[oó]n|examen|la\s*prueba|cita[cs]i[oó]n)?)?'
        r'|d[íi]a\s*:?\s*(?:\w+,?\s*)?)'
        r'\s*:?\s*(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})',
        re.IGNORECASE
    )
    m = p.search(texto[:3000])
    if m:
        dia, mes_str, anyo = m.group(1), m.group(2).lower(), m.group(3)
        mes = MESES.get(mes_str)
        if mes:
            return f"{anyo}-{mes}-{dia.zfill(2)}"

    # "15/07/2026" en las primeras líneas
    p2 = re.compile(r'\b(\d{1,2})/(\d{1,2})/(\d{4})\b')
    for m2 in p2.finditer(texto[:1500]):
        anyo = m2.group(3)
        if anyo == ANYO:
            return f"{anyo}-{m2.group(2).zfill(2)}-{m2.group(1).zfill(2)}"

    return None


# Palabras de función y términos administrativos que NUNCA son parte de un nombre
_NO_NOMBRE = {
    # Artículos
    "LA","EL","LOS","LAS","UN","UNA","UNOS","UNAS","DEL","AL",
    # Preposiciones / conjunciones
    "EN","CON","POR","PARA","SIN","SOBRE","BAJO","ENTRE","HASTA","DESDE",
    "SU","SUS","MI","TU","SE","LE","O","U","E","NI","PERO","SINO","MAS","Y",
    "DE","A",  # cortas — solo se comprueban como palabras completas
    # Términos de documentos administrativos de oposiciones
    "CASO","PROGRAMA","PROGRAMACION","DIDACTICA","FICHA","BAREMACION",
    "DOCUMENTACION","ACREDITATIVA","ACREDITATIVO","INTERVENCION","PROCESO",
    "SELECTIVO","INGRESO","ADQUISICION","NUEVAS","NUEVOS","PLAZAS","PLAZA",
    "CONVOCATORIA","CUERPO","ESPECIALIDAD","MAESTROS","EDUCACION","FISICA",
    "OFICIAL","DEBIDAMENTE","COMPULSADA","INSTANCIA","SOLICITUD","IMPRESO",
    "TITULO","CERTIFICADO","COPIA","ORIGINAL","SELLADA","FIRMADA",
    # Admin / lugar
    "TRIBUNAL","CARM","OPOSICION","PAGINA","FECHA","HORA","LUGAR","DIRECCION",
    "MURCIA","REGION","CONSEJERIA","CULTURA","DEPORTES","SECRETARIA",
    "ASPIRANTE","APELLIDOS","NOMBRE","DNI","NIF","ORDEN","NUMERO","LISTA",
    "RELACION","DEFINITIVA","PROVISIONAL","PUNTUACION","NOTA","ACTUACION",
    "EJERCICIO","PRUEBA","JUNTA","CENTRO","INSTITUTO","COLEGIO","ESCUELA",
    "AVENIDA","CALLE","BAREMO","TOTAL","MEDIA","RESULTADO","TURNO","LIBRE",
}


def _es_nombre_valido(texto):
    """
    Devuelve True solo si el texto parece un nombre de persona.
    Reglas:
      - 2 a 5 palabras
      - Ninguna palabra pertenece a _NO_NOMBRE
      - Si tiene coma: la parte de apellidos ≥4 chars, la de nombre ≥3 chars
      - Sin coma: todas las palabras empiezan en mayúscula/solo letras
    """
    palabras = re.findall(r"[A-ZÁÉÍÓÚÜÑ']+", normalizar(texto))
    if len(palabras) < 2 or len(palabras) > 5:
        return False
    for p in palabras:
        if p in _NO_NOMBRE:
            return False
    if "," in texto:
        partes = [p.strip() for p in texto.split(",", 1)]
        if len(partes[0]) < 4 or len(partes[1]) < 3:
            return False
    return True


def parse_nombres(texto):
    """
    Extrae nombres del texto del PDF de una citación.

    Soporta tres formatos habituales en documentos CARM:
      A) "1   GARCÍA LÓPEZ, MARÍA JESÚS   09:00 h"  (número + apellidos, nombre + hora)
      B) "1.- GARCÍA LÓPEZ, MARÍA"                  (número + apellidos, nombre)
      C) "1   PABLO ALCARAZ BAEZA"                  (número + nombre sin coma)

    Rechaza cualquier línea con palabras de función o términos administrativos.
    """
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    personas = []

    # ── Patrón principal: número de orden + nombre (con o sin coma) ────────────
    p_num = re.compile(
        r'^(\d{1,3})(?:[.\-\)]+\s*|\s+)'          # número + separador (. - ) o espacios)
        r'([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\s\-\',]{4,55}?)'  # nombre/apellidos
        r'(?:\s{2,}(\d{1,2}[:.h]\d{2}\s*(?:h(?:oras?)?)?))?'  # hora opcional (≥2 espacios)
        r'\s*$',
        re.IGNORECASE
    )

    for linea in lineas:
        m = p_num.match(linea)
        if not m:
            continue
        candidato = _limpiar(m.group(2))
        if _es_nombre_valido(candidato):
            personas.append({
                "orden":  m.group(1),
                "nombre": candidato,
                "hora":   _hora(m.group(3)),
            })

    if personas:
        return personas

    # ── Fallback: nombres sin número de orden ──────────────────────────────────
    # Solo acepta líneas de letras puras (sin dígitos ni caracteres raros)
    p_sinnum = re.compile(r"^[A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\s\-\',]{4,55}$", re.IGNORECASE)
    orden = 1
    for linea in lineas:
        if re.search(r'\d', linea):        # descarta líneas con números
            continue
        if not p_sinnum.match(linea):
            continue
        candidato = _limpiar(linea)
        if _es_nombre_valido(candidato):
            personas.append({"orden": str(orden), "nombre": candidato, "hora": ""})
            orden += 1
    return personas


def _limpiar(s):
    return re.sub(r'\s+', ' ', s).strip(".,- ")

def _hora(s):
    if not s:
        return ""
    m = re.search(r'(\d{1,2})[:.h](\d{2})', s.strip())
    return f"{m.group(1).zfill(2)}:{m.group(2)}" if m else ""


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
                        "especialidad":   esp["name"],
                        "trib_code":      trib,
                        "tribunal":       f"T{get_tribunal_num(trib)}",
                        "url":            pub["url"],
                        "text":           pub["text"],
                        "tipo":           classify_doc(pub["text"]),
                        "timestamp":      datetime.now().isoformat(),
                        "personas":       [],
                        "fecha_citacion": None,
                    }
    except Exception as e:
        log(f"❌ Error en barrido: {e}", "ERROR")
        import traceback; traceback.print_exc()
    log(f"  Total encontradas: {len(found)}")
    return found


# ─── Proceso ───────────────────────────────────────────────────────────────────
def process(found, state, config):
    known     = state.get("known", {})
    new_items = []

    for key, pub in found.items():
        if key not in known:
            new_items.append(pub)

        existing = known.get(key, {})

        # Extraer nombres solo de "citación a determinados aspirantes"
        if pub["tipo"] == "citacion" and es_citacion_aspirantes(pub["text"]) and not existing.get("personas"):
            log(f"  📄 Extrayendo {pub['tribunal']}: {pub['text'][:50]}")
            datos = extraer_datos_pdf(pub["url"])
            pub["personas"]       = datos["personas"]
            pub["fecha_citacion"] = datos["fecha_citacion"]
            if datos["personas"]:
                log(f"     → {len(datos['personas'])} personas · fecha {datos['fecha_citacion']}")
            else:
                log("     → Sin nombres extraídos")
        else:
            pub["personas"]       = existing.get("personas", [])
            pub["fecha_citacion"] = existing.get("fecha_citacion")

        known[key] = pub

    if new_items:
        log(f"\n🆕 {len(new_items)} novedad(es):")
        for it in new_items:
            log(f"  [{it['tipo'].upper()}] {it['tribunal']}: {it['text']}")
        enviar_email(config, new_items)
    else:
        log("  ✅ Sin novedades")

    state["known"]     = known
    state["last_scan"] = datetime.now().isoformat()
    state["scans"]     = state.get("scans", 0) + 1
    save_state(state)
    generate_dashboard(known)
    return len(new_items)


# ─── Dashboard ─────────────────────────────────────────────────────────────────
def generate_dashboard(known):
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    pubs = sorted(known.values(), key=lambda x: x.get("timestamp",""), reverse=True)

    tribs = sorted({p["tribunal"] for p in pubs},
                   key=lambda x: int(re.sub(r'\D','',x) or 0))
    n_cit = sum(1 for p in pubs if p["tipo"] == "citacion")
    n_cal = len(pubs) - n_cit

    # ── Citaciones: agrupar por DÍA DE ACTUACIÓN → tribunal ──────────────────
    # por_dia_trib[fecha][trib] = {"docs": [...], "personas": [...]}
    por_dia_trib = defaultdict(lambda: defaultdict(lambda: {"docs": [], "personas": []}))
    for p in pubs:
        if p["tipo"] == "citacion":
            # Usar siempre la fecha del acto extraída del PDF;
            # si no hay (ej. apertura de cabeceras sin fecha), agrupar aparte
            fecha = p.get("fecha_citacion") or "0000-sin-fecha"
            trib  = p["tribunal"]
            por_dia_trib[fecha][trib]["docs"].append(p)
            por_dia_trib[fecha][trib]["personas"].extend(p.get("personas", []))

    bloques_dias = ""
    for fecha in sorted(por_dia_trib.keys()):
        if fecha == "0000-sin-fecha":
            fecha_fmt  = "Fecha pendiente"
            dia_nombre = ""
        else:
            fecha_fmt  = fmt_fecha(fecha)
            dia_nombre = _nombre_dia(fecha)
        por_trib   = por_dia_trib[fecha]
        total_dia  = sum(len(v["personas"]) for v in por_trib.values())

        tablas_trib = ""
        for trib in sorted(por_trib.keys(), key=lambda x: int(re.sub(r'\D','',x) or 0)):
            entry    = por_trib[trib]
            docs     = entry["docs"]
            personas = entry["personas"]

            # Enlace(s) al documento — más reciente primero
            doc_links = "".join(
                f'<a href="{d["url"]}" target="_blank" class="doc-link">📄 {d["text"][:55]}</a>'
                for d in sorted(docs, key=lambda x: x.get("timestamp",""), reverse=True)
            )

            # Tabla de personas
            hay_horas = any(pe.get("hora") for pe in personas)
            col_hora  = "<th>Hora</th>" if hay_horas else ""
            filas     = "".join(
                f"<tr><td class='ord'>{pe.get('orden','')}</td>"
                f"<td>{pe['nombre']}</td>"
                f"{'<td class=hora>' + pe.get('hora','') + '</td>' if hay_horas else ''}</tr>"
                for pe in personas
            )
            tabla = (f"<table class='pers-table'>"
                     f"<thead><tr><th style='width:34px'>Nº</th><th>Nombre</th>{col_hora}</tr></thead>"
                     f"<tbody>{filas}</tbody></table>") if personas else \
                    "<p class='empty' style='font-size:12px'>Extrayendo nombres…</p>"

            cnt = f"{len(personas)} pers." if personas else "sin nombres"
            tablas_trib += f"""
            <div class="trib-block">
              <div class="trib-tag">{trib} <span class="cnt">{cnt}</span></div>
              <div class="doc-links">{doc_links}</div>
              {tabla}
            </div>"""

        bloques_dias += f"""
        <div class="dia-block">
          <div class="dia-header">
            <span class="dia-fecha">{dia_nombre} {fecha_fmt}</span>
            <span class="dia-total">{total_dia} personas · {len(por_trib)} tribunal(es)</span>
          </div>
          <div class="trib-grid">{tablas_trib}</div>
        </div>"""

    if not bloques_dias:
        bloques_dias = '<p class="empty">Sin citaciones publicadas todavía.</p>'

    # ── Calificaciones ────────────────────────────────────────────────────────
    cals = [p for p in pubs if p["tipo"] == "calificacion"]
    cal_por_trib = defaultdict(list)
    for p in cals:
        cal_por_trib[p["tribunal"]].append(p)

    cal_rows = ""
    for t in sorted(cal_por_trib.keys(), key=lambda x: int(re.sub(r'\D','',x) or 0)):
        for d in cal_por_trib[t]:
            pub_date = d.get("timestamp","")[:10]
            cal_rows += (f"<tr><td><strong>{t}</strong></td>"
                         f"<td><a href='{d['url']}' target='_blank' class='cal-link'>"
                         f"📋 {d['text'][:65]}</a></td>"
                         f"<td style='white-space:nowrap;color:var(--muted);font-size:12px'>"
                         f"{fmt_fecha(pub_date)}</td></tr>")

    seccion_cal = ""
    if cals:
        seccion_cal = f"""
  <div class="card">
    <h2>📋 Calificaciones publicadas</h2>
    <table>
      <thead><tr>
        <th style="width:80px">Tribunal</th>
        <th>Documento</th>
        <th style="width:90px">Publicado</th>
      </tr></thead>
      <tbody>{cal_rows}</tbody>
    </table>
  </div>"""

    # ── Resumen por tribunal ──────────────────────────────────────────────────
    trib_rows = ""
    for t in tribs:
        tp = [p for p in pubs if p["tribunal"] == t]
        nc = sum(1 for p in tp if p["tipo"] == "citacion")
        nk = len(tp) - nc
        trib_rows += (f"<tr><td><strong>{t}</strong></td>"
                      f"<td class=\"{'ok' if nc else 'empty'}\">{nc or '—'}</td>"
                      f"<td class=\"{'ok' if nk else 'empty'}\">{nk or '—'}</td></tr>")

    last_scan = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="900">
<title>EF CARM 2026</title>
<style>
:root {{
  --bg:#f0f2f5; --card:#fff; --border:#e5e7eb; --primary:#1e3a5f;
  --accent:#2563eb; --text:#333; --muted:#9ca3af; --ok:#166534;
}}
* {{ box-sizing:border-box; margin:0; padding:0 }}
body {{ font-family:'Segoe UI',Arial,sans-serif; background:var(--bg); color:var(--text); font-size:14px }}
a {{ color:var(--accent); text-decoration:none }}
a:hover {{ text-decoration:underline }}

header {{ background:linear-gradient(135deg,#1e3a5f,#2563eb); color:#fff; padding:18px 24px }}
header h1 {{ font-size:19px; font-weight:700 }}
header p  {{ font-size:12px; opacity:.75; margin-top:3px }}

.wrap {{ padding:16px; max-width:1200px; margin:0 auto }}

.stats {{ display:flex; gap:12px; flex-wrap:wrap; margin:16px 0 }}
.stat {{ background:var(--card); border-radius:10px; padding:14px 20px;
         flex:1; min-width:110px; box-shadow:0 1px 4px rgba(0,0,0,.07); text-align:center }}
.stat .num {{ font-size:26px; font-weight:700; color:var(--accent) }}
.stat .lbl {{ font-size:11px; color:var(--muted); margin-top:2px }}

.card {{ background:var(--card); border-radius:10px; padding:16px;
         margin-bottom:14px; box-shadow:0 1px 4px rgba(0,0,0,.07) }}
.card h2 {{ font-size:14px; font-weight:700; color:var(--primary);
            margin-bottom:12px; padding-bottom:8px; border-bottom:2px solid var(--border) }}

table {{ width:100%; border-collapse:collapse; font-size:13px }}
th {{ background:#f8fafc; text-align:left; padding:8px 10px; font-weight:600;
      color:#555; border-bottom:2px solid var(--border) }}
td {{ padding:7px 10px; border-bottom:1px solid #f3f4f6; vertical-align:middle }}
tr:last-child td {{ border-bottom:none }}
tr:hover td {{ background:#fafbff }}
td.ok    {{ color:var(--ok); font-weight:600 }}
td.empty {{ color:var(--muted) }}

/* Bloques por día */
.dia-block {{ margin-bottom:20px }}
.dia-header {{
  display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap;
  background:#1e3a5f; color:#fff; border-radius:8px 8px 0 0; padding:10px 16px;
}}
.dia-fecha {{ font-size:14px; font-weight:700 }}
.dia-total {{ font-size:12px; opacity:.75 }}
.trib-grid {{
  display:grid; grid-template-columns:repeat(auto-fill,minmax(280px,1fr)); gap:1px;
  background:var(--border); border:1px solid var(--border); border-top:none;
  border-radius:0 0 8px 8px; overflow:hidden;
}}
.trib-block {{ background:var(--card); padding:12px }}
.trib-tag {{ font-size:12px; font-weight:700; color:var(--accent);
             margin-bottom:6px; display:flex; align-items:center; gap:6px }}
.trib-tag .cnt {{ font-weight:400; color:var(--muted) }}

/* Enlace(s) de documento dentro del bloque de tribunal */
.doc-links {{ margin-bottom:8px }}
.doc-link {{
  display:block; font-size:12px; padding:4px 8px; margin-bottom:4px;
  background:#eff6ff; border:1px solid #bfdbfe; border-radius:6px;
  color:#1d4ed8; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
}}
.doc-link:hover {{ background:#dbeafe }}
.cal-link {{ color:#166534 }}

.pers-table {{ font-size:12px; width:100% }}
.pers-table th {{ font-size:11px; padding:5px 8px; background:#f8fafc }}
.pers-table td {{ padding:5px 8px; border-bottom:1px solid #f3f4f6 }}
.pers-table tr:last-child td {{ border-bottom:none }}
.pers-table tr:hover td {{ background:#f0f7ff }}
td.ord  {{ color:var(--muted); font-size:11px; width:28px }}
td.hora {{ color:#374151; white-space:nowrap; font-size:12px }}

.empty {{ color:var(--muted); font-style:italic; padding:8px 0; font-size:13px }}
.ts {{ font-size:11px; color:var(--muted); text-align:right; margin-top:8px }}
</style>
</head>
<body>
<header>
  <h1>📋 Oposiciones EF — CARM 2026</h1>
  <p>Publicaciones de tribunales · Educación Física · Actualización automática</p>
</header>
<div class="wrap">

  <div class="stats">
    <div class="stat"><div class="num">{len(tribs)}</div><div class="lbl">Tribunales</div></div>
    <div class="stat"><div class="num" style="color:#1d4ed8">{n_cit}</div><div class="lbl">Citaciones</div></div>
    <div class="stat"><div class="num" style="color:var(--ok)">{n_cal}</div><div class="lbl">Calificaciones</div></div>
    <div class="stat"><div class="num" style="color:#6b7280">{len(pubs)}</div><div class="lbl">Total docs</div></div>
  </div>

  <div class="card">
    <h2>🔔 Citaciones — personas por día de actuación</h2>
    {bloques_dias}
  </div>
{seccion_cal}
  <div class="card">
    <h2>📊 Estado por tribunal</h2>
    <table>
      <thead><tr><th>Tribunal</th><th>Citaciones</th><th>Calificaciones</th></tr></thead>
      <tbody>{trib_rows}</tbody>
    </table>
  </div>

  <div class="ts">Última actualización: {last_scan} · Se recarga cada 15 min</div>
</div>
</body>
</html>"""

    out = DOCS_DIR / "index.html"
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"  🌐 Dashboard → {out}")


def _nombre_dia(fecha_iso):
    try:
        dias = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
        d = datetime.fromisoformat(fecha_iso)
        return dias[d.weekday()]
    except Exception:
        return ""


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

    config = load_config()

    if args.reset:
        STATE_FILE.unlink(missing_ok=True)
        log("Estado reiniciado.")
        return

    if args.discover:
        discover()
        return

    estado_email = "configurado" if config.get("gmail_usuario") else "no configurado"
    state = load_state()
    log(f"Publicaciones conocidas: {len(state.get('known',{}))}")
    log(f"Email: {estado_email}")

    count = 0
    while True:
        count += 1
        log(f"\n── Barrido #{count} ({datetime.now().strftime('%H:%M:%S')}) ──")
        found = scan()
        if found is not None:
            process(found, state, config)
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
