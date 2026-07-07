#!/usr/bin/env python3
"""
══════════════════════════════════════════════════════════════
  MONITOR DE PUBLICACIONES EF — educarm (CARM) 2026
  Descarga citaciones y calificaciones · Dashboard web
  Compatible con ejecución local Y GitHub Actions
══════════════════════════════════════════════════════════════

USO LOCAL:
  python monitor_ef.py              # Bucle cada 5 min
  python monitor_ef.py --once       # Un solo barrido
  python monitor_ef.py --discover   # Ver especialidades

EN GITHUB ACTIONS:
  Se ejecuta automáticamente vía cron. Variables de entorno:
    GMAIL_USUARIO, GMAIL_CONTRASENA_APP, EMAIL_DESTINO

REQUISITOS:
  pip install requests beautifulsoup4 openpyxl
"""

import argparse
import hashlib
import json
import os
import platform
import re
import smtplib
import sys
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("❌ Faltan dependencias. Ejecuta:")
    print("   pip install requests beautifulsoup4 openpyxl")
    sys.exit(1)


# ─── Configuración fija ───────────────────────────────────────────────────────
BASE_URL   = "https://servicios.educarm.es/admin/index2.php"
ANYO       = "2026"
CONV       = "OPOPRI26"
COD_CUERPO = "0597"

FILTRO_ESP = ["FISICA"]   # Solo Educación Física

BASE_DIR       = Path(__file__).parent
CONFIG_FILE    = BASE_DIR / "config_ef.json"
STATE_FILE     = BASE_DIR / "state_ef.json"
CITACIONES_DIR = BASE_DIR / "citaciones"
CALIFIC_DIR    = BASE_DIR / "calificaciones"
WEB_DIR        = BASE_DIR / "docs"          # GitHub Pages sirve desde /docs
LOG_FILE       = BASE_DIR / "monitor_ef.log"
DEFAULT_INTERVAL = 5

# Detectar entorno CI (GitHub Actions)
IS_CI = os.getenv("GITHUB_ACTIONS") == "true"

# Palabras clave para clasificar documentos
KW_CITACION    = ["citaci", "llamamiento", "cita ", "convocatoria", "llamamient"]
KW_CALIFICACION = ["calificaci", "nota", "puntuaci", "resultado", "definitiv", "provisi"]


# ─── Config ───────────────────────────────────────────────────────────────────
def load_config():
    defaults = {
        "gmail_usuario": "",
        "gmail_contrasena_app": "",
        "email_destino": "",
        "excel_alumnos": str(BASE_DIR / "Consulta" / "Cruce_Tribunales_Academia_con_Baremo.xlsx"),
    }
    # Leer config.json si existe
    cfg = dict(defaults)
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    # Las variables de entorno tienen prioridad (GitHub Actions Secrets)
    env_map = {
        "gmail_usuario":        "GMAIL_USUARIO",
        "gmail_contrasena_app": "GMAIL_CONTRASENA_APP",
        "email_destino":        "EMAIL_DESTINO",
    }
    for key, env_var in env_map.items():
        val = os.getenv(env_var)
        if val:
            cfg[key] = val
    if IS_CI:
        log("  🤖 Modo GitHub Actions detectado.")
    return cfg


# ─── Sesión HTTP ──────────────────────────────────────────────────────────────
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": BASE_URL,
    "Accept-Language": "es-ES,es;q=0.9",
})


# ─── Utilidades ───────────────────────────────────────────────────────────────
def normalizar(texto):
    reemplazos = str.maketrans("áéíóúÁÉÍÓÚàèìòùüÜñÑ", "aeiouAEIOUaeiouuUnN")
    return texto.upper().translate(reemplazos)


def log(msg, level="INFO"):
    ts   = datetime.now().strftime("%H:%M:%S")
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


def sanitize_filename(name):
    name = re.sub(r'[/\\:*?"<>|]', '_', name)
    name = re.sub(r'\s+', '_', name)
    return name.strip("_. ")[:80]


def make_key(especialidad, tribunal, url):
    return hashlib.md5(f"{especialidad}|{tribunal}|{url}".encode()).hexdigest()


def get_tribunal_num(cod_tribunal):
    """Extrae número limpio del código de tribunal → '1', '2', etc."""
    m = re.search(r'(\d+)', str(cod_tribunal))
    return m.group(1) if m else cod_tribunal


def beep(n=3):
    if IS_CI:
        return  # Sin sonido en GitHub Actions
    sistema = platform.system()
    for _ in range(n):
        if sistema == "Windows":
            try:
                import winsound
                winsound.Beep(1000, 500)
            except Exception:
                pass
        time.sleep(0.3)


# ─── Clasificación de documentos ──────────────────────────────────────────────
def classify_doc(text):
    """Devuelve 'citacion' o 'calificacion' según el texto del enlace."""
    t = normalizar(text)
    for kw in KW_CITACION:
        if normalizar(kw) in t:
            return "citacion"
    for kw in KW_CALIFICACION:
        if normalizar(kw) in t:
            return "calificacion"
    # Por defecto: si no hay pista clara, miramos si tiene "lista"
    if "LISTA" in t or "RELACION" in t:
        return "calificacion"
    return "otro"


# ─── Carga de alumnos de la academia desde Excel ──────────────────────────────
def load_academy_students(excel_path):
    """
    Lee el Excel y devuelve un dict:
      {nombre_normalizado: {"tribunal": "1", "baremo": 8.5, "nombre_real": "..."}}
    Detecta automáticamente las columnas relevantes.
    """
    alumnos = {}
    path = Path(excel_path)
    if not path.exists():
        log(f"  ⚠️  Excel no encontrado: {excel_path}", "WARN")
        log("     Los alumnos de la academia no se marcarán en el dashboard.", "WARN")
        return alumnos

    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            # Detectar cabecera (primera fila con texto)
            header_row = None
            header_idx = 0
            for i, row in enumerate(rows):
                if any(cell is not None and str(cell).strip() for cell in row):
                    header_row = [str(c).upper().strip() if c else "" for c in row]
                    header_idx = i
                    break
            if header_row is None:
                continue

            log(f"  📊 Hoja '{ws.title}' — columnas: {header_row[:8]}")

            # Buscar columnas de nombre, tribunal y baremo/puntuacion
            col_nombre   = _find_col(header_row, ["NOMBRE", "APELLIDO", "CANDIDATO", "OPOSITOR", "ALUMNO"])
            col_tribunal = _find_col(header_row, ["TRIBUNAL", "TRIB", "N_TRIBUNAL", "NUM_TRIBUNAL"])
            col_baremo   = _find_col(header_row, ["BAREMO", "NOTA", "PUNTUACION", "TOTAL", "MEDIA"])

            if col_nombre is None:
                log(f"  ⚠️  No se encontró columna de nombre en hoja '{ws.title}'", "WARN")
                continue

            for row in rows[header_idx + 1:]:
                if not row or all(c is None for c in row):
                    continue
                nombre_val = row[col_nombre] if col_nombre < len(row) else None
                if not nombre_val:
                    continue
                nombre_str  = str(nombre_val).strip()
                nombre_norm = normalizar(nombre_str)
                tribunal    = str(row[col_tribunal]).strip() if col_tribunal is not None and col_tribunal < len(row) else ""
                baremo_raw  = row[col_baremo] if col_baremo is not None and col_baremo < len(row) else None
                try:
                    baremo = float(baremo_raw) if baremo_raw is not None else None
                except (ValueError, TypeError):
                    baremo = None
                alumnos[nombre_norm] = {
                    "nombre_real": nombre_str,
                    "tribunal":    get_tribunal_num(tribunal) if tribunal else "",
                    "baremo":      baremo,
                }

        log(f"  🎓 {len(alumnos)} alumnos de la academia cargados.")
    except ImportError:
        log("  ⚠️  openpyxl no instalado. Ejecuta: pip install openpyxl", "WARN")
    except Exception as e:
        log(f"  ⚠️  Error leyendo Excel: {e}", "WARN")

    return alumnos


def _find_col(header_row, keywords):
    """Devuelve el índice de la primera columna que contiene alguna keyword."""
    for kw in keywords:
        for i, h in enumerate(header_row):
            if kw in normalizar(h):
                return i
    return None


def is_academy_student(nombre_texto, academy_students):
    """Comprueba si algún nombre del Excel aparece en el texto dado."""
    texto_norm = normalizar(str(nombre_texto))
    for nombre_norm in academy_students:
        # Comprobar si el nombre (o sus partes) aparece en el texto
        partes = nombre_norm.split()
        if len(partes) >= 2:
            # Al menos apellido + parte del nombre
            if partes[0] in texto_norm and partes[1] in texto_norm:
                return nombre_norm
        elif len(partes) == 1 and len(partes[0]) > 4:
            if partes[0] in texto_norm:
                return nombre_norm
    return None


# ─── Email ────────────────────────────────────────────────────────────────────
def enviar_email(config, items_nuevos):
    usuario  = config.get("gmail_usuario", "")
    password = config.get("gmail_contrasena_app", "")
    destino  = config.get("email_destino", "")

    if not usuario or not password or not destino:
        log("  ⚠️  Email no configurado en config_ef.json.", "WARN")
        return

    n = len(items_nuevos)
    asunto = f"[EF CARM 2026] {n} publicación{'es' if n > 1 else ''} nueva{'s' if n > 1 else ''}"

    filas = ""
    for it in items_nuevos:
        tipo_badge = "🔔 Citación" if it["tipo"] == "citacion" else "📋 Calificación"
        filas += f"""
        <tr>
          <td style="padding:6px 12px;border-bottom:1px solid #eee;">{it['tribunal']}</td>
          <td style="padding:6px 12px;border-bottom:1px solid #eee;">{tipo_badge}</td>
          <td style="padding:6px 12px;border-bottom:1px solid #eee;">{it['text']}</td>
          <td style="padding:6px 12px;border-bottom:1px solid #eee;">
            <a href="{it['url']}" style="color:#1a73e8;">PDF</a>
          </td>
        </tr>"""

    html = f"""
    <html><body style="font-family:Arial,sans-serif;color:#333;">
      <h2 style="color:#d32f2f;">🔔 Nuevas publicaciones EF — CARM 2026</h2>
      <p><strong>Fecha:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
      <table style="border-collapse:collapse;width:100%;margin-top:12px;">
        <thead>
          <tr style="background:#f5f5f5;">
            <th style="padding:8px 12px;text-align:left;">Tribunal</th>
            <th style="padding:8px 12px;text-align:left;">Tipo</th>
            <th style="padding:8px 12px;text-align:left;">Documento</th>
            <th style="padding:8px 12px;text-align:left;">Enlace</th>
          </tr>
        </thead>
        <tbody>{filas}</tbody>
      </table>
      <p style="margin-top:20px;color:#888;font-size:12px;">Monitor automático EF educarm CARM 2026</p>
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


# ─── AJAX ─────────────────────────────────────────────────────────────────────
def ajax_post(action, extra=None):
    data = {"anyo": ANYO, "convocatoria": CONV}
    if extra:
        data.update(extra)
    url = (
        f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
        f"&module=publicacionesTribunales&action={action}"
    )
    r = SESSION.post(url, data=data, timeout=20)
    r.raise_for_status()
    return r.json()


def get_especialidades(todas=False):
    resp  = ajax_post("ajaxOpcionesEspecialidad", {"codCuerpo": COD_CUERPO})
    codes = resp["codEspecialidades"].split("#")
    names = resp["denEspecialidades"].split("#")
    result = []
    for c, n in zip(codes, names):
        if not c:
            continue
        if todas or any(f in normalizar(n) for f in FILTRO_ESP):
            result.append({"code": c, "name": n})
    return result


def get_tribunales(cod_especialidad):
    resp = ajax_post("ajaxOpcionesTribunal", {
        "codCuerpo":      COD_CUERPO,
        "codEspecialidad": cod_especialidad,
    })
    return [c for c in resp["codTribunales"].split("#") if c]


def get_publicaciones_html(cod_especialidad, cod_tribunal):
    url = (
        f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
        f"&module=publicacionesTribunales&action=getPublicaciones"
        f"&anyo={ANYO}&convocatoria={CONV}"
    )
    data = {
        "cuerpos_lista":        COD_CUERPO,
        "especialidades_lista": cod_especialidad,
        "tribunales_lista":     cod_tribunal,
    }
    r = SESSION.post(url, data=data, timeout=20)
    r.raise_for_status()
    r.encoding = "iso-8859-1"
    return r.text


def extraer_publicaciones(html):
    soup = BeautifulSoup(html, "html.parser")
    pubs = []
    KW   = (".pdf", "download", "fichero", "documento", "getfile", "descarga")

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        text = a.get_text(strip=True)
        if not text or len(text) < 3:
            continue
        if any(kw in href.lower() for kw in KW):
            href = _fix_url(href)
            pubs.append({"url": href, "text": text})

    if not pubs:
        for a in soup.select("table a[href]"):
            href = a["href"].strip()
            text = a.get_text(strip=True)
            if href and text and len(text) > 2 and not href.startswith("#"):
                href = _fix_url(href)
                pubs.append({"url": href, "text": text})

    return pubs


def _fix_url(href):
    if href.startswith("//"):
        return "https:" + href
    elif href.startswith("/"):
        return "https://servicios.educarm.es" + href
    elif not href.startswith("http"):
        return "https://servicios.educarm.es/admin/" + href
    return href


# ─── Descarga ─────────────────────────────────────────────────────────────────
def download_pdf(url, dest_path):
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    if dest_path.exists():
        return True  # ya descargado
    try:
        r = SESSION.get(url, timeout=30, stream=True)
        r.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)
        return True
    except Exception as e:
        log(f"  ⚠️  Error descargando {url}: {e}", "WARN")
        return False


def build_dest_path(tipo, trib_code, text, timestamp):
    """
    Construye la ruta de destino del PDF:
      citaciones/t1_20260615_nombre_documento.pdf
    """
    tnum     = get_tribunal_num(trib_code)
    fecha    = datetime.fromisoformat(timestamp).strftime("%Y%m%d") if timestamp else datetime.now().strftime("%Y%m%d")
    filename = sanitize_filename(text)
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    filename = f"t{tnum}_{fecha}_{filename}"

    if tipo == "citacion":
        return CITACIONES_DIR / filename
    elif tipo == "calificacion":
        return CALIFIC_DIR / filename
    else:
        return CALIFIC_DIR / ("otros_" + filename)


# ─── Barrido ──────────────────────────────────────────────────────────────────
def scan():
    log("🚀 Barrido EF...")
    found = {}

    try:
        SESSION.get(
            f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
            f"&module=publicacionesTribunales&anyo={ANYO}&convocatoria={CONV}",
            timeout=15,
        )

        especialidades = get_especialidades()
        if not especialidades:
            log("  ⚠️  No se encontraron especialidades EF. Ejecuta: python monitor_ef.py --discover", "WARN")
            return found

        log(f"  📚 Especialidades EF: {[e['name'] for e in especialidades]}")

        for esp in especialidades:
            tribunales = get_tribunales(esp["code"])
            log(f"  → {esp['name']}: {len(tribunales)} tribunal(es)")

            for trib_code in tribunales:
                html = get_publicaciones_html(esp["code"], trib_code)
                pubs = extraer_publicaciones(html)

                for pub in pubs:
                    tipo = classify_doc(pub["text"])
                    key  = make_key(esp["name"], trib_code, pub["url"])
                    found[key] = {
                        "especialidad":  esp["name"],
                        "trib_code":     trib_code,
                        "tribunal":      f"T{get_tribunal_num(trib_code)}",
                        "url":           pub["url"],
                        "text":          pub["text"],
                        "tipo":          tipo,
                        "timestamp":     datetime.now().isoformat(),
                    }

    except Exception as e:
        log(f"❌ Error durante barrido: {e}", "ERROR")
        import traceback; traceback.print_exc()

    log(f"  📊 Total encontradas: {len(found)}")
    return found


# ─── Procesar ─────────────────────────────────────────────────────────────────
def process_results(found, state, config, academy_students):
    known     = state.get("known", {})
    new_items = []

    for key, pub in found.items():
        if key not in known:
            new_items.append(pub)
        known[key] = pub  # actualizar siempre (por si cambia algo)

    # Descargar todo (nuevo y ya conocido por si falló antes)
    for key, pub in found.items():
        dest = build_dest_path(pub["tipo"], pub["trib_code"], pub["text"], pub.get("timestamp"))
        if not dest.exists():
            if download_pdf(pub["url"], dest):
                log(f"  💾 {pub['tipo'].upper()} {pub['tribunal']} → {dest.name}")

    if new_items:
        log(f"\n🆕 ¡{len(new_items)} NOVEDAD(ES)!")
        for it in new_items:
            log(f"  → [{it['tipo'].upper()}] {it['tribunal']}: {it['text']}")
        enviar_email(config, new_items)
        beep(4)
    else:
        log("  ✅ Sin novedades")

    state["known"]     = known
    state["last_scan"] = datetime.now().isoformat()
    state["scans"]     = state.get("scans", 0) + 1
    save_state(state)

    # Regenerar HTML del dashboard
    generate_dashboard(known, academy_students)

    return len(new_items)


# ─── Dashboard HTML ───────────────────────────────────────────────────────────
def generate_dashboard(known, academy_students):
    WEB_DIR.mkdir(parents=True, exist_ok=True)

    # Preparar datos para el HTML
    pubs_list = sorted(known.values(), key=lambda x: x.get("timestamp", ""), reverse=True)

    # Enriquecer con flag de academia
    for pub in pubs_list:
        # Buscar alumnos de academia en el texto del PDF (por tribunal)
        tnum    = get_tribunal_num(pub.get("trib_code", ""))
        matches = [
            info["nombre_real"]
            for norm, info in academy_students.items()
            if info.get("tribunal") == tnum
        ]
        pub["alumnos_academia"] = matches

    # Agrupar por tribunal
    tribunales = {}
    for pub in pubs_list:
        t = pub["tribunal"]
        tribunales.setdefault(t, {"citaciones": [], "calificaciones": [], "otros": []})
        tipo = pub.get("tipo", "otro")
        if tipo == "citacion":
            tribunales[t]["citaciones"].append(pub)
        elif tipo == "calificacion":
            tribunales[t]["calificaciones"].append(pub)
        else:
            tribunales[t]["otros"].append(pub)

    last_scan = datetime.now().strftime("%d/%m/%Y %H:%M")
    n_cit  = sum(1 for p in pubs_list if p.get("tipo") == "citacion")
    n_cal  = sum(1 for p in pubs_list if p.get("tipo") == "calificacion")
    n_trib = len(tribunales)
    n_acad = sum(1 for p in pubs_list if p.get("alumnos_academia"))

    # Serializar datos como JSON para el HTML
    data_json = json.dumps({
        "last_scan":   last_scan,
        "publicaciones": pubs_list,
        "tribunales":  {t: {k: len(v) for k, v in docs.items()} for t, docs in tribunales.items()},
        "stats": {
            "citaciones":   n_cit,
            "calificaciones": n_cal,
            "tribunales":   n_trib,
            "con_academia": n_acad,
        }
    }, ensure_ascii=False)

    # Generar filas de tabla por tribunal
    trib_rows = ""
    for t in sorted(tribunales.keys()):
        docs = tribunales[t]
        nc = len(docs["citaciones"])
        nk = len(docs["calificaciones"])
        alumnos_en_trib = [
            info["nombre_real"]
            for info in academy_students.values()
            if info.get("tribunal") == get_tribunal_num(t)
        ]
        badge_academia = (
            f'<span class="badge badge-academy">🎓 {len(alumnos_en_trib)} academia</span>'
            if alumnos_en_trib else ""
        )
        trib_rows += f"""
        <tr>
          <td><strong>{t}</strong></td>
          <td>{"✅ " + str(nc) if nc else "—"}</td>
          <td>{"✅ " + str(nk) if nk else "—"}</td>
          <td>{badge_academia}</td>
        </tr>"""

    # Generar filas de publicaciones
    pub_rows = ""
    for pub in pubs_list[:50]:  # Máx 50 en el feed
        tipo_badge = (
            '<span class="badge badge-cit">Citación</span>'
            if pub.get("tipo") == "citacion"
            else '<span class="badge badge-cal">Calificación</span>'
        )
        fecha = pub.get("timestamp", "")[:10] if pub.get("timestamp") else ""
        alum  = pub.get("alumnos_academia", [])
        acad_badge = (
            f'<span class="badge badge-academy">🎓 {", ".join(alum[:2])}</span>'
            if alum else ""
        )
        pub_rows += f"""
        <tr class="{'row-academy' if alum else ''}">
          <td>{pub.get('tribunal','')}</td>
          <td>{tipo_badge}</td>
          <td><a href="{pub.get('url','')}" target="_blank">{pub.get('text','')[:60]}</a> {acad_badge}</td>
          <td>{fecha}</td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Oposiciones EF CARM 2026 — Dashboard</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Segoe UI', Arial, sans-serif;
    background: #f0f2f5;
    color: #333;
    font-size: 14px;
  }}
  header {{
    background: linear-gradient(135deg, #1e3a5f, #2563eb);
    color: white;
    padding: 20px 24px 16px;
  }}
  header h1 {{ font-size: 20px; font-weight: 700; }}
  header p  {{ font-size: 12px; opacity: 0.8; margin-top: 4px; }}
  .container {{ padding: 16px; max-width: 1100px; margin: 0 auto; }}

  /* Stats */
  .stats {{
    display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px;
  }}
  .stat-card {{
    background: white; border-radius: 10px; padding: 14px 20px;
    flex: 1; min-width: 120px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08);
    text-align: center;
  }}
  .stat-num  {{ font-size: 28px; font-weight: 700; color: #2563eb; }}
  .stat-label {{ font-size: 11px; color: #888; margin-top: 2px; }}

  /* Sections */
  .section {{
    background: white; border-radius: 10px; padding: 16px;
    margin-bottom: 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08);
  }}
  .section h2 {{
    font-size: 14px; font-weight: 700; color: #1e3a5f;
    margin-bottom: 12px; padding-bottom: 8px;
    border-bottom: 2px solid #e5e7eb;
  }}

  /* Tables */
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th {{
    background: #f8fafc; text-align: left;
    padding: 8px 10px; font-weight: 600; color: #555;
    border-bottom: 2px solid #e5e7eb;
  }}
  td {{ padding: 7px 10px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: #fafbff; }}
  tr.row-academy td {{ background: #fff8e1; }}
  tr.row-academy:hover td {{ background: #fff3cd; }}

  /* Badges */
  .badge {{
    display: inline-block; border-radius: 12px;
    padding: 2px 8px; font-size: 11px; font-weight: 600;
  }}
  .badge-cit     {{ background: #dbeafe; color: #1d4ed8; }}
  .badge-cal     {{ background: #dcfce7; color: #166534; }}
  .badge-academy {{ background: #fef9c3; color: #92400e; }}

  a {{ color: #2563eb; text-decoration: none; }}
  a:hover {{ text-decoration: underline; }}

  .update-time {{
    font-size: 11px; color: #aaa; text-align: right; margin-top: 4px;
  }}
  .legend {{
    font-size: 12px; color: #888; margin-top: 10px;
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  }}
</style>
</head>
<body>

<header>
  <h1>📋 Oposiciones EF — CARM 2026</h1>
  <p>Monitor de publicaciones de tribunales · Educación Física</p>
</header>

<div class="container">

  <!-- Stats -->
  <div class="stats" style="margin-top:16px;">
    <div class="stat-card">
      <div class="stat-num">{n_trib}</div>
      <div class="stat-label">Tribunales</div>
    </div>
    <div class="stat-card">
      <div class="stat-num" style="color:#1d4ed8;">{n_cit}</div>
      <div class="stat-label">Citaciones</div>
    </div>
    <div class="stat-card">
      <div class="stat-num" style="color:#166534;">{n_cal}</div>
      <div class="stat-label">Calificaciones</div>
    </div>
    <div class="stat-card">
      <div class="stat-num" style="color:#92400e;">{n_acad}</div>
      <div class="stat-label">Docs con alumnos academia</div>
    </div>
  </div>

  <!-- Resumen por tribunal -->
  <div class="section">
    <h2>📊 Estado por tribunal</h2>
    <table>
      <thead>
        <tr>
          <th>Tribunal</th>
          <th>Citaciones</th>
          <th>Calificaciones</th>
          <th>Academia</th>
        </tr>
      </thead>
      <tbody>
        {trib_rows}
      </tbody>
    </table>
  </div>

  <!-- Últimas publicaciones -->
  <div class="section">
    <h2>🕐 Últimas publicaciones</h2>
    <table>
      <thead>
        <tr>
          <th style="width:70px">Tribunal</th>
          <th style="width:110px">Tipo</th>
          <th>Documento</th>
          <th style="width:90px">Fecha</th>
        </tr>
      </thead>
      <tbody>
        {pub_rows}
      </tbody>
    </table>
    <div class="legend">
      <span class="badge badge-academy">🎓 Alumno academia</span>
      Las filas en amarillo corresponden a tribunales con alumnos de tu academia.
    </div>
  </div>

  <div class="update-time">Última actualización: {last_scan}</div>

</div>

<script>
// Datos completos embebidos para uso futuro
const DATA = {data_json};
</script>

</body>
</html>"""

    out = WEB_DIR / "index.html"
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"  🌐 Dashboard generado → {out}")
    # En CI también imprimimos la URL de Pages para el log
    if IS_CI:
        repo = os.getenv("GITHUB_REPOSITORY", "usuario/repo")
        log(f"  🔗 URL GitHub Pages: https://{repo.split('/')[0]}.github.io/{repo.split('/')[1]}/", "INFO")


# ─── Modo descubrimiento ──────────────────────────────────────────────────────
def discover():
    log("🔍 Consultando todas las especialidades disponibles...")
    try:
        SESSION.get(
            f"{BASE_URL}?aplicacion=PUBLICACIONES_TRIBUNALES"
            f"&module=publicacionesTribunales&anyo={ANYO}&convocatoria={CONV}",
            timeout=15,
        )
        todas = get_especialidades(todas=True)
        log(f"\n  {len(todas)} especialidades en CARM 2026:")
        for e in todas:
            marcado = " ← MONITORIZADA" if any(f in normalizar(e["name"]) for f in FILTRO_ESP) else ""
            log(f"    [{e['code']}] {e['name']}{marcado}")
    except Exception as ex:
        log(f"❌ {ex}", "ERROR")


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Monitor publicaciones EF — CARM 2026")
    parser.add_argument("--once",     action="store_true", help="Un solo barrido")
    parser.add_argument("--interval", type=int, default=DEFAULT_INTERVAL,
                        help=f"Minutos entre barridos (default: {DEFAULT_INTERVAL})")
    parser.add_argument("--discover", action="store_true", help="Listar todas las especialidades")
    parser.add_argument("--reset",    action="store_true", help="Borrar estado y empezar de cero")
    args = parser.parse_args()

    CITACIONES_DIR.mkdir(parents=True, exist_ok=True)
    CALIFIC_DIR.mkdir(parents=True, exist_ok=True)
    WEB_DIR.mkdir(parents=True, exist_ok=True)

    config = load_config()

    log("╔═══════════════════════════════════════════════╗")
    log("║   MONITOR EF — PUBLICACIONES TRIBUNALES 2026  ║")
    log("╚═══════════════════════════════════════════════╝")

    if args.reset:
        if STATE_FILE.exists():
            STATE_FILE.unlink()
            log("🗑️  Estado reiniciado.")

    if args.discover:
        discover()
        return

    # Cargar alumnos de la academia
    academy_students = load_academy_students(config.get("excel_alumnos", ""))

    if args.reset:
        return

    state = load_state()
    log(f"  Publicaciones conocidas: {len(state.get('known', {}))}")
    log(f"  Alumnos academia: {len(academy_students)}")
    log(f"  Intervalo: cada {args.interval} min")
    log(f"  Email destino: {config.get('email_destino','(no configurado)')}")
    log(f"  Citaciones → {CITACIONES_DIR}")
    log(f"  Calificaciones → {CALIFIC_DIR}")
    log(f"  Dashboard → {WEB_DIR / 'index.html'}")

    scan_count = 0
    while True:
        scan_count += 1
        log(f"\n── Barrido #{scan_count} ({datetime.now().strftime('%H:%M:%S')}) ──")
        found = scan()
        if found is not None:
            process_results(found, state, config, academy_students)
            state = load_state()

        if args.once:
            log("\n✅ Barrido único completado.")
            break

        log(f"  ⏳ Próximo barrido en {args.interval} min...")
        try:
            time.sleep(args.interval * 60)
        except KeyboardInterrupt:
            log("\n👋 Monitor detenido.")
            break


if __name__ == "__main__":
    main()
