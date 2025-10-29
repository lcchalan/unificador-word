# -*- coding: utf-8 -*-
"""
Núcleo para unificar por títulos (H1/H2/H3) y extraer tablas.
Diseñado para funcionar en servidores (Render) sin escribir a disco.

Expone:
- headings_from_docx(doc_bytes: bytes) -> List[{"level":int,"text":str}]
- procesar(archivos, niveles, titulos_exactos, enforce_whitelist=False)
    -> {"unificado.docx": bytes, "tablas.xlsx": bytes}
"""

from typing import List, Dict, Tuple, Optional
import io, re, unicodedata

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docxcompose.composer import Composer
from openpyxl import Workbook

# ---------------------------------------------------------------------
# Normalización y utilidades
# ---------------------------------------------------------------------

_NUM_PREFIX = re.compile(r'^\s*(\d+[\.\)]\s*|\d+\s*-\s*)')

def quitar_acentos(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def norm(s: str) -> str:
    return re.sub(r'\s+', ' ', quitar_acentos((s or '').strip().lower()))

def base_title(s: str) -> str:
    return norm(_NUM_PREFIX.sub('', s or ''))

def estilo_es_titulo(style_name: str) -> Tuple[bool, int]:
    """
    Acepta Heading/Título 1..6 (ES/EN).
    """
    if not style_name:
        return (False, 0)
    s = norm(style_name)
    m = re.search(r'(heading|titulo|título)\s*(\d+)', s)
    return (bool(m), int(m.group(2)) if m else 0)

def iter_block_items(doc: Document):
    """
    Itera párrafos y tablas preservando orden.
    Devuelve ('p', Paragraph) o ('t', Table).
    """
    body = doc._element.body
    for child in list(body.iterchildren()):
        if child.tag.endswith('p'):
            yield ('p', Paragraph(child, doc))
        elif child.tag.endswith('tbl'):
            yield ('t', Table(child, doc))

def tabla_2d(tbl: Table) -> List[List[str]]:
    out = []
    for r in tbl.rows:
        row = []
        for c in r.cells:
            row.append((c.text or '').strip())
        out.append(row)
    return out
def procesar_grouped(archivos: List[Dict], group_level: int,
                     titulos_objetivo: List[str],
                     enforce_whitelist: bool = False) -> Dict[str, bytes]:
    """
    Genera un DOCX por cada título objetivo (exacto por texto 'base') en el nivel group_level.
    Ej.: group_level=1 (H1), titulos_objetivo=["A"] -> {"A.docx": bytes}

    Si titulos_objetivo está vacío, agrupa por TODOS los títulos encontrados en ese nivel.
    No genera Excel; solo los DOCX por título.
    """
    # normaliza selección de títulos
    objetivos_norm = set(base_title(t) for t in titulos_objetivo if (t or '').strip())

    # primero descubrimos todos los títulos candidatos del nivel indicado
    # y pre-creamos un Composer por cada uno
    composers: Dict[str, Composer] = {}     # clave: texto base, valor: composer
    nombres_visibles: Dict[str, str] = {}   # clave: texto base, valor: texto original (bonito) para filename

    def get_or_make_composer(key_norm: str, visible_text: str) -> Composer:
        if key_norm not in composers:
            tmp_doc = Document()
            tmp_stream = io.BytesIO()
            tmp_doc.save(tmp_stream); tmp_stream.seek(0)
            composers[key_norm] = Composer(Document(tmp_stream))
            # guarda el nombre visible (primera vez que aparece)
            nombres_visibles.setdefault(key_norm, visible_text.strip() or key_norm)
        return composers[key_norm]

    # 1) recolecta y decide qué títulos vamos a generar
    all_titles_norm: set = set()
    sample_visible_for: Dict[str, str] = {}
    for name, content in [(a['name'], a['content']) for a in archivos if a.get('content')]:
        bloques = _extraer_bloques(content)
        for kind, lvl, txt in bloques:
            if kind == 'h' and lvl == group_level and txt.strip():
                bt = base_title(txt)
                if enforce_whitelist and not allowed_by_whitelist(lvl, txt):
                    continue
                all_titles_norm.add(bt)
                sample_visible_for.setdefault(bt, txt)

    # si no se pidió lista específica, usamos todos los títulos encontrados
    target_keys = objetivos_norm if objetivos_norm else all_titles_norm

    # 2) por cada archivo, extrae los rangos de cada título target y apéndalos en su propio DOCX
    for name, content in [(a['name'], a['content']) for a in archivos if a.get('content')]:
        bloques = _extraer_bloques(content)
        # precalculamos rangos de todos los headings
        heads = [(i, b) for i, b in enumerate(bloques) if b[0] == 'h']
        # Para cada título del nivel objetivo:
        for key in target_keys:
            # buscamos todas las apariciones del título base en este doc
            indices = []
            for idx, b in enumerate(bloques):
                if b[0] == 'h' and b[1] == group_level and base_title(b[2]) == key:
                    if enforce_whitelist and not allowed_by_whitelist(b[1], b[2]):
                        continue
                    indices.append(idx)
            if not indices:
                continue

            # compón un 'part' por cada aparición y apéndalo al Composer de ese título
            comp = get_or_make_composer(key, sample_visible_for.get(key, key))
            for i_start in indices:
                # fin del rango: siguiente heading de igual o menor nivel
                j_end = len(bloques)
                for j in range(i_start + 1, len(bloques)):
                    if bloques[j][0] == 'h' and bloques[j][1] <= group_level:
                        j_end = j
                        break

                part = Document()
                # agrega el heading original
                h = part.add_heading(level=min(max(group_level, 1), 6))
                h.add_run(bloques[i_start][2])

                # agrega contenido
                for b in bloques[i_start + 1:j_end]:
                    if b[0] == 'p':
                        part.add_paragraph(b[2])
                    elif b[0] == 't':
                        rows = b[2]
                        if rows and rows[0]:
                            t = part.add_table(rows=len(rows), cols=len(rows[0]))
                            for r_i, row in enumerate(rows):
                                for c_i, val in enumerate(row):
                                    t.cell(r_i, c_i).text = val
                comp.append(part)

    # 3) serializa todos los compositores a bytes
    out: Dict[str, bytes] = {}
    for key_norm, comp in composers.items():
        if key_norm not in target_keys:
            continue
        nice = sample_visible_for.get(key_norm, key_norm)
        # crea nombre de archivo seguro
        safe = re.sub(r'[\\/:*?"<>|]+', '_', nice).strip()
        if not safe:
            safe = key_norm or "titulo"
        filename = f"{safe}.docx"
        stream = io.BytesIO()
        comp.doc.save(stream); stream.seek(0)
        out[filename] = stream.read()

    return out


# ---------------------------------------------------------------------
# (Opcional) Lista blanca basada en tu estructura H1/H2/H3
# Actívala pasando enforce_whitelist=True en procesar(...)
# ---------------------------------------------------------------------

H1_ALLOW = set(map(base_title, [
    "Introducción",
    "Diagnóstico estratégico",
    "Misión, visión y valores de la carrera",
    "Objetivos estratégicos",
    "Anexos (enlaces a documentos relevantes)",
    "Conclusiones",
    "Bibliografía",
]))

H2_ALLOW = set(map(base_title, [
    # Diagnóstico estratégico
    "Análisis FODA", "Análisis PESTEL", "Matriz Ansoff",
    # Misión, visión y valores
    "Misión", "Visión", "Valores", "Perfil profesional y perfil de egreso",
    # Macro-líneas
    "Misionalidad y desarrollo humano integral",
    "Educación para el futuro",
    "Ciudadanía global, internacionalización y relacionamiento estratégico",
    "Innovación, emprendimiento y sostenibilidad",
    # Otros
    "Otros indicadores relacionados con la Investigación",
]))

H3_ALLOW = set(map(base_title, [
    # Línea 1
    "01. Plan de formación integral del estudiante",
    "06. Plan de admisión, acogida y acompañamiento académico de estudiantes",
    "23. Plan de seguimiento y mejora de indicadores del perfil docente",
    "25. Plan de formación integral del docente",
    "26. Plan de mejora del proceso de evaluación integral docente",
    "Otras estrategias planificadas por la carrera relacionadas con la línea estratégica 1",
    # Línea 2
    "03. Plan implantación del marco de competencias UTPL",
    "04. Plan de prospectiva y creación de nueva oferta",
    "07. Plan de acciones curriculares para el fortalecimiento de las competencias genéricas",
    "11. Plan de fortalecimiento de prácticas preprofesionales y proyectos de vinculación",
    "12. Plan de fortalecimiento de criterios para la evaluación de la calidad de carreras y programas académicos",
    "13. Plan de acciones curriculares para el fortalecimiento de la empleabilidad del graduado UTPL",
    "16. Plan de mejora del proceso de elaboración y seguimiento de planes docentes",
    "18. Plan de mejora de ambientes de aprendizaje",
    "19. Plan de mejora de evaluación de los aprendizajes",
    "20. Plan de mejora del proceso de integración curricular",
    "21. Plan de mejora del proceso de titulación",
    "22. Plan de seguimiento y mejora de la labor tutorial",
    "Otras estrategias planificadas por la carrera relacionadas con la línea estratégica 2",
    # Línea 3
    "08. Plan de internacionalización del currículo",
    "24. Plan de intervención de personal académico en territorio",
    "Otras estrategias planificadas por la carrera relacionadas con la línea estratégica 3",
    # Línea 4
    "05. Plan de acciones académicas orientadas a la comunicación y promoción de la oferta",
    "09. Plan de innovación educativa",
    "10. Plan de implantación de metodologías activas en el currículo",
    "28. Plan de formación de líderes académicos",
    "29. Plan de posicionamiento institucional en innovación educativa",
    "30. Plan de investigación sobre innovación educativa, EaD, MP",
    "Otras estrategias planificadas por la carrera relacionadas con la línea estratégica 4",
]))

def allowed_by_whitelist(level: int, text: str) -> bool:
    bt = base_title(text)
    if level == 1:
        return bt in H1_ALLOW
    if level == 2:
        return bt in H2_ALLOW
    if level == 3:
        return bt in H3_ALLOW
    return False

# ---------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------

def headings_from_docx(doc_bytes: bytes) -> List[Dict]:
    """
    Devuelve lista de headings: [{"level": 1..3, "text": "..."}, ...]
    (Solo H1/H2/H3)
    """
    doc = Document(io.BytesIO(doc_bytes))
    out = []
    for kind, obj in iter_block_items(doc):
        if kind != 'p':
            continue
        style = getattr(obj.style, 'name', '')
        is_h, lvl = estilo_es_titulo(style)
        if is_h and 1 <= lvl <= 3:
            txt = (obj.text or '').strip()
            if txt:
                out.append({"level": lvl, "text": txt})
    return out

def _extraer_bloques(doc_bytes: bytes):
    """
    Convierte el docx a una lista ordenada de bloques:
      ('h', level, text)   para títulos
      ('p', 0, text)       para párrafos
      ('t', 0, tabla2d)    para tablas
    """
    doc = Document(io.BytesIO(doc_bytes))
    bloques = []
    for kind, obj in iter_block_items(doc):
        if kind == 'p':
            style = getattr(obj.style, 'name', '')
            is_h, lvl = estilo_es_titulo(style)
            txt = (obj.text or '').strip()
            if is_h and txt and 1 <= lvl <= 6:
                bloques.append(('h', lvl, txt))
            elif txt:
                bloques.append(('p', 0, txt))
        else:
            t2d = tabla_2d(obj)
            if t2d and any(any(cell for cell in row) for row in t2d):
                bloques.append(('t', 0, t2d))
    return bloques

def _rangos_por_titulo(bloques, niveles: List[int],
                       titulos_exactos_norm: set,
                       enforce_whitelist: bool):
    """
    Calcula rangos [i_start, i_end) para cada título incluido.
    Lógica de inclusión:
      - Solo H1/H2/H3
      - level ∈ niveles (si viene vacío, se ignora)
      - y (si titulos_exactos_norm no está vacío) el título base ∈ titulos_exactos_norm
      - y (si enforce_whitelist True) debe pasar lista blanca.
    """
    heads = [(i, b) for i, b in enumerate(bloques) if b[0] == 'h' and 1 <= b[1] <= 3]
    rangos = []
    for k, (i, b) in enumerate(heads):
        lvl, txt = b[1], b[2]
        include = True
        if niveles:
            include = include and (lvl in niveles)
        if titulos_exactos_norm:
            include = include and (base_title(txt) in titulos_exactos_norm)
        if enforce_whitelist:
            include = include and allowed_by_whitelist(lvl, txt)
        if not include:
            continue
        j = len(bloques)
        for kk in range(k + 1, len(heads)):
            if heads[kk][1][1] <= lvl:
                j = heads[kk][0]
                break
        rangos.append((i, j, lvl, txt))
    return rangos

def procesar(archivos: List[Dict], niveles: List[int], titulos_exactos: List[str],
             enforce_whitelist: bool = False) -> Dict[str, bytes]:
    """
    entradas:
      archivos = [{"name": "...", "content": bytes}, ...]
      niveles = [1,2,3]
      titulos_exactos = ["Introducción", "Metodología", ...]  (opcional)
      enforce_whitelist = True/False  (si True, usa la estructura oficial)
    salida:
      {"unificado.docx": bytes, "tablas.xlsx": bytes}
    """
    # Normaliza filtro por título
    titulos_exactos_norm = set(base_title(t) for t in titulos_exactos if (t or '').strip())

    # Documento base para composer
    base_doc = Document()
    tmp = io.BytesIO()
    base_doc.save(tmp)
    tmp.seek(0)
    composer = Composer(Document(tmp))

    # Excel de tablas
    wb = Workbook()
    ws = wb.active
    ws.title = "Tablas"
    ws.append(["Archivo", "Título", "Nivel", "#Tabla", "Fila", "Col", "Valor"])

    for name, content in [(a['name'], a['content']) for a in archivos if a.get('content')]:
        bloques = _extraer_bloques(content)
        rangos = _rangos_por_titulo(bloques, niveles, titulos_exactos_norm, enforce_whitelist)

        part = Document()
        tabla_idx = 0

        for (i_start, i_end, lvl, txt) in rangos:
            # Título
            h = part.add_heading(level=min(max(lvl, 1), 6))
            h.alignment = 0
            h.add_run(txt)

            # Contenido entre títulos
            for b in bloques[i_start + 1:i_end]:
                if b[0] == 'p':
                    part.add_paragraph(b[2])
                elif b[0] == 't':
                    rows = b[2]
                    if not rows:
                        continue
                    tabla = part.add_table(rows=len(rows), cols=len(rows[0]))
                    tabla_idx += 1
                    for r_i, row in enumerate(rows, start=1):
                        for c_i, val in enumerate(row, start=1):
                            tabla.cell(r_i - 1, c_i - 1).text = val
                            ws.append([name, txt, lvl, tabla_idx, r_i, c_i, val])
            part.add_paragraph('')  # separador

        composer.append(part)

    # DOCX a bytes
    out_docx = io.BytesIO()
    composer.doc.save(out_docx)
    out_docx.seek(0)

    # XLSX a bytes
    out_xlsx = io.BytesIO()
    wb.save(out_xlsx)
    out_xlsx.seek(0)

    return {"unificado.docx": out_docx.read(), "tablas.xlsx": out_xlsx.read()}

# ---------------------------------------------------------------------
# Solo para pruebas en local (NO se ejecuta en Render al importar)
# ---------------------------------------------------------------------
if __name__ == "__main__":
    # Ejemplo rápido en local:
    import os

    INPUT_DIR = os.environ.get("INPUT_DIR", ".")
    files = []
    for fname in os.listdir(INPUT_DIR):
        if fname.lower().endswith(".docx"):
            with open(os.path.join(INPUT_DIR, fname), "rb") as f:
                files.append({"name": fname, "content": f.read()})

    res = procesar(files, niveles=[1, 2, 3], titulos_exactos=[], enforce_whitelist=False)
    with open("unificado.docx", "wb") as f:
        f.write(res["unificado.docx"])
    with open("tablas.xlsx", "wb") as f:
        f.write(res["tablas.xlsx"])
    print("Generados unificado.docx y tablas.xlsx en la carpeta actual.")
