# app.py
import base64
from flask import Flask, request, jsonify

# Importa la lógica de Word
from lector_word import (
    procesar,                 # modo clásico: 1 Word + 1 Excel
    procesar_grouped,         # modo por título: varios Word (un archivo por título)
    headings_from_docx        # listar headings para panel/depuración
)

app = Flask(__name__)


@app.route("/", methods=["GET"])
def health():
    return "Unificador de Títulos: OK", 200


@app.route("/api/headings", methods=["POST"])
def api_headings():
    """
    Entrada:
      {
        "archivos": [{"name": "...", "content": "<BASE64>"}]
      }
    Salida:
      { "<name>": [ {"level": 1, "text": "..."}, ... ], ... }
    """
    data = request.get_json(force=True, silent=False)
    out = {}
    for a in data.get("archivos", []):
        name = a.get("name", "archivo.docx")
        content_b64 = a.get("content")
        if not content_b64:
            out[name] = []
            continue
        try:
            b = base64.b64decode(content_b64)
            out[name] = headings_from_docx(b)
        except Exception as e:
            out[name] = []
    return jsonify(out)


@app.route("/api/merge", methods=["POST"])
def api_merge():
    """
    Modos:
    - Clásico (por defecto): devuelve un objeto con 2 claves fijas:
        {
          "unificado.docx": "<BASE64>",
          "tablas.xlsx": "<BASE64>"
        }

    - Agrupado por título (group_by_title = true):
        * Si return_array = true:
            { "files": [ {"filename":"A.docx","content":"<BASE64>"}, ... ] }
          (RECOMENDADO para Power Automate)
        * Si return_array = false (compat):
            { "A.docx":"<BASE64>", "B.docx":"<BASE64>", ... }
    """
    data = request.get_json(force=True, silent=False)

    # Entrada de archivos [{ name, content(base64) }]
    archivos = []
    for a in data.get("archivos", []):
        if not a or "content" not in a:
            continue
        try:
            archivos.append({
                "name": a.get("name", "archivo.docx"),
                "content": base64.b64decode(a["content"])
            })
        except Exception:
            # ignora archivos corruptos
            pass

    niveles = data.get("niveles", [1, 2, 3])
    titulos = data.get("titulos_exactos", [])
    enforce_whitelist = bool(data.get("enforce_whitelist", False))

    group_by_title = bool(data.get("group_by_title", False))
    group_level = int(data.get("group_level", 1))           # 1=H1, 2=H2, 3=H3
    return_array = bool(data.get("return_array", False))    # para PA

    if group_by_title:
        grouped = procesar_grouped(
            archivos=archivos,
            group_level=group_level,
            titulos_objetivo=titulos,
            enforce_whitelist=enforce_whitelist,
        )
        if return_array:
            files = [
                {"filename": k, "content": base64.b64encode(v).decode("utf-8")}
                for k, v in grouped.items()
            ]
            return jsonify({"files": files})

        # compat: objeto con claves dinámicas
        out = {k: base64.b64encode(v).decode("utf-8") for k, v in grouped.items()}
        return jsonify(out)

    # Modo clásico: 1 Word + 1 Excel
    result = procesar(
        archivos=archivos,
        niveles=niveles,
        titulos=titulos,
        enforce_whitelist=enforce_whitelist,
    )
    out = {k: base64.b64encode(v).decode("utf-8") for k, v in result.items()}
    return jsonify(out)


if __name__ == "__main__":
    # para pruebas locales
    app.run(host="0.0.0.0", port=8000, debug=False)
