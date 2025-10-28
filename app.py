import os
import base64
from flask import Flask, request, jsonify, abort
from lector_word import procesar, headings_from_docx

app = Flask(__name__)

# --- Seguridad opcional por API_KEY (agregar en Render -> Environment Variables) ---
API_KEY = os.getenv("API_KEY", "").strip()

@app.before_request
def check_api_key():
    if API_KEY:  # solo si definiste API_KEY en el entorno
        client_key = request.headers.get("Authorization", "")
        if client_key != API_KEY:
            abort(401)

@app.route("/")
def home():
    return "Unificador de Títulos: OK"

@app.route("/api/headings", methods=["POST"])
def api_headings():
    """
    Body:
    {
      "archivos": [{"name":"x.docx","content":"<base64>"}]
    }
    Respuesta:
    {
      "x.docx": [{"level":1,"text":"..."}, ...],
      ...
    }
    """
    data = request.get_json(force=True, silent=False)
    archivos = data.get("archivos", [])
    out = {}
    for a in archivos:
        name = a["name"]
        content_b = base64.b64decode(a["content"])
        out[name] = headings_from_docx(content_b)
    return jsonify(out)

@app.route("/api/merge", methods=["POST"])
def api_merge():
    """
    Body:
    {
      "archivos": [{"name":"x.docx","content":"<base64>"}],
      "niveles": [1,2,3],
      "titulos_exactos": ["Introducción","Metodología"],
      "enforce_whitelist": false   # opcional: activar lista blanca de tu estructura
    }
    Respuesta:
    { "unificado.docx": "<base64>", "tablas.xlsx": "<base64>" }
    """
    data = request.get_json(force=True, silent=False)

    archivos = []
    for a in data.get("archivos", []):
        archivos.append({"name": a["name"], "content": base64.b64decode(a["content"])})
    niveles = data.get("niveles", [1, 2, 3])
    titulos = data.get("titulos_exactos", [])
    enforce_whitelist = bool(data.get("enforce_whitelist", False))

    res = procesar(archivos, niveles, titulos, enforce_whitelist=enforce_whitelist)
    out = {k: base64.b64encode(v).decode("utf-8") for k, v in res.items()}
    return jsonify(out)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
