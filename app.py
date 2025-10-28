# app.py
from flask import Flask, request, jsonify
from lector_word import procesar, headings_from_docx  # tu script ya funcional
import base64

app = Flask(__name__)

@app.route("/api/headings", methods=["POST"])
def get_headings():
    data = request.get_json()
    archivos = data.get("archivos", [])
    result = {}
    for a in archivos:
        content = base64.b64decode(a["content"])
        result[a["name"]] = headings_from_docx(content)
    return jsonify(result)

@app.route("/api/merge", methods=["POST"])
def merge():
    data = request.get_json()
    archivos = []
    for a in data.get("archivos", []):
        archivos.append({
            "name": a["name"],
            "content": base64.b64decode(a["content"])
        })
    niveles = data.get("niveles", [1, 2, 3])
    titulos_exactos = data.get("titulos_exactos", [])
    res = procesar(archivos, niveles, titulos_exactos)
    out = {k: base64.b64encode(v).decode("utf-8") for k, v in res.items()}
    return jsonify(out)

@app.route("/")
def home():
    return "Unificador de Títulos funcionando ✅"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
