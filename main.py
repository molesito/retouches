"""
Servicio web en Flask para procesar archivos .docx recibidos desde n8n por HTTP
y aplicar:
1) Bordes negros a TODAS las tablas.
2) Eliminar la repetición de la primera fila cuando la tabla se parte entre páginas.

Endpoints:
- GET /health
- POST /process  (multipart/form-data con campo 'file' o cuerpo binario crudo)
"""

from flask import Flask, request, send_file, jsonify
from io import BytesIO
from docx import Document
from docx.oxml.shared import OxmlElement, qn
import datetime

app = Flask(__name__)

# -------- Utilidades DOCX --------
def aplicar_bordes_negros_tabla(table, ancho_octavos_pt: int = 8, color_hex: str = "000000") -> None:
    """
    Aplica bordes negros de línea simple a la tabla (externos e internos).
    """
    tbl = table._tbl
    tblPr = tbl.tblPr

    # Buscar <w:tblBorders> si ya existe
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)

    def _crear_borde(tag: str):
        el = OxmlElement(f'w:{tag}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(ancho_octavos_pt))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color_hex)
        return el

    # Limpiar existentes y establecer los seis bordes
    for tag in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        previo = tblBorders.find(qn(f'w:{tag}'))
        if previo is not None:
            tblBorders.remove(previo)
        tblBorders.append(_crear_borde(tag))

def eliminar_repeticion_fila_encabezado(table) -> None:
    """
    Quita la marca <w:tblHeader/> de las filas para que no se repitan
    automáticamente al dividir la tabla entre páginas.
    """
    for row in table.rows:
        tr = row._tr
        trPr = tr.trPr
        if trPr is None:
            continue
        a_borrar = []
        for child in list(trPr):
            if child.tag == qn('w:tblHeader'):
                a_borrar.append(child)
        for child in a_borrar:
            trPr.remove(child)

def procesar_docx(stream_entrada: BytesIO) -> BytesIO:
    """
    Procesa el DOCX en memoria, aplicando bordes y quitando repetición.
    Devuelve un BytesIO con el DOCX resultante.
    """
    doc = Document(stream_entrada)
    for table in doc.tables:
        aplicar_bordes_negros_tabla(table)
        eliminar_repeticion_fila_encabezado(table)

    salida = BytesIO()
    doc.save(salida)
    salida.seek(0)
    return salida

# -------- Endpoints --------
@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "timestamp": datetime.datetime.utcnow().isoformat() + "Z"})

@app.route('/process', methods=['POST'])
def process():
    try:
        in_bytes = None
        filename_in = 'input.docx'

        # 1) multipart/form-data con campo 'file'
        if 'file' in request.files:
            f = request.files['file']
            filename_in = f.filename or filename_in
            in_bytes = f.read()
        else:
            # 2) cuerpo binario crudo
            in_bytes = request.get_data(cache=False)
            if not in_bytes:
                return jsonify({"error": "No se recibió archivo"}), 400

        entrada = BytesIO(in_bytes)
        salida = procesar_docx(entrada)

        base = filename_in.rsplit('.', 1)[0] if '.' in filename_in else filename_in
        filename_out = f"{base}_procesado.docx"

        return send_file(
            salida,
            as_attachment=True,
            download_name=filename_out,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # Para pruebas locales: python main.py
    # Luego visitar: http://localhost:5000/health
    app.run(host='0.0.0.0', port=5000, debug=False)
