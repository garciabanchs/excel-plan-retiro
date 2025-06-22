import logging
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import requests
from io import BytesIO
import os
import traceback

logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "<h1>Servidor Plan de Retiro activo ✅</h1>"

@app.route("/modificar-plan-retiro", methods=["POST"])
def modificar():
    try:
        data = request.get_json()
        logging.debug(f"Datos recibidos: {data}")

        logging.debug("Cargando archivo Excel")
        wb = load_workbook("PlanDeRetiroInstrucciones.xlsx")

        logging.debug("Modificando hoja Plan de Retiro")
        ws_plan = wb["Plan de Retiro"]
        ws_plan["C2"] = data.get("edad_actual")
        ws_plan["C3"] = data.get("edad_retiro")
        ws_plan["C4"] = data.get("ingreso_anual")
        ws_plan["C5"] = data.get("activo_financiero")

        tasa_interes = data.get("tasa_interes")
        if tasa_interes is not None:
            ws_plan["C8"] = tasa_interes / 100
        else:
            ws_plan["C8"] = None

        fraccion_ahorro = data.get("fraccion_ahorro")
        if fraccion_ahorro is not None:
            ws_plan["C10"] = fraccion_ahorro / 100
        else:
            ws_plan["C10"] = None

        logging.debug("Modificando hoja Cómo contactarme")
        ws_contact = wb["Cómo contactarme"]
        ws_contact["C8"].value = data.get("nombre_persona")

        img_url = "https://raw.githubusercontent.com/garciabanchs/excel-plan-retiro/main/imagen_circular.png"
        logging.debug(f"Descargando imagen desde URL: {img_url}")
        response = requests.get(img_url)
        logging.debug(f"Status descarga imagen: {response.status_code}")
        if response.status_code == 200:
            img_bytes = BytesIO(response.content)
            img = Image(img_bytes)
            ws_contact.add_image(img, "A2")
            logging.debug("Imagen insertada en Excel")
        else:
            logging.error(f"Error descargando imagen, status: {response.status_code}")

        output_dir = "downloads"
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, "PlanModificado.xlsx")

        logging.debug("Guardando archivo modificado")
        wb.save(output_file)

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        traceback_str = traceback.format_exc()
        logging.error(f"Error en modificar-plan-retiro:\n{traceback_str}")
        return jsonify({"error": str(e), "traceback": traceback_str}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
