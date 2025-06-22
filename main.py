from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import requests
from io import BytesIO
import os
import traceback

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "<h1>Servidor Plan de Retiro activo ✅</h1>"

@app.route("/modificar-plan-retiro", methods=["POST"])
def modificar():
    try:
        data = request.get_json()

        wb = load_workbook("PlanDeRetiroInstrucciones.xlsx")

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

        ws_contact = wb["Cómo contactarme"]
        ws_contact["C8"].value = data.get("nombre_persona")

        img_url = "https://raw.githubusercontent.com/garciabanchs/excel-plan-retiro/main/imagen_circular.png"
        response = requests.get(img_url)
        if response.status_code == 200:
            img_bytes = BytesIO(response.content)
            img = Image(img_bytes)
            ws_contact.add_image(img, "A2")
        else:
            print(f"Error descargando imagen, status: {response.status_code}")

        output_dir = "downloads"
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, "PlanModificado.xlsx")
        wb.save(output_file)

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        traceback_str = traceback.format_exc()
        print(traceback_str)
        return jsonify({"error": str(e), "traceback": traceback_str}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
