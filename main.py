from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
import os

app = Flask(__name__)
CORS(app)  # Permite CORS

@app.route("/")
def home():
    return "<h1>Servidor Plan de Retiro activo ✅</h1>"

@app.route("/modificar-plan-retiro", methods=["POST"])
def modificar():
    try:
        data = request.get_json()

        wb = load_workbook("PlanDeRetiroInstrucciones.xlsx")

        # Hoja Plan de Retiro para datos financieros
        ws_plan = wb["Plan de Retiro"]
        ws_plan["C2"] = data.get("edad_actual")
        ws_plan["C3"] = data.get("edad_retiro")
        ws_plan["C4"] = data.get("ingreso_anual")
        ws_plan["C8"] = data.get("activo_financiero")
        ws_plan["C10"] = data.get("tasa_interes")
        ws_plan["C12"] = data.get("fraccion_ahorro")  # Ajusta si quieres otra celda

        # Hoja Cómo contactarme para el nombre
        ws_contact = wb["Cómo contactarme"]
        ws_contact["C8"] = data.get("nombre_persona")

        output_file = "downloads/PlanModificado.xlsx"
        os.makedirs("downloads", exist_ok=True)
        wb.save(output_file)

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
