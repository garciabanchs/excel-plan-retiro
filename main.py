from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
import os
import traceback

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
        ws_plan["C5"] = data.get("activo_financiero")
        ws_plan["C8"] = data.get("tasa_interes")
        ws_plan["C10"] = data.get("fraccion_ahorro")
       
        # Hoja Cómo contactarme para el nombre (solo modificar C8)
        ws_contact = wb["Cómo contactarme"]
        ws_contact["C8"] = data.get("nombre_persona")

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
