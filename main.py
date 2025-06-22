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
        ws = wb["Plan de Retiro"]  # Abrimos la hoja correcta

        # Escribir datos en las celdas específicas:
        ws["C2"] = data.get("edad_actual")
        ws["C3"] = data.get("edad_retiro")
        ws["C4"] = data.get("ingreso_anual")
        ws["C8"] = data.get("activo_financiero")
        ws["C10"] = data.get("tasa_interes")
        ws["C12"] = data.get("fraccion_ahorro")      # Cambia si prefieres otra celda
        ws["C14"] = data.get("nombre_persona")       # Cambia si prefieres otra celda

        output_file = "downloads/PlanModificado.xlsx"
        os.makedirs("downloads", exist_ok=True)
        wb.save(output_file)

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

