from flask import Flask, request, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route("/modificar-plan-retiro", methods=["POST"])
def modificar():
    data = request.get_json()

    # Cargar Excel base
    wb = load_workbook("PlanDeRetiroInstrucciones.xlsx")
    ws = wb.active

    # Aplicar modificaciones
    ws["C1"] = data["edad_actual"]
    ws["C2"] = data["edad_retiro"]
    ws["C3"] = data["ingreso_anual"]
    ws["C4"] = data["activo_financiero"]
    ws["C5"] = data["tasa_interes"]
    ws["C6"] = data["fraccion_ahorro"]
    ws["C7"] = data["nombre_persona"]

    output_file = "downloads/PlanModificado.xlsx"
    wb.save(output_file)

    return send_file(output_file, as_attachment=True)

if __name__ == "__main__":
    os.makedirs("downloads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)
