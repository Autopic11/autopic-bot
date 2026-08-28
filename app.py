from flask import Flask, request, jsonify
import os
import requests
from datetime import datetime

app = Flask(__name__)

FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "AUTOPIC/DOCUMENTACIÓN/inventario.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "Hoja2")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
REFRESH_TOKEN = os.getenv("MS_REFRESH_TOKEN")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "") # si no tienes, déjalo vacío

def get_access_token():
    url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
        "scope": "offline_access Files.ReadWrite Files.ReadWrite.All Sites.ReadWrite.All User.Read"
    }
    if CLIENT_SECRET:
        data["client_secret"] = CLIENT_SECRET
    
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

@app.route("/")
def home():
    return f"Autopic Bot corriendo! Archivo: {FILE_PATH} Hoja: {SHEET_NAME}"

@app.route("/agregar", methods=["GET", "POST"])
def agregar():
    try:
        # 1. Lee datos de GET (?producto=) o POST (JSON)
        json_data = request.get_json(silent=True) or {}
        args_data = request.args.to_dict()
        # junta ambos
        data = {**args_data, **json_data}
        
        producto = data.get("producto", "")
        sku = data.get("sku", "")
        cantidad = data.get("cantidad", "")
        marca = data.get("marca", "")

        if not producto:
            return jsonify({"ok": False, "error": "Falta producto. Ejemplo: ?producto=Filtro&sku=FA-001&cantidad=5"}), 400

        # 2. Conecta a OneDrive
        token = get_access_token()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

        # 3. Obtener última fila usada
        file_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{FILE_PATH}:/workbook/worksheets/{SHEET_NAME}/usedRange"
        resp = requests.get(file_url, headers=headers)
        
        if resp.status_code != 200:
            # si la hoja está vacía, empieza en fila 1
            next_row = 2
        else:
            used = resp.json()
            row_count = used.get("rowCount", 1)
            next_row = row_count + 1

        # 4. Escribir fila nueva: A=Fecha, B=Producto, C=SKU, D=Cantidad, E=Marca
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        valores = [[fecha, producto, sku, cantidad, marca]]
        
        range_address = f"A{next_row}:E{next_row}"
        patch_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{FILE_PATH}:/workbook/worksheets/{SHEET_NAME}/range(address='{range_address}')"
        
        body = {"values": valores}
        r2 = requests.patch(patch_url, headers=headers, json=body)
        r2.raise_for_status()

        return jsonify({
            "ok": True, 
            "mensaje": f"Agregado en fila {next_row}",
            "datos": {"fecha": fecha, "producto": producto, "sku": sku, "cantidad": cantidad, "marca": marca}
        })

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
