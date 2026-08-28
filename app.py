from flask import Flask, request, jsonify
import os, requests
from datetime import datetime

app = Flask(__name__)

FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "AUTOPIC/DOCUMENTACIÓN/inventario.xlsx")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
REFRESH_TOKEN = os.getenv("MS_REFRESH_TOKEN")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")

def get_token():
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

def get_next_row(headers, sheet):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{FILE_PATH}:/workbook/worksheets/{sheet}/range(address='A7:A1000')"
    resp = requests.get(url, headers=headers)
    if resp.status_code!= 200:
        return 7
    for i, row in enumerate(resp.json().get("values", [])):
        if not row[0]:
            return 7 + i
    return 7 + len(resp.json().get("values", []))

@app.route("/")
def home():
    return "Autopic Bot v4 OK - Soporta INVENTARIO y FOLIOS VENDIDOS"

@app.route("/agregar", methods=["GET", "POST"])
def agregar():
    try:
        d = {**(request.args.to_dict()), **(request.get_json(silent=True) or {})}
        hoja = d.get("hoja", "GENERAL").upper()

        token = get_token()
        h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        fila = get_next_row(h, hoja)

        # --- LÓGICA PARA FOLIOS VENDIDOS ---
        if hoja == "FOLIOS VENDIDOS":
            id_p = d.get("id", "")
            nombre = d.get("nombre", "")
            desc = d.get("descripcion", "")
            cant = d.get("cantidad", "")
            unidad = d.get("unidad", "PZA")
            fecha = d.get("fecha", datetime.now().strftime("%Y-%m-%d"))
            cliente = d.get("cliente", "")
            num_serie = d.get("numero_serie", d.get("serie", ""))
            proveedor = d.get("proveedor", "")
            estado = d.get("estado", "")
            embalaje = d.get("embalaje", "")

            if not id_p:
                return jsonify({"ok": False, "error": "Falta ID para FOLIOS VENDIDOS"}), 400

            valores = [[id_p, nombre, desc, cant, unidad, fecha, cliente, num_serie, proveedor, estado, embalaje]]
            url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{FILE_PATH}:/workbook/worksheets/{hoja}/range(address='A{fila}:K{fila}')"
            r = requests.patch(url, headers=h, json={"values": valores})
            r.raise_for_status()
            return jsonify({"ok": True, "hoja": hoja, "fila": fila, "tipo": "VENTA"})

        # --- LÓGICA PARA INVENTARIO GENERAL, DELTA, PHOENIX, CARLO ---
        else:
            id_p = d.get("id", "")
            nombre = d.get("nombre", "")
            desc = d.get("descripcion", "")
            cant = d.get("cantidad", "")
            unidad = d.get("unidad", "PIEZA")
            ubicacion = d.get("ubicacion", "")
            proveedor = d.get("proveedor", "")
            num_serie = d.get("numero_serie", d.get("serie", ""))
            marca = d.get("marca", "")

            if not id_p:
                return jsonify({"ok": False, "error": "Falta ID"}), 400

            # A:H
            url_ah = f"https://graph.microsoft.com/v1.0/me/drive/root:/{FILE_PATH}:/workbook/worksheets/{hoja}/range(address='A{fila}:H{fila}')"
            valores_ah = [[id_p, nombre, desc, cant, unidad, ubicacion, proveedor, num_serie]]
            r1 = requests.patch(url_ah, headers=h, json={"values": valores_ah})
            r1.raise_for_status()

            # J=Marca
            if marca:
                url_j = f"https://graph.microsoft.com/v1.0/me/drive/root:/{FILE_PATH}:/workbook/worksheets/{hoja}/range(address='J{fila}:J{fila}')"
                requests.patch(url_j, headers=h, json={"values": [[marca]]})

            return jsonify({"ok": True, "hoja": hoja, "fila": fila, "tipo": "INVENTARIO"})

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
