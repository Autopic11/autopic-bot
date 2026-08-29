from flask import Flask, request
import requests, os, json
from datetime import datetime

app = Flask(__name__)

# --- TUS DATOS DE META (de tu captura) ---
WHATSAPP_TOKEN = os.environ.get("WHATSAPP_TOKEN")
PHONE_ID = os.environ.get("PHONE_ID") or os.environ.get("PHONE_NUMBER_ID") or "1265790226622285"
VERIFY_TOKEN = os.environ.get("VERIFY_TOKEN") or "autopic123"

# --- TUS DATOS DE MICROSOFT (de tu captura) ---
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID")
MS_REFRESH_TOKEN = os.environ.get("MS_REFRESH_TOKEN")
ONEDRIVE_FILE_PATH = os.environ.get("ONEDRIVE_FILE_PATH") or "AUTOPIC/DOCUMENTACIÓN/INVENTARIO GENERAL PRUEBA.xlsx"
SHEET_NAME_DEFAULT = os.environ.get("SHEET_NAME", "Hoja2")

# Para renovar token de Microsoft
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "") # si no tienes, déjalo vacío
_ms_token_cache = {"token": None, "expiry": 0}

def get_ms_access_token():
    import time
    # si tenemos token cacheado y no expira en 5 min, usarlo
    if _ms_token_cache["token"] and time.time() < _ms_token_cache["expiry"] - 300:
        return _ms_token_cache["token"]

    url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    data = {
        "client_id": MS_CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": MS_REFRESH_TOKEN,
        "scope": "https://graph.microsoft.com/.default offline_access Files.ReadWrite"
    }
    if MS_CLIENT_SECRET:
        data["client_secret"] = MS_CLIENT_SECRET

    r = requests.post(url, data=data)
    print(f"RENOVAR MS TOKEN: {r.status_code} {r.text[:1000]}")
    if r.status_code == 200:
        j = r.json()
        _ms_token_cache["token"] = j["access_token"]
        _ms_token_cache["expiry"] = time.time() + j.get("expires_in", 3600)
        # Actualizar refresh token si viene uno nuevo
        if "refresh_token" in j:
            _ms_token_cache["refresh"] = j["refresh_token"]
        return j["access_token"]
    else:
        return None

def get_headers_graph():
    token = get_ms_access_token()
    return {"Authorization": f"Bearer {token}"} if token else {}

# --- FUNCIONES WHATSAPP ---
sesiones = {}

def enviar_texto(to, body):
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    payload = {"messaging_product": "whatsapp", "to": to, "type": "text", "text": {"body": body}}
    r = requests.post(url, headers=headers, json=payload)
    print(f"-> TEXTO a {to}: {r.status_code} {r.text[:800]}")
    return r

def enviar_botones(to, texto, botones):
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    btns = [{"type": "reply", "reply": {"id": b["id"], "title": b["title"]}} for b in botones]
    payload = {
        "messaging_product": "whatsapp", "to": to, "type": "interactive",
        "interactive": {"type": "button", "body": {"text": texto}, "action": {"buttons": btns}}
    }
    r = requests.post(url, headers=headers, json=payload)
    print(f"-> BOTONES a {to}: {r.status_code} {r.text[:800]}")
    return r

def get_excel(hoja):
    try:
        # Usar ruta de OneDrive directa
        encoded_path = ONEDRIVE_FILE_PATH.replace(" ", "%20")
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_path}:/workbook/worksheets/{hoja}/usedRange"
        h = get_headers_graph()
        r = requests.get(url, headers=h)
        print(f"GET EXCEL {hoja}: {r.status_code}")
        if r.status_code!= 200:
            print(r.text[:1500])
            return []
        return r.json().get('values', [])
    except Exception as e:
        print(f"ERROR get_excel: {e}")
        import traceback; traceback.print_exc()
        return []

def buscar_producto(hoja, modelo):
    valores = get_excel(hoja)
    if len(valores) < 2: return None
    encabezados = valores[0]
    for idx, fila in enumerate(valores[1:], start=2):
        fila_str = " ".join([str(c).upper() for c in fila])
        if modelo.upper() in fila_str:
            return {"fila": idx, "encabezados": encabezados, "datos": fila}
    return None

def escribir_fila(hoja, fila_num, array_valores):
    try:
        encoded_path = ONEDRIVE_FILE_PATH.replace(" ", "%20")
        col_fin = chr(64 + len(array_valores))
        rango = f"A{fila_num}:{col_fin}{fila_num}"
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_path}:/workbook/worksheets/{hoja}/range(address='{rango}')"
        h = get_headers_graph()
        h["Content-Type"] = "application/json"
        body = {"values": [array_valores]}
        r = requests.patch(url, headers=h, json=body)
        print(f"ESCRIBIR {hoja} FILA {fila_num}: {r.status_code} {r.text[:800]}")
        return r.status_code in [200,201]
    except Exception as e:
        print(f"ERROR escribir: {e}")
        return False

@app.route('/webhook', methods=['GET','POST'])
def webhook():
    if request.method == 'GET':
        if request.args.get('hub.verify_token') == VERIFY_TOKEN:
            return request.args.get('hub.challenge')
        return "Error", 403

    data = request.json
    print(f"WEBHOOK: {json.dumps(data)[:3000]}")
    try:
        entry = data['entry'][0]['changes'][0]['value']
        if 'messages' not in entry: return "ok", 200
        msg = entry['messages'][0]
        de = msg['from']
        texto = msg.get('text', {}).get('body', '').strip()
        texto_up = texto.upper()
        btn_id = msg.get('interactive', {}).get('button_reply', {}).get('id', '')

        if 'INVENTARIO' in texto_up or 'HOLA' in texto_up or btn_id == 'menu_principal':
            sesiones[de] = {"estado": "menu_principal"}
            enviar_botones(de, "🤖 *AUTOPIC INVENTARIO*\nElige:", [
                {"id":"opt_buscar","title":"🔍 BUSCAR"},
                {"id":"opt_agregar","title":"➕ AGREGAR"},
                {"id":"opt_folios","title":"📄 FOLIOS"}
            ])
        elif btn_id == 'opt_buscar':
            sesiones[de] = {"estado":"esperando_hoja_buscar"}
            enviar_botones(de, "🔍 ¿En qué hoja busco?", [
                {"id":"hoja_GENERAL","title":"GENERAL"},
                {"id":"hoja_DELTA","title":"DELTA"},
                {"id":"hoja_PHOENIX CONTACT","title":"PHOENIX"}
            ])
            enviar_texto(de, "Más: Escribe CARLO para CARLO GAVAZZI")
        elif btn_id == 'opt_agregar':
            sesiones[de] = {"estado":"esperando_hoja_agregar"}
            enviar_botones(de, "➕ ¿En qué hoja agrego?", [
                {"id":"agregar_GENERAL","title":"GENERAL"},
                {"id":"agregar_DELTA","title":"DELTA"},
                {"id":"agregar_PHOENIX CONTACT","title":"PHOENIX"}
            ])
        elif btn_id == 'opt_folios':
            enviar_botones(de, "📄 *FOLIOS VENDIDOS*", [
                {"id":"folios_buscar","title":"🔍 BUSCAR"},
                {"id":"folios_agregar","title":"➕ AGREGAR"}
            ])
        elif btn_id.startswith('hoja_') or texto_up == 'CARLO':
            hoja = btn_id.replace('hoja_','') if btn_id.startswith('hoja_') else 'CARLO GAVAZZI'
            sesiones[de] = {"estado":"esperando_modelo_buscar", "hoja": hoja}
            enviar_texto(de, f"📝 En *{hoja}*\nINGRESA MODELO:")
        elif sesiones.get(de,{}).get('estado') == 'esperando_modelo_buscar':
            hoja = sesiones[de]['hoja']
            enviar_texto(de, f"⏳ Buscando *{texto}* en {hoja}...")
            res = buscar_producto(hoja, texto)
            if res:
                detalle = f"✅ *ENCONTRADO EN {hoja} FILA {res['fila']}*\n\n"
                for i, h in enumerate(res['encabezados']):
                    if i < len(res['datos']):
                        detalle += f"*{h}:* {res['datos'][i]}\n"
                enviar_texto(de, detalle)
                sesiones[de] = {"estado":"preguntar_descontar", "hoja": hoja, "fila": res['fila'], "datos": res['datos']}
                enviar_botones(de, "¿DESCONTAR?", [{"id":"descontar_si","title":"SI"},{"id":"descontar_no","title":"NO"}])
            else:
                enviar_texto(de, f"❌ *{texto}* NO EXISTE EN {hoja}")
        elif btn_id == 'descontar_si':
            sesiones[de]['estado'] = 'esperando_cantidad_descontar'
            enviar_texto(de, "¿CUANTAS PIEZAS?")
        elif sesiones.get(de,{}).get('estado') == 'esperando_cantidad_descontar':
            try:
                cant = int(texto)
                hoja = sesiones[de]['hoja']
                fila = sesiones[de]['fila']
                cant_actual = int(float(str(sesiones[de]['datos'][3]).replace(',','')))
                if cant > cant_actual:
                    enviar_texto(de, f"❌ Stock insuficiente: {cant_actual}")
                else:
                    nueva = cant_actual - cant
                    encoded_path = ONEDRIVE_FILE_PATH.replace(" ", "%20")
                    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_path}:/workbook/worksheets/{hoja}/range(address='D{fila}')"
                    h = get_headers_graph(); h["Content-Type"]="application/json"
                    requests.patch(url, headers=h, json={"values":[[nueva]]})
                    enviar_texto(de, f"✅ Descontado. Nuevo stock {hoja}: {nueva}")
            except:
                enviar_texto(de, "Escribe solo número. Ej: 2")
        elif btn_id == 'descontar_no':
            enviar_botones(de, "Menú:", [{"id":"opt_buscar","title":"🔍 BUSCAR"},{"id":"opt_agregar","title":"➕ AGREGAR"},{"id":"opt_folios","title":"📄 FOLIOS"}])
        elif btn_id.startswith('agregar_'):
            hoja = btn_id.replace('agregar_','')
            sesiones[de] = {"estado":"agregar_paso", "hoja":hoja, "paso":1, "datos":{}}
            enviar_texto(de, f"📝 Agregando en *{hoja}* - Paso 1/9\nID:")
        elif sesiones.get(de,{}).get('estado') == 'agregar_paso':
            hoja = sesiones[de]['hoja']
            paso = sesiones[de]['paso']
            nombres = ["ID","MODELO","DESCRIPCION","CANTIDAD","UNIDAD","UBICACION","PROVEEDOR","NUMERO DE SERIE","MARCA"]
            sesiones[de]['datos'][paso] = texto
            if paso < 9:
                sesiones[de]['paso'] += 1
                enviar_texto(de, f"Paso {sesiones[de]['paso']}/9 {nombres[sesiones[de]['paso']-1]}:")
            else:
                d = sesiones[de]['datos']
                valores = get_excel(hoja)
                siguiente = len(valores) + 1
                fila_arr = [d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], "", d[9]]
                ok = escribir_fila(hoja, siguiente, fila_arr)
                enviar_texto(de, f"✅ AGREGADO EN {hoja} FILA {siguiente}" if ok else "❌ Error al escribir")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback; traceback.print_exc()
    return "ok", 200

@app.route('/')
def home():
    return f"AUTOPIC BOT V3 ACTIVO - PHONE_ID {PHONE_ID} - FILE {ONEDRIVE_FILE_PATH}"

if __name__ == '__main__':
    app.run(port=10000)
