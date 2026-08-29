from flask import Flask, request
import requests, os, json
from datetime import datetime
import time

app = Flask(__name__)

# --- CORREGIDO: Lee todas las variantes ---
WHATSAPP_TOKEN = os.environ.get("WHATSAPP_TOKEN", "").strip()
PHONE_ID = (os.environ.get("PHONE_ID") or os.environ.get("PHONE_NUMBER_ID") or "1265790226622285").strip()
VERIFY_TOKEN = (os.environ.get("VERIFY_TOKEN") or "autopic123").strip()

MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID", "").strip()
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "").strip()
MS_REFRESH_TOKEN = os.environ.get("MS_REFRESH_TOKEN", "").strip()
ONEDRIVE_FILE_PATH = os.environ.get("ONEDRIVE_FILE_PATH") or "AUTOPIC/DOCUMENTACIÓN/INVENTARIO GENERAL PRUEBA.xlsx"

_ms_token_cache = {"token": None, "expiry": 0}

def get_ms_access_token():
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
    try:
        r = requests.post(url, data=data, timeout=20)
        print(f"RENOVAR MS TOKEN: {r.status_code}")
        if r.status_code == 200:
            j = r.json()
            _ms_token_cache["token"] = j["access_token"]
            _ms_token_cache["expiry"] = time.time() + j.get("expires_in", 3600)
            return j["access_token"]
        else:
            print(r.text[:2000])
            return None
    except Exception as e:
        print(f"ERROR MS TOKEN: {e}")
        return None

def get_headers_graph():
    token = get_ms_access_token()
    return {"Authorization": f"Bearer {token}"} if token else {}

sesiones = {}

def enviar_texto(to, body):
    if not WHATSAPP_TOKEN or not PHONE_ID:
        print(f"❌ FALTA TOKEN O PHONE_ID: TOKEN existe? {bool(WHATSAPP_TOKEN)} PHONE_ID={PHONE_ID}")
        return None
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
        encoded_path = ONEDRIVE_FILE_PATH.replace(" ", "%20")
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_path}:/workbook/worksheets/{hoja}/usedRange"
        h = get_headers_graph()
        r = requests.get(url, headers=h, timeout=20)
        print(f"GET EXCEL {hoja}: {r.status_code}")
        if r.status_code!= 200:
            print(r.text[:1500])
            return []
        return r.json().get('values', [])
    except Exception as e:
        print(f"ERROR get_excel: {e}")
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
        r = requests.patch(url, headers=h, json=body, timeout=20)
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
        if 'messages' not in entry:
            return "ok", 200
        msg = entry['messages'][0]
        de = msg['from']
        texto = msg.get('text', {}).get('body', '').strip()
        texto_up = texto.upper()
        btn_id = msg.get('interactive', {}).get('button_reply', {}).get('id', '')

        # --- CORREGIDO: AHORA CONTESTA A HOLA, INVENTARIO, INVENT, O CUALQUIER COSA ---
        if any(x in texto_up for x in ['INVENTARIO', 'HOLA', 'MENU', 'INVENT', 'AYUDA']) or btn_id == 'menu_principal' or texto_up == '':
            # Si manda cualquier cosa sin sesion, lo mandamos al menu
            sesiones[de] = {"estado": "menu_principal"}
            enviar_botones(de, "🤖 *AUTOPIC INVENTARIO*\nElige una opción:", [
                {"id":"opt_buscar","title":"🔍 BUSCAR"},
                {"id":"opt_agregar","title":"➕ AGREGAR"},
                {"id":"opt_folios","title":"📄 FOLIOS"}
            ])
            return "ok", 200

        if btn_id == 'opt_buscar' or 'BUSCAR' in texto_up:
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
        elif btn_id.startswith('hoja_') or texto_up == 'CARLO' or texto_up.startswith('HOJA'):
            hoja = btn_id.replace('hoja_','') if btn_id.startswith('hoja_') else ('CARLO GAVAZZI' if texto_up == 'CARLO' else texto.replace('HOJA','').strip() or 'GENERAL')
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
                enviar_botones(de, "¿DESCONTAR STOCK?", [{"id":"descontar_si","title":"SI"},{"id":"descontar_no","title":"NO"}])
            else:
                enviar_texto(de, f"❌ *{texto}* NO EXISTE EN {hoja}\nEscribe otro modelo o escribe MENU")
        elif btn_id == 'descontar_si':
            sesiones[de]['estado'] = 'esperando_cantidad_descontar'
            enviar_texto(de, "¿CUANTAS PIEZAS? (solo número)")
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
                    requests.patch(url, headers=h, json={"values":[[nueva]]}, timeout=20)
                    enviar_texto(de, f"✅ Descontado. Nuevo stock {hoja}: {nueva}")
                    sesiones[de] = {"estado": "menu_principal"}
            except:
                enviar_texto(de, "Escribe solo número. Ej: 2")
        elif btn_id == 'descontar_no':
            enviar_botones(de, "Menú principal:", [{"id":"opt_buscar","title":"🔍 BUSCAR"},{"id":"opt_agregar","title":"➕ AGREGAR"},{"id":"opt_folios","title":"📄 FOLIOS"}])

    except Exception as e:
        print(f"ERROR WEBHOOK: {e}")
        import traceback; traceback.print_exc()
    return "ok", 200

@app.route('/')
def home():
    return f"AUTOPIC BOT ACTIVO - PHONE_ID {PHONE_ID} - FILE {ONEDRIVE_FILE_PATH} - TOKEN OK? {bool(WHATSAPP_TOKEN)}"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
