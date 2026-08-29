from flask import Flask, request
import requests, os, json
from datetime import datetime

app = Flask(__name__)

# --- CONFIGURACIÓN - PON ESTO EN RENDER ENV ---
WHATSAPP_TOKEN = os.environ.get("WHATSAPP_TOKEN")
PHONE_ID = os.environ.get("PHONE_ID")
VERIFY_TOKEN = "autopic123"
MS_TOKEN = os.environ.get("MS_TOKEN")
DRIVE_ID = os.environ.get("DRIVE_ID")
FILE_ID = os.environ.get("FILE_ID")

sesiones = {} # memoria temporal {numero: {estado, hoja, datos, paso}}

def enviar_texto(to, body):
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    payload = {"messaging_product": "whatsapp", "to": to, "type": "text", "text": {"body": body}}
    requests.post(url, headers=headers, json=payload)

def enviar_botones(to, texto, botones):
    # botones = [{"id":"buscar","title":"🔍 BUSCAR"}]
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    btns = [{"type": "reply", "reply": {"id": b["id"], "title": b["title"]}} for b in botones]
    payload = {
        "messaging_product": "whatsapp", "to": to, "type": "interactive",
        "interactive": {"type": "button", "body": {"text": texto}, "action": {"buttons": btns}}
    }
    requests.post(url, headers=headers, json=payload)

def get_excel(hoja):
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{FILE_ID}/workbook/worksheets/{hoja}/usedRange"
    h = {"Authorization": f"Bearer {MS_TOKEN}"}
    r = requests.get(url, headers=h)
    return r.json().get('values', [])

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
    # array_valores debe ser lista de 10-11 columnas
    col_fin = chr(64 + len(array_valores)) # A=1
    rango = f"A{fila_num}:{col_fin}{fila_num}"
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{FILE_ID}/workbook/worksheets/{hoja}/range(address='{rango}')"
    h = {"Authorization": f"Bearer {MS_TOKEN}", "Content-Type": "application/json"}
    body = {"values": [array_valores]}
    r = requests.patch(url, headers=h, json=body)
    return r.status_code in [200,201]

@app.route('/webhook', methods=['GET','POST'])
def webhook():
    if request.method == 'GET':
        if request.args.get('hub.verify_token') == VERIFY_TOKEN:
            return request.args.get('hub.challenge')
        return "Error", 403

    data = request.json
    try:
        entry = data['entry'][0]['changes'][0]['value']
        if 'messages' not in entry: return "ok", 200
        msg = entry['messages'][0]
        de = msg['from']
        texto = msg.get('text', {}).get('body', '').strip()
        texto_up = texto.upper()
        btn_id = msg.get('interactive', {}).get('button_reply', {}).get('id', '')

        # MENU PRINCIPAL
        if 'INVENTARIO' in texto_up or btn_id == 'menu_principal':
            sesiones[de] = {"estado": "menu_principal"}
            enviar_botones(de, "🤖 *AUTOPIC INVENTARIO*\nElige:", [
                {"id":"opt_buscar","title":"🔍 BUSCAR"},
                {"id":"opt_agregar","title":"➕ AGREGAR"},
                {"id":"opt_folios","title":"📄 FOLIOS"}
            ])

        # OPCION BUSCAR
        elif btn_id == 'opt_buscar':
            sesiones[de] = {"estado":"esperando_hoja_buscar"}
            enviar_botones(de, "🔍 ¿En qué hoja busco?", [
                {"id":"hoja_GENERAL","title":"GENERAL"},
                {"id":"hoja_DELTA","title":"DELTA"},
                {"id":"hoja_PHOENIX CONTACT","title":"PHOENIX"}
            ])
            enviar_texto(de, "Más: Escribe CARLO para CARLO GAVAZZI")

        # OPCION AGREGAR
        elif btn_id == 'opt_agregar':
            sesiones[de] = {"estado":"esperando_hoja_agregar"}
            enviar_botones(de, "➕ ¿En qué hoja agrego?", [
                {"id":"agregar_GENERAL","title":"GENERAL"},
                {"id":"agregar_DELTA","title":"DELTA"},
                {"id":"agregar_PHOENIX CONTACT","title":"PHOENIX"}
            ])

        # OPCION FOLIOS VENDIDOS
        elif btn_id == 'opt_folios':
            enviar_botones(de, "📄 *FOLIOS VENDIDOS*", [
                {"id":"folios_buscar","title":"🔍 BUSCAR"},
                {"id":"folios_agregar","title":"➕ AGREGAR"}
            ])

        # --- BUSCAR HOJA SELECCIONADA ---
        elif btn_id.startswith('hoja_') or texto_up == 'CARLO':
            hoja = btn_id.replace('hoja_','') if btn_id.startswith('hoja_') else 'CARLO GAVAZZI'
            sesiones[de] = {"estado":"esperando_modelo_buscar", "hoja": hoja}
            enviar_texto(de, f"📝 En *{hoja}*\n\nINGRESA MODELO:")

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
                enviar_botones(de, "¿DESCONTAR PRODUCTO?", [{"id":"descontar_si","title":"SI"},{"id":"descontar_no","title":"NO"}])
            else:
                enviar_texto(de, f"❌ EL PRODUCTO *{texto}* NO EXISTE EN {hoja}")
                sesiones[de] = {"estado":"menu_principal"}

        # DESCONTAR
        elif btn_id == 'descontar_si':
            sesiones[de]['estado'] = 'esperando_cantidad_descontar'
            enviar_texto(de, "¿CUANTAS PIEZAS?")

        elif sesiones.get(de,{}).get('estado') == 'esperando_cantidad_descontar':
            try:
                cant_descontar = int(texto)
                hoja = sesiones[de]['hoja']
                fila = sesiones[de]['fila']
                # columna D es cantidad (indice 3)
                cant_actual = int(sesiones[de]['datos'][3])
                if cant_descontar > cant_actual:
                    enviar_texto(de, f"❌ NO HAY SUFICIENTE STOCK. Stock actual: {cant_actual}")
                else:
                    nueva = cant_actual - cant_descontar
                    # actualizar solo columna D
                    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{FILE_ID}/workbook/worksheets/{hoja}/range(address='D{fila}')"
                    h = {"Authorization": f"Bearer {MS_TOKEN}", "Content-Type": "application/json"}
                    requests.patch(url, headers=h, json={"values":[[nueva]]})
                    enviar_texto(de, f"✅ DESCONTADO. Nueva cantidad en {hoja}: {nueva}")
                sesiones[de] = {"estado":"menu_principal"}
            except:
                enviar_texto(de, "Escribe solo el número. Ej: 2")

        elif btn_id == 'descontar_no':
            enviar_botones(de, "Menú:", [{"id":"opt_buscar","title":"🔍 BUSCAR"},{"id":"opt_agregar","title":"➕ AGREGAR"},{"id":"opt_folios","title":"📄 FOLIOS"}])

        # --- AGREGAR INVENTARIO POR PASOS ---
        elif btn_id.startswith('agregar_'):
            hoja = btn_id.replace('agregar_','')
            sesiones[de] = {"estado":"agregar_paso", "hoja":hoja, "paso":1, "datos":{}}
            enviar_texto(de, f"📝 Agregando en *{hoja}* - Paso 1/9\nID (ej. P010):")

        elif sesiones.get(de,{}).get('estado') == 'agregar_paso':
            hoja = sesiones[de]['hoja']
            paso = sesiones[de]['paso']
            pasos_nombres = ["ID","MODELO","DESCRIPCION","CANTIDAD","UNIDAD","UBICACION","PROVEEDOR","NUMERO DE SERIE","MARCA"]
            sesiones[de]['datos'][paso] = texto
            if paso < 9:
                sesiones[de]['paso'] += 1
                enviar_texto(de, f"Paso {sesiones[de]['paso']}/9\n{pasos_nombres[sesiones[de]['paso']-1]}:")
            else:
                # ya tiene todo
                d = sesiones[de]['datos']
                valores = get_excel(hoja)
                siguiente = len(valores) + 1
                # A ID, B MODELO, C DESC, D CANT, E UNID, F UBIC, G PROV, H SERIE, I VACIA, J MARCA
                fila_arr = [d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], "", d[9]]
                ok = escribir_fila(hoja, siguiente, fila_arr)
                if ok:
                    enviar_texto(de, f"✅ PRODUCTO AGREGADO EN {hoja} FILA {siguiente}: {d[1]} - {d[2]}")
                else:
                    enviar_texto(de, "❌ Error al escribir")
                sesiones[de] = {"estado":"menu_principal"}

    except Exception as e:
        print(f"ERROR: {e}")
    return "ok", 200

@app.route('/')
def home():
    return "AUTOPIC BOT V3 ACTIVO"

if __name__ == '__main__':
    app.run(port=10000)
