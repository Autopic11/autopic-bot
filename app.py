from flask import Flask, request
import requests, os

app = Flask(__name__)
TOKEN = os.environ.get("WHATSAPP_TOKEN","").strip()
PHONE_ID = os.environ.get("PHONE_ID","1265790226622285").strip()
VERIFY = os.environ.get("VERIFY_TOKEN","autopic123").strip()

def send(to, txt):
    url=f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    h={"Authorization":f"Bearer {TOKEN}","Content-Type":"application/json"}
    d={"messaging_product":"whatsapp","to":to,"type":"text","text":{"body":txt}}
    r=requests.post(url, headers=h, json=d)
    print(f"SEND {to} -> {r.status_code} {r.text[:300]}")

@app.route('/webhook', methods=['GET','POST'])
def webhook():
    if request.method=='GET':
        if request.args.get('hub.verify_token')==VERIFY:
            return request.args.get('hub.challenge')
        return "fail",403
    data=request.json
    print(f"POST WEBHOOK RECIBIDO: {data}")
    try:
        val=data['entry'][0]['changes'][0]['value']
        if 'messages' in val:
            m=val['messages'][0]
            de=m['from']
            txt=m.get('text',{}).get('body','')
            print(f"Mensaje de {de}: {txt}")
            send(de, f"✅ Recibí tu: {txt}\nBot ACTIVO - ahora sí contesta!")
    except Exception as e:
        print(f"ERROR: {e}")
    return "ok",200

@app.route('/')
def home():
    return f"AUTOPIC BOT ACTIVO - PHONE_ID {PHONE_ID} - TOKEN OK? {bool(TOKEN)}"

@app.route('/test')
def test():
    n=request.args.get('numero')
    send(n,"✅ PRUEBA MINI BOT")
    return f"enviado a {n}"

if __name__=='__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT",10000)))
