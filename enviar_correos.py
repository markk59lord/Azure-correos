import msal
import requests
import pandas as pd

# Leer correos desde archivo Excel
excel_path = r"C:\Users\Inventario 2\Desktop\inventarios.xlsx"  # Ruta al archivo
df = pd.read_excel(excel_path)

# Validar que exista la columna 'correo'
if 'correo' not in df.columns:
    print("❌ La columna 'correo' no se encontró en el archivo Excel.")
    exit()

# Datos de tu aplicación registrada en Azure
client_id = '8ae73c31-e6bd-4206-8bd5-a1e3a06ae07a'
tenant_id = '4d1ba89f-9d3a-43fa-9e69-226f396d4086'

authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["Mail.Send"]

# Crear la aplicación pública
app = msal.PublicClientApplication(client_id=client_id, authority=authority)

# Obtener token
accounts = app.get_accounts()
if accounts:
    token_response = app.acquire_token_silent(scopes, account=accounts[0])
else:
    token_response = app.acquire_token_interactive(scopes)

if "access_token" not in token_response:
    print(f"❌ Error al obtener token: {token_response.get('error_description')}")
    exit()

access_token = token_response["access_token"]
print("✅ Token de acceso obtenido correctamente.")

# Enviar correos uno a uno
for email in df['correo'].dropna():
    print(f"📤 Enviando a: {email}")
    email_data = {
        "message": {
            "subject": "Correo de prueba individual",
            "body": {
                "contentType": "Text",
                "content": f"Hola {email}, este es un correo de prueba enviado individualmente."
            },
            "toRecipients": [
                {"emailAddress": {"address": email}}
            ]
        },
        "saveToSentItems": "true"
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    graph_url = "https://graph.microsoft.com/v1.0/me/sendMail"
    response = requests.post(graph_url, headers=headers, json=email_data)

    if response.status_code == 202:
        print("✅ Correo enviado.")
    else:
        print(f"❌ Error ({response.status_code}): {response.text}")
