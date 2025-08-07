import sys
import requests

CLIENT_ID = 'QXGVf1neHIouY41kH8o0TVGAPIPbQebbulB371893AD6AMqE'
CLIENT_SECRET = 'eNAFwASrEs6mneWCY2GcVwIgmYhFnPX2RB3wbjvrWsp2ClwSx3gaW7DmsO9jpip4'

# Argumente prüfen
if len(sys.argv) != 3:
    print("❌ Erwartet: project_guid und model_guid als Argumente.")
    sys.exit(1)

project_guid = sys.argv[1]
model_guid = sys.argv[2]
project_id = 'b.' + project_guid  # Forge erwartet 'b.'-Prefix

def get_token():
    url = "https://developer.api.autodesk.com/authentication/v1/authenticate"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "data:read"
    }
    resp = requests.post(url, data=data)
    resp.raise_for_status()
    return resp.json()['access_token']

def get_item_info(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/items/{model_guid}"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()

try:
    token = get_token()
    item_info = get_item_info(token)
    size = item_info.get('data', {}).get('attributes', {}).get('storageSize')

    if size:
        size_mb = round(int(size) / (1024 * 1024), 2)
        print(f"📦 Modellgröße: {size_mb} MB")
    else:
        print("❗ Dateigröße nicht gefunden.")
except Exception as e:
    print(f"❌ Fehler: {str(e)}")
    sys.exit(1)
