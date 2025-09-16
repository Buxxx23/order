
# Wareneingangsbestellung Rotogal – Microsoft Cloud (OneDrive + Graph Mail)

Bereit für **Streamlit Community Cloud**. Speichert PDFs in **OneDrive** und versendet E‑Mails über **Microsoft Graph**.

## Deploy (Streamlit Cloud)
1. Lege ein GitHub‑Repo an und lade `app.py`, `requirements.txt`, `README.md` ins Repo.
2. https://share.streamlit.io → **New app** → Repository/Branch wählen → `app.py` als Main file → **Deploy**.

## Microsoft 365 vorbereiten
1) In **Microsoft Entra ID** (Azure AD) → **App registrations** → **New registration**.  
   - Notiere dir **Application (client) ID** und **Directory (tenant) ID**.
2) **API permissions → Microsoft Graph → Application permissions** hinzufügen:  
   - `Files.ReadWrite.All`  
   - `Mail.Send`  
   → **Grant admin consent** klicken.
3) **Certificates & secrets → New client secret** → Secret‑Wert sicher speichern.

## App konfigurieren
- In der laufenden App (Sidebar) oder in **Streamlit Secrets** hinterlegen:
```
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...
GRAPH_USER_UPN=name@firma.de
ONEDRIVE_FOLDER=Bestellungen/Rotogal
EMAIL_TO=einkauf@firma.de
```
- Aktivieren: **Auto‑upload PDF to OneDrive** und/oder **Auto‑send email**.

## Lokal starten (optional)
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Hinweise
- Upload via: `https://graph.microsoft.com/v1.0/users/{UPN}/drive/root:/<Folder>/<Filename>:/content`
- Mail via: `https://graph.microsoft.com/v1.0/users/{UPN}/sendMail`
- Die App nutzt den **Client‑Credentials‑Flow** (App‑Only). Der Upload/Mailversand erfolgt im Kontext des angegebenen **User UPN**.
