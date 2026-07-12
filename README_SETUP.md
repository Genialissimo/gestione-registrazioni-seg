# Configurazione iniziale — Gestione Registrazioni SEG

Questa guida va seguita **una sola volta** per collegare l'app al tuo
Google Sheet. Fammi sapere ad ogni passaggio se qualcosa non torna: la
facciamo insieme.

## 1. Crea un progetto Google Cloud (gratuito)

1. Vai su https://console.cloud.google.com/
2. Crea un nuovo progetto (es. "gestione-seg")
3. Nel menu "API e servizi" → "Libreria", attiva:
   - **Google Sheets API**
   - **Google Drive API**

## 2. Crea l'account di servizio (l'account "robot" che userà l'app)

Nota: questo è diverso dal tuo account Google normale — è un account
tecnico che permette al programma (non a te) di leggere/scrivere sul
foglio, senza bisogno di una password tua.

1. In "API e servizi" → "Credenziali" → "Crea credenziali" → "Account di servizio"
2. Dagli un nome (es. "seg-app")
3. Una volta creato, apri l'account di servizio → scheda "Chiavi" → "Aggiungi chiave" → "Crea nuova chiave" → formato **JSON**
4. Si scarica un file `.json`: tienilo da parte, servirà al punto 4 (NON va mai messo su GitHub, contiene una password)
5. Copia anche l'indirizzo email dell'account di servizio (tipo `seg-app@nome-progetto.iam.gserviceaccount.com`) — ti servirà al punto 3

## 3. Condividi il tuo foglio Google già pronto

1. Apri il tuo Google Sheet
2. Bottone "Condividi" → incolla l'email dell'account di servizio del punto 2 → permesso **Editor**
3. Copia l'ID del foglio dall'URL:
   `https://docs.google.com/spreadsheets/d/`**`QUESTO-È-L-ID`**`/edit`
   (ti servirà al punto 4)

## 4. Configura i "secrets" dell'app

Se pubblichi su **Streamlit Community Cloud**: vai nelle impostazioni
dell'app → "Secrets" e incolla (adattando ai tuoi valori):

```toml
sheet_id = "L_ID_DEL_TUO_FOGLIO_DEL_PUNTO_3"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "seg-app@nome-progetto.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

Tutti questi valori si trovano dentro il file `.json` scaricato al punto 2
(basta copiarli negli stessi campi). Se lavori in locale invece che su
Streamlit Cloud, lo stesso contenuto va in un file `.streamlit/secrets.toml`
nella cartella del progetto (mai da caricare su GitHub).

## 5. Avvio

Una volta pubblicata l'app (GitHub + Streamlit Community Cloud) e
configurati i secrets, apri il link: la sidebar mostrerà subito "Collegato"
con il nome del tuo foglio.

---

Quando sei pronto, dimmi a che punto sei arrivato (es. "ho creato l'account
di servizio, ora?") e ti guido nel passaggio successivo.
