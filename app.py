"""
app.py
Gestione Registrazioni SEG - Web App (Streamlit + Google Sheets)
Scheletro iniziale: collegamento a un unico foglio Google + Home.
I form di registrazione/anagrafiche verranno aggiunti in seguito.
"""

from datetime import datetime

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────────────────────────
# CONFIGURAZIONE PAGINA
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestione Registrazioni SEG",
    page_icon="📒",
    layout="wide",
    initial_sidebar_state="expanded",
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ─────────────────────────────────────────────────────────────────
# CONNESSIONE A GOOGLE
# ─────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_client() -> gspread.Client:
    """Autentica il programma verso Google tramite l'account di servizio
    definito nei 'secrets' dell'app (vedi README_SETUP.md)."""
    credenziali = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(credenziali)


@st.cache_resource(show_spinner=False)
def apri_foglio_dati():
    """Apre il foglio Google dati (l'ID è definito una volta sola nei
    secrets, in 'sheet_id'). Ritorna (workbook, errore)."""
    try:
        client = get_client()
        wb = client.open_by_key(st.secrets["sheet_id"])
        return wb, None
    except gspread.exceptions.APIError:
        email_sa = st.secrets["gcp_service_account"]["client_email"]
        return None, (
            "Impossibile aprire il foglio dati. Controlla che sia stato "
            f"condiviso (come Editor) con:\n`{email_sa}`"
        )
    except Exception as e:
        return None, f"Errore durante il collegamento: {e}"


# ─────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────
workbook, errore = apri_foglio_dati()
collegato = workbook is not None

with st.sidebar:
    st.markdown("## 📒 Gestione Registrazioni SEG")
    st.divider()

    if collegato:
        st.success(f"✅  Collegato: {workbook.title}")
    else:
        st.error("⚠️  Non collegato")
        if errore:
            st.caption(errore)

    st.divider()
    st.subheader("Registrazioni")
    st.caption("Le voci qui sotto si attivano dopo il collegamento; le costruiremo insieme.")
    st.button("📝  Nuova registrazione", disabled=not collegato, use_container_width=True)
    st.button("📖  Visualizza registrazioni", disabled=not collegato, use_container_width=True)
    st.button("🗂️  Anagrafiche", disabled=not collegato, use_container_width=True)


# ─────────────────────────────────────────────────────────────────
# HOME
# ─────────────────────────────────────────────────────────────────
st.title("Pannello di controllo")

if collegato:
    col1, col2 = st.columns([3, 1])
    with col1:
        st.success(f"✅  Foglio collegato: **{workbook.title}**")
    with col2:
        st.link_button("Apri su Google Sheets", workbook.url, use_container_width=True)
else:
    st.warning("⚠️  Nessun foglio dati collegato. Controlla la configurazione (vedi README_SETUP.md).")

st.divider()

st.subheader("Sezioni")
c1, c2, c3 = st.columns(3)
card_data = [
    ("👤", "Anagrafiche", "Gestione delle anagrafiche (in arrivo)."),
    ("💳", "Registrazioni", "Registrazione movimenti (in arrivo)."),
    ("📋", "Report", "Schede e riepiloghi (in arrivo)."),
]
for col, (icon, titolo, desc) in zip((c1, c2, c3), card_data):
    with col:
        with st.container(border=True):
            st.markdown(f"#### {icon}  {titolo}")
            st.caption(desc)
            st.button("Apri →", key=f"card_{titolo}",
                      disabled=not collegato, use_container_width=True)

st.caption(f"Ultimo aggiornamento pagina: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
