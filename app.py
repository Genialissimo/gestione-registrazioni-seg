"""
app.py
Gestione Registrazioni SEG - Web App (Streamlit + Google Sheets)
Collegamento a un unico foglio Google + Home + Visualizza registrazioni.
Altri form (Nuova registrazione, Anagrafiche) verranno aggiunti in seguito.
"""

from datetime import datetime

import pandas as pd
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

NOME_FOGLIO_RISPOSTE = "Risposte del modulo 9"


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


@st.cache_data(ttl=60, show_spinner=False)
def leggi_foglio_come_df(_workbook, nome_foglio: str):
    """Legge un foglio (tab) del workbook e lo ritorna come DataFrame.
    Ritorna (dataframe, errore). Il parametro workbook è preceduto da '_'
    per dire a Streamlit di non provare a metterlo in cache lui stesso
    (gli oggetti gspread non sono 'hashable')."""
    try:
        ws = _workbook.worksheet(nome_foglio)
    except gspread.WorksheetNotFound:
        nomi_disponibili = ", ".join(f"'{f.title}'" for f in _workbook.worksheets())
        return None, (
            f"Il foglio '{nome_foglio}' non esiste nel documento collegato. "
            f"Fogli disponibili: {nomi_disponibili}."
        )
    except Exception as e:
        return None, f"Errore durante la lettura del foglio: {e}"

    valori = ws.get_all_records()
    if not valori:
        return pd.DataFrame(), None
    return pd.DataFrame(valori), None


# ─────────────────────────────────────────────────────────────────
# STATO SESSIONE (navigazione tra pagine)
# ─────────────────────────────────────────────────────────────────
if "pagina" not in st.session_state:
    st.session_state.pagina = "home"


def vai_a(pagina: str):
    st.session_state.pagina = pagina


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
    st.button("🏠  Home", use_container_width=True,
              on_click=vai_a, args=("home",))
    st.button("📝  Nuova registrazione", disabled=not collegato,
              use_container_width=True)
    st.button("📖  Visualizza registrazioni", disabled=not collegato,
              use_container_width=True, on_click=vai_a, args=("registrazioni",))
    st.button("🗂️  Anagrafiche", disabled=not collegato, use_container_width=True)


# ─────────────────────────────────────────────────────────────────
# PAGINA: HOME
# ─────────────────────────────────────────────────────────────────
def mostra_home():
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


# ─────────────────────────────────────────────────────────────────
# PAGINA: VISUALIZZA REGISTRAZIONI
# ─────────────────────────────────────────────────────────────────
def mostra_registrazioni():
    st.title("Visualizza registrazioni")
    st.caption(f"Dati letti dal foglio «{NOME_FOGLIO_RISPOSTE}».")

    if not collegato:
        st.warning("⚠️  Nessun foglio dati collegato.")
        return

    df, err = leggi_foglio_come_df(workbook, NOME_FOGLIO_RISPOSTE)

    if err:
        st.error(err)
        return

    if df.empty:
        st.info("Il foglio è collegato correttamente ma non contiene ancora righe di dati.")
        return

    fr_top = st.columns([3, 1, 1])
    with fr_top[0]:
        ricerca = st.text_input("🔍 Cerca in tutte le colonne",
                                 placeholder="Digita per filtrare le righe…")
    with fr_top[1]:
        st.metric("Righe totali", len(df))
    with fr_top[2]:
        if st.button("🔄 Aggiorna", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    df_mostrato = df
    if ricerca:
        maschera = df.apply(
            lambda riga: riga.astype(str).str.contains(ricerca, case=False, na=False).any(),
            axis=1,
        )
        df_mostrato = df[maschera]
        st.caption(f"{len(df_mostrato)} righe corrispondenti su {len(df)} totali.")

    st.dataframe(
        df_mostrato,
        use_container_width=True,
        hide_index=True,
        height=560,
    )


# ─────────────────────────────────────────────────────────────────
# ROUTING
# ─────────────────────────────────────────────────────────────────
if st.session_state.pagina == "registrazioni":
    mostra_registrazioni()
else:
    mostra_home()
