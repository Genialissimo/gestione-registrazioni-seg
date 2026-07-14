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
RIGA_INTESTAZIONE_RISPOSTE = 9

NOME_FOGLIO_ANAGRAFICA = "Anagrafica"
RIGA_INTESTAZIONE_ANAGRAFICA = 1

# Opzioni fisse per i campi a scelta della scheda anagrafica, basate sulla
# cartolina di registrazione del proclamatore (modulo S-21).
OPZIONI_SESSO = ["Maschio", "Femmina"]
OPZIONI_INCARICO = ["(nessuno)", "Anziano", "Servitore di ministero"]
OPZIONI_TIPO = ["Proclamatore", "Pioniere Regolare", "Pioniere speciale", "Missionario sul campo"]
OPZIONI_ATTIVI_INATTIVI = ["A", "I"]
ETICHETTE_ATTIVI_INATTIVI = {"A": "Attivo", "I": "Inattivo"}


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
def leggi_foglio_come_df(_workbook, nome_foglio: str, riga_intestazione: int = 1):
    """Legge un foglio (tab) del workbook e lo ritorna come DataFrame.
    'riga_intestazione' è il numero di riga (1-based) in cui si trovano i
    nomi delle colonne: tutto ciò che sta sopra viene ignorato, tutto ciò
    che sta sotto diventa dati. Ritorna (dataframe, errore). Il parametro
    workbook è preceduto da '_' per dire a Streamlit di non provare a
    metterlo in cache lui stesso (gli oggetti gspread non sono 'hashable')."""
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

    tutti_i_valori = ws.get_all_values()
    if len(tutti_i_valori) < riga_intestazione:
        return pd.DataFrame(), None

    intestazioni = tutti_i_valori[riga_intestazione - 1]
    righe_dati = tutti_i_valori[riga_intestazione:]

    # Rende univoci eventuali nomi di colonna duplicati o vuoti, che
    # altrimenti causerebbero un errore nella creazione del DataFrame.
    intestazioni_pulite = []
    contatori = {}
    for i, nome in enumerate(intestazioni):
        nome = nome.strip() or f"Colonna {i + 1}"
        if nome in contatori:
            contatori[nome] += 1
            nome = f"{nome} ({contatori[nome]})"
        else:
            contatori[nome] = 0
        intestazioni_pulite.append(nome)

    # Scarta le righe completamente vuote (capita spesso in fondo al foglio).
    righe_dati = [r for r in righe_dati if any(cella.strip() for cella in r)]

    df = pd.DataFrame(righe_dati, columns=intestazioni_pulite)
    return df, None


def trova_indice_colonna(intestazioni: list, parola_chiave: str):
    """Cerca (senza distinguere maiuscole/minuscole) la prima colonna la cui
    intestazione contiene 'parola_chiave'. Ritorna l'indice (0-based) o None
    se non trovata — usato per non dipendere da un testo esatto al carattere."""
    parola_chiave = parola_chiave.lower()
    for i, nome in enumerate(intestazioni):
        if parola_chiave in nome.lower():
            return i
    return None


def calcola_eta(data_str: str) -> str:
    """Calcola gli anni compiuti da una data in formato gg/mm/aaaa. Ritorna
    stringa vuota se la data non è valida."""
    if not data_str:
        return ""
    try:
        d = datetime.strptime(data_str, "%d/%m/%Y")
    except ValueError:
        return ""
    oggi = datetime.now()
    anni = oggi.year - d.year - ((oggi.month, oggi.day) < (d.month, d.day))
    return str(anni)


def calcola_eta_dettagliata(data_str: str) -> str:
    """Calcola anni e mesi compiuti da una data in formato gg/mm/aaaa, nel
    formato 'anni,mesi' (es. '51,7' = 51 anni e 7 mesi). Ritorna stringa
    vuota se la data non è valida."""
    if not data_str:
        return ""
    try:
        d = datetime.strptime(data_str, "%d/%m/%Y")
    except ValueError:
        return ""
    oggi = datetime.now()
    anni = oggi.year - d.year
    mesi = oggi.month - d.month
    if oggi.day < d.day:
        mesi -= 1
    if mesi < 0:
        anni -= 1
        mesi += 12
    return f"{anni},{mesi}"


def opzioni_da_colonna(df: pd.DataFrame, nome_colonna: str) -> list:
    """Ritorna i valori unici (non vuoti) già presenti in una colonna del
    DataFrame, ordinati alfabeticamente — usati per popolare menu a tendina
    con le voci realmente in uso nel foglio (es. i nomi dei Gruppi)."""
    if nome_colonna not in df.columns:
        return []
    valori = df[nome_colonna].astype(str).str.strip()
    valori = sorted({v for v in valori if v and v.lower() != "nan"})
    return valori


def prossimo_id_anagrafica(df: pd.DataFrame) -> int:
    if "ID" not in df.columns or df.empty:
        return 1
    numeri = pd.to_numeric(df["ID"], errors="coerce").dropna()
    return int(numeri.max()) + 1 if not numeri.empty else 1


def salva_riga_anagrafica(_workbook, valori: dict, riga_da_aggiornare: int = None):
    """Scrive una nuova riga in fondo al foglio Anagrafica, oppure aggiorna
    una riga esistente (numero di riga del foglio, 1-based, intestazione
    inclusa) se 'riga_da_aggiornare' è specificato. 'valori' è un dizionario
    {nome_colonna: valore}; le colonne del foglio non presenti nel
    dizionario vengono lasciate vuote (nuova riga) o non toccate (modifica
    parziale non supportata qui: si riscrive l'intera riga in ordine).
    Ritorna (successo: bool, errore: str|None)."""
    try:
        ws = _workbook.worksheet(NOME_FOGLIO_ANAGRAFICA)
        intestazioni = ws.row_values(RIGA_INTESTAZIONE_ANAGRAFICA)
        riga_completa = [valori.get(nome, "") for nome in intestazioni]

        if riga_da_aggiornare is None:
            ws.append_row(riga_completa, value_input_option="USER_ENTERED")
        else:
            ultima_colonna = gspread.utils.rowcol_to_a1(1, len(intestazioni)).rstrip("0123456789")
            intervallo = f"A{riga_da_aggiornare}:{ultima_colonna}{riga_da_aggiornare}"
            ws.update(intervallo, [riga_completa], value_input_option="USER_ENTERED")
        return True, None
    except Exception as e:
        return False, f"Errore durante il salvataggio: {e}"


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
    st.button("🗂️  Anagrafiche", disabled=not collegato, use_container_width=True,
              on_click=vai_a, args=("anagrafiche",))


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
    c1, c2 = st.columns(2)
    card_data = [
        ("💳", "Registrazioni", "Registrazione movimenti (in arrivo)."),
        ("📋", "Report", "Schede e riepiloghi (in arrivo)."),
    ]
    for col, (icon, titolo, desc) in zip((c1, c2), card_data):
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
    st.title("Rapporti consegnati")
    st.caption(f"Dati letti dal foglio «{NOME_FOGLIO_RISPOSTE}» (intestazione riga {RIGA_INTESTAZIONE_RISPOSTE}).")

    if not collegato:
        st.warning("⚠️  Nessun foglio dati collegato.")
        return

    df, err = leggi_foglio_come_df(workbook, NOME_FOGLIO_RISPOSTE, RIGA_INTESTAZIONE_RISPOSTE)

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
# PAGINA: ANAGRAFICHE
# ─────────────────────────────────────────────────────────────────
def _form_anagrafica(df: pd.DataFrame, riga_esistente: dict = None, numero_riga_foglio: int = None):
    """Disegna il form di inserimento/modifica di un Proclamatore.
    Se 'riga_esistente' è None è un inserimento nuovo, altrimenti è una
    modifica (i campi vengono precompilati con i valori attuali)."""
    e = riga_esistente or {}

    def parse_data(s):
        try:
            return datetime.strptime(s, "%d/%m/%Y").date()
        except Exception:
            return None

    with st.form("form_anagrafica", clear_on_submit=False):
        nome_cognome = st.text_input("Cognome e Nome *", value=e.get("Cognome e Nome", ""))

        eta_nascita = calcola_eta_dettagliata(e.get("Data Nascita", ""))
        if eta_nascita:
            st.markdown(f"**Data di nascita** &nbsp; "
                        f"<span style='color:#D32F2F'>(anni {eta_nascita})</span>",
                        unsafe_allow_html=True)
        else:
            st.markdown("**Data di nascita**")
        data_nascita = st.date_input("Data di nascita", value=parse_data(e.get("Data Nascita", "")),
                                      format="DD/MM/YYYY", min_value=datetime(1900, 1, 1),
                                      label_visibility="collapsed")

        sesso_corrente = e.get("Sesso", "")
        sesso_default = ("Maschio" if sesso_corrente.upper().startswith("M")
                          else "Femmina" if sesso_corrente.upper().startswith("F") else "Maschio")
        sesso = st.selectbox("Sesso", OPZIONI_SESSO, index=OPZIONI_SESSO.index(sesso_default))

        senza_battesimo = st.checkbox("Non ancora battezzato/a",
                                       value=not bool(e.get("Data Battesimo", "")))
        data_battesimo = None
        if not senza_battesimo:
            eta_battesimo = calcola_eta_dettagliata(e.get("Data Battesimo", ""))
            if eta_battesimo:
                st.markdown(f"**Data del battesimo** &nbsp; "
                            f"<span style='color:#D32F2F'>(anni {eta_battesimo})</span>",
                            unsafe_allow_html=True)
            else:
                st.markdown("**Data del battesimo**")
            data_battesimo = st.date_input("Data del battesimo", value=parse_data(e.get("Data Battesimo", "")),
                                            format="DD/MM/YYYY", min_value=datetime(1900, 1, 1),
                                            label_visibility="collapsed")

        incarico_corrente = e.get("Incarico", "") or "(nessuno)"
        if incarico_corrente not in OPZIONI_INCARICO:
            incarico_corrente = "(nessuno)"
        incarico = st.selectbox("Incarico", OPZIONI_INCARICO,
                                 index=OPZIONI_INCARICO.index(incarico_corrente))

        tipo_corrente = e.get("Tipo", "") or "Proclamatore"
        if tipo_corrente not in OPZIONI_TIPO:
            tipo_corrente = "Proclamatore"
        tipo = st.selectbox("Tipo di servizio", OPZIONI_TIPO,
                             index=OPZIONI_TIPO.index(tipo_corrente))
        pr_dal = None
        if tipo in ("Pioniere Regolare", "Pioniere speciale", "Missionario sul campo"):
            pr_dal = st.date_input(f"{tipo} dal", value=parse_data(e.get("PR dal", "")),
                                    format="DD/MM/YYYY", min_value=datetime(1900, 1, 1))

        opzioni_gruppo = opzioni_da_colonna(df, "Gruppo")
        gruppo_corrente = e.get("Gruppo", "")
        elenco_gruppo = opzioni_gruppo + ["➕ Nuovo…"]
        if gruppo_corrente and gruppo_corrente not in elenco_gruppo:
            elenco_gruppo = [gruppo_corrente] + elenco_gruppo
        scelta_gruppo = st.selectbox("Gruppo", elenco_gruppo or ["➕ Nuovo…"],
                                      index=(elenco_gruppo.index(gruppo_corrente)
                                             if gruppo_corrente in elenco_gruppo else 0))
        if scelta_gruppo == "➕ Nuovo…":
            scelta_gruppo = st.text_input("Nome del nuovo gruppo")

        note = st.text_area("Note", value=e.get("Note", ""), height=100)

        st.divider()
        st.caption("Promemoria regolarità (da aggiornare quando manca il rapporto mensile)")
        irregolare = st.checkbox("Irregolare", value=e.get("Irregolare", "").strip().upper() in ("X", "SI", "SÌ"))
        irregolare_mesi = st.number_input("Irregolare da mesi", min_value=0, max_value=36, step=1,
                                           value=int(e.get("Irregolare da Mesi", 0) or 0))
        attivi_inattivi_corrente = e.get("Attivi / Inattivi", "A") or "A"
        if attivi_inattivi_corrente not in OPZIONI_ATTIVI_INATTIVI:
            attivi_inattivi_corrente = "A"
        etichetta_stato = st.selectbox("Stato", list(ETICHETTE_ATTIVI_INATTIVI.values()),
                                        index=OPZIONI_ATTIVI_INATTIVI.index(attivi_inattivi_corrente))
        attivi_inattivi = {v: k for k, v in ETICHETTE_ATTIVI_INATTIVI.items()}[etichetta_stato]
        dal = st.date_input("Stato attuale dal", value=parse_data(e.get("Dal", "")),
                             format="DD/MM/YYYY", min_value=datetime(1900, 1, 1))

        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            invia = st.form_submit_button("✔ Salva", use_container_width=True, type="primary")
        with col_btn2:
            annulla = st.form_submit_button("Annulla", use_container_width=True)

    if annulla:
        st.session_state.anagrafica_modifica = None
        st.session_state.anagrafica_nuovo = False
        st.rerun()

    if invia:
        if not nome_cognome.strip():
            st.error("Il campo 'Cognome e Nome' è obbligatorio.")
            return

        data_nascita_str = data_nascita.strftime("%d/%m/%Y") if data_nascita else ""
        data_battesimo_str = data_battesimo.strftime("%d/%m/%Y") if data_battesimo else ""

        valori = {
            "ID": e.get("ID") or str(prossimo_id_anagrafica(df)),
            "Cognome e Nome": nome_cognome.strip(),
            "Data Nascita": data_nascita_str,
            "Data Battesimo": data_battesimo_str,
            "Sesso": "M" if sesso == "Maschio" else "F",
            "Incarico": "" if incarico == "(nessuno)" else incarico,
            "Tipo": tipo,
            "PR dal": pr_dal.strftime("%d/%m/%Y") if pr_dal else "",
            "Gruppo": scelta_gruppo,
            "Note": note.strip(),
            "Irregolare": "X" if irregolare else "",
            "Irregolare da Mesi": str(irregolare_mesi) if irregolare else "",
            "Attivi / Inattivi": attivi_inattivi,
            "Dal": dal.strftime("%d/%m/%Y") if dal else "",
            "Anni Età": calcola_eta(data_nascita_str),
            "Anni Batt": calcola_eta(data_battesimo_str),
        }

        ok, err = salva_riga_anagrafica(workbook, valori, riga_da_aggiornare=numero_riga_foglio)
        if ok:
            st.cache_data.clear()
            st.session_state.anagrafica_modifica = None
            st.session_state.anagrafica_nuovo = False
            st.success("✔ Salvato correttamente.")
            st.rerun()
        else:
            st.error(err)


def mostra_anagrafiche():
    st.title("Anagrafiche")
    st.caption(f"Dati letti dal foglio «{NOME_FOGLIO_ANAGRAFICA}».")

    if not collegato:
        st.warning("⚠️  Nessun foglio dati collegato.")
        return

    if "anagrafica_nuovo" not in st.session_state:
        st.session_state.anagrafica_nuovo = False
    if "anagrafica_modifica" not in st.session_state:
        st.session_state.anagrafica_modifica = None

    df, err = leggi_foglio_come_df(workbook, NOME_FOGLIO_ANAGRAFICA, RIGA_INTESTAZIONE_ANAGRAFICA)
    if err:
        st.error(err)
        return

    # ── Form di inserimento nuovo Proclamatore ──────────────────────────
    if st.session_state.anagrafica_nuovo:
        st.subheader("➕ Nuovo Proclamatore")
        _form_anagrafica(df)
        return

    # ── Form di modifica Proclamatore esistente ─────────────────────────
    if st.session_state.anagrafica_modifica is not None:
        idx = st.session_state.anagrafica_modifica
        riga = df.iloc[idx].to_dict()
        numero_riga_foglio = RIGA_INTESTAZIONE_ANAGRAFICA + 1 + idx  # +1 per l'intestazione
        st.subheader(f"✏️ Modifica: {riga.get('Cognome e Nome', '')}")
        _form_anagrafica(df, riga_esistente=riga, numero_riga_foglio=numero_riga_foglio)
        return

    # ── Contenitore riservato per i pulsanti: verrà riempito più sotto, ma
    # essendo dichiarato qui appare sempre in cima alla pagina (utile su
    # telefono, dove altrimenti "Modifica" resterebbe nascosto sotto la
    # griglia finché non si scorre) ─────────────────────────────────────
    contenitore_azioni = st.container()

    ricerca = st.text_input("🔍 Cerca per nome, gruppo, tipo…", placeholder="Digita per filtrare…")

    if df.empty:
        with contenitore_azioni:
            if st.button("➕ Nuovo", use_container_width=True, type="primary"):
                st.session_state.anagrafica_nuovo = True
                st.rerun()
        st.info("Il foglio è collegato correttamente ma non contiene ancora Proclamatori.")
        return

    df_mostrato = df.reset_index(drop=True)
    if ricerca:
        maschera = df_mostrato.apply(
            lambda riga: riga.astype(str).str.contains(ricerca, case=False, na=False).any(), axis=1
        )
        df_mostrato = df_mostrato[maschera]

    st.caption(f"{len(df_mostrato)} Proclamatori su {len(df)} totali. Tocca una riga per selezionarla, poi «Modifica».")

    colonne_da_mostrare = [c for c in
                            ["Cognome e Nome", "Data Nascita", "Sesso", "Incarico", "Tipo",
                             "Gruppo", "Attivi / Inattivi", "Anni Età", "Anni Batt"]
                            if c in df_mostrato.columns]

    df_griglia = df_mostrato[colonne_da_mostrare].copy()
    if "Attivi / Inattivi" in df_griglia.columns:
        df_griglia["Attivi / Inattivi"] = df_griglia["Attivi / Inattivi"].map(
            lambda v: ETICHETTE_ATTIVI_INATTIVI.get(v, v)
        )
    df_griglia = df_griglia.rename(columns={
        "Cognome e Nome": "Nome",
        "Data Nascita": "Nascita",
        "Attivi / Inattivi": "Stato",
    })

    evento = st.dataframe(
        df_griglia,
        use_container_width=True,
        hide_index=True,
        height=480,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "Nome": st.column_config.TextColumn(width="medium"),
            "Nascita": st.column_config.TextColumn(width="small"),
            "Sesso": st.column_config.TextColumn(width="small"),
            "Incarico": st.column_config.TextColumn(width="medium"),
            "Tipo": st.column_config.TextColumn(width="medium"),
            "Gruppo": st.column_config.TextColumn(width="medium"),
            "Stato": st.column_config.TextColumn(width="small"),
        },
    )

    righe_selezionate = evento.selection.rows if evento and evento.selection else []
    idx_originale = None
    nome_sel = None
    if righe_selezionate:
        posizione = righe_selezionate[0]
        idx_originale = df_mostrato.index[posizione]
        nome_sel = df_mostrato.loc[idx_originale, "Cognome e Nome"]

    # ── Ora riempiamo il contenitore riservato in cima alla pagina ──────
    with contenitore_azioni:
        c_nuovo, c_modifica, c_aggiorna = st.columns(3)
        with c_nuovo:
            if st.button("➕ Nuovo", use_container_width=True, type="primary"):
                st.session_state.anagrafica_nuovo = True
                st.rerun()
        with c_modifica:
            if st.button("✏️ Modifica", use_container_width=True, disabled=idx_originale is None):
                st.session_state.anagrafica_modifica = idx_originale
                st.rerun()
        with c_aggiorna:
            if st.button("🔄", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
        if nome_sel:
            st.caption(f"Selezionato: **{nome_sel}**")


# ─────────────────────────────────────────────────────────────────
# ROUTING
# ─────────────────────────────────────────────────────────────────
if st.session_state.pagina == "registrazioni":
    mostra_registrazioni()
elif st.session_state.pagina == "anagrafiche":
    mostra_anagrafiche()
else:
    mostra_home()