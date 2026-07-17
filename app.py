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

NOME_FOGLIO_TUTTI = "Tutti"
RIGA_INTESTAZIONE_TUTTI = 4
# Indici di colonna (0-based) nel foglio 'Tutti': B=1, C=2, D=3, E=4, G=6, H=7, I=8, J=9
COL_TUTTI_NOME = 1
COL_TUTTI_MESE = 2
COL_TUTTI_TIPO_SERVIZIO = 3   # D: es. "Pioniere Ausiliario"
COL_TUTTI_MINISTERO = 4       # E: "Si"/"No" — ha partecipato al ministero
COL_TUTTI_ORE = 6
COL_TUTTI_CRED_ORE = 7
COL_TUTTI_STUDI = 8
COL_TUTTI_OSSERVAZIONI = 9

MESI_ITALIANI = {
    1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile", 5: "Maggio", 6: "Giugno",
    7: "Luglio", 8: "Agosto", 9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre",
}

# Opzioni fisse per i campi a scelta della scheda anagrafica, basate sulla
# cartolina di registrazione del proclamatore (modulo S-21).
OPZIONI_SESSO = ["Maschio", "Femmina"]
OPZIONI_INCARICO = ["(nessuno)", "Anziano", "Servitore di ministero"]
OPZIONI_TIPO = ["Proclamatore", "Pioniere Regolare", "Pioniere speciale", "Missionario sul campo"]
OPZIONI_HAI_SERVITO = ["Proclamatore", "Pioniere Ausiliario", "Pioniere Regolare",
                       "Pioniere Speciale", "Rappresentante sul campo"]
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


def a_float_it(s: str) -> float:
    """Converte una stringa numerica in formato italiano (es. '15,00' o
    '15') in float. Ritorna 0.0 se vuota o non valida."""
    s = (s or "").strip()
    if not s:
        return 0.0
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return 0.0


def formatta_numero_it(v: float) -> str:
    """Formatta un numero come stringa in stile italiano: 42,00"""
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formatta_mese_esteso(mese_anno: str) -> str:
    """Converte 'AAAA-MM' (es. '2025-09') nel nome del mese per esteso
    (es. 'Settembre'). Ritorna il valore originale se non riconosciuto."""
    try:
        _, mese = mese_anno.split("-")
        return MESI_ITALIANI.get(int(mese), mese_anno)
    except Exception:
        return mese_anno


@st.cache_data(ttl=60, show_spinner=False)
def leggi_foglio_tutti(_workbook):
    """Legge il foglio 'Tutti' (archivio storico di tutti i rapporti di
    tutti i mesi) leggendo le colonne per posizione (B, C, D, E, G, H, I, J),
    dato che l'intestazione non è a riga 1. Calcola inoltre le due caselle
    'Ha partecipato al ministero' e 'Pioniere ausiliario' secondo le regole
    indicate. Ritorna (dataframe, errore)."""
    try:
        ws = _workbook.worksheet(NOME_FOGLIO_TUTTI)
    except gspread.WorksheetNotFound:
        nomi_disponibili = ", ".join(f"'{f.title}'" for f in _workbook.worksheets())
        return None, (
            f"Il foglio '{NOME_FOGLIO_TUTTI}' non esiste nel documento collegato. "
            f"Fogli disponibili: {nomi_disponibili}."
        )
    except Exception as e:
        return None, f"Errore durante la lettura del foglio: {e}"

    tutti_i_valori = ws.get_all_values()
    righe_dati = tutti_i_valori[RIGA_INTESTAZIONE_TUTTI:]

    record = []
    ultima_colonna_utile = max(COL_TUTTI_NOME, COL_TUTTI_MESE, COL_TUTTI_TIPO_SERVIZIO,
                                COL_TUTTI_MINISTERO, COL_TUTTI_ORE, COL_TUTTI_CRED_ORE,
                                COL_TUTTI_STUDI, COL_TUTTI_OSSERVAZIONI)
    for i, r in enumerate(righe_dati):
        if len(r) <= ultima_colonna_utile:
            r = r + [""] * (ultima_colonna_utile + 1 - len(r))
        nome = r[COL_TUTTI_NOME].strip()
        if not nome:
            continue

        mese_anno = r[COL_TUTTI_MESE].strip()
        tipo_servizio = r[COL_TUTTI_TIPO_SERVIZIO].strip()
        ministero_testo = r[COL_TUTTI_MINISTERO].strip()
        ore_testo = r[COL_TUTTI_ORE].strip()
        cred_ore_testo = r[COL_TUTTI_CRED_ORE].strip()

        ha_partecipato = (ministero_testo.lower() in ("si", "sì")) or (a_float_it(ore_testo) > 0)
        pioniere_ausiliario = "pioniere ausiliario" in tipo_servizio.lower()

        # Numero di riga reale nel foglio (1-based) e valori grezzi delle
        # colonne C..J, necessari per poter salvare le modifiche più avanti
        # senza perdere il contenuto delle colonne che non gestiamo (es. F).
        riga_foglio = RIGA_INTESTAZIONE_TUTTI + 1 + i
        grezza = r[2:10]  # colonne C,D,E,F,G,H,I,J (indici 2..9)

        record.append({
            "Nome": nome,
            "Mese/Anno": mese_anno,
            "Anno di servizio": formatta_mese_esteso(mese_anno),
            "Ha partecipato al ministero": ha_partecipato,
            "Studi Biblici": r[COL_TUTTI_STUDI].strip(),
            "Pioniere ausiliario": pioniere_ausiliario,
            "Ore": ore_testo,
            "Cred. Ore": cred_ore_testo,
            "Osservazioni": r[COL_TUTTI_OSSERVAZIONI].strip(),
            "RigaFoglio": riga_foglio,
            "_grezza": grezza,
        })
    return pd.DataFrame(record), None


def anno_teocratico_di(mese_anno: str):
    """Dato un valore 'Mese/Anno' in formato AAAA-MM (es. '2026-06'),
    ritorna l'anno teocratico di appartenenza (l'anno teocratico X va da
    settembre X ad agosto X+1). Ritorna None se il formato non è valido."""
    try:
        anno_str, mese_str = mese_anno.split("-")
        anno, mese = int(anno_str), int(mese_str)
    except Exception:
        return None
    return anno if mese >= 9 else anno - 1


def prossimo_id_anagrafica(df: pd.DataFrame) -> int:
    if "ID" not in df.columns or df.empty:
        return 1
    numeri = pd.to_numeric(df["ID"], errors="coerce").dropna()
    return int(numeri.max()) + 1 if not numeri.empty else 1


def salva_riga_foglio(_workbook, nome_foglio: str, riga_intestazione: int,
                       valori: dict, riga_da_aggiornare: int = None):
    """Scrive una nuova riga in fondo a un foglio, oppure aggiorna una riga
    esistente (numero di riga del foglio, 1-based) se 'riga_da_aggiornare'
    è specificato. 'valori' è un dizionario {nome_colonna: valore}; le
    colonne del foglio non presenti nel dizionario vengono lasciate vuote
    (nuova riga) o svuotate (modifica: si riscrive l'intera riga in
    ordine). Ritorna (successo: bool, errore: str|None)."""
    try:
        ws = _workbook.worksheet(nome_foglio)
        intestazioni = ws.row_values(riga_intestazione)
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


def salva_riga_tutti(_workbook, riga_foglio: int, nuova_grezza: list):
    """Aggiorna una riga del foglio 'Tutti' (colonne C:J) con i nuovi
    valori, preservando inalterata la colonna F (che non gestiamo nel
    form). 'nuova_grezza' deve avere 8 elementi (C,D,E,F,G,H,I,J).
    Ritorna (successo: bool, errore: str|None)."""
    try:
        ws = _workbook.worksheet(NOME_FOGLIO_TUTTI)
        ws.update(f"C{riga_foglio}:J{riga_foglio}", [nuova_grezza], value_input_option="USER_ENTERED")
        return True, None
    except Exception as e:
        return False, f"Errore durante il salvataggio: {e}"


def salva_riga_anagrafica(_workbook, valori: dict, riga_da_aggiornare: int = None):
    """Scorciatoia per salvare una riga nel foglio Anagrafica (vedi
    'salva_riga_foglio' per i dettagli)."""
    return salva_riga_foglio(_workbook, NOME_FOGLIO_ANAGRAFICA, RIGA_INTESTAZIONE_ANAGRAFICA,
                              valori, riga_da_aggiornare)


# ─────────────────────────────────────────────────────────────────
# CONNESSIONE (navigazione tramite le card)
# ─────────────────────────────────────────────────────────────────
if "pagina" not in st.session_state:
    st.session_state.pagina = "home"


def vai_a(pagina: str):
    st.session_state.pagina = pagina


workbook, errore = apri_foglio_dati()
collegato = workbook is not None


# ─────────────────────────────────────────────────────────────────
# PAGINA: HOME
# ─────────────────────────────────────────────────────────────────
def mostra_home():
    st.title("Pannello di controllo")

    st.subheader("Sezioni")
    c1, c2, c3 = st.columns(3)
    card_data = [
        ("📖", "Rapp. consegnati", "Visualizza e modifica i rapporti di servizio consegnati.", "registrazioni"),
        ("🗂️", "Anagrafiche", "Gestisci i dati dei Proclamatori.", "anagrafiche"),
        ("📚", "Storico rapporti", "Storico dei rapporti di servizio per Proclamatore.", "storico"),
    ]
    for col, (icon, titolo, desc, pagina) in zip((c1, c2, c3), card_data):
        with col:
            with st.container(border=True):
                st.markdown(f"#### {icon}  {titolo}")
                st.caption(desc)
                st.button("Apri →", key=f"card_{titolo}", disabled=not collegato,
                          use_container_width=True, on_click=vai_a, args=(pagina,))

    st.caption(f"Ultimo aggiornamento pagina: {datetime.now().strftime('%d/%m/%Y %H:%M')}")


# ─────────────────────────────────────────────────────────────────
# PAGINA: VISUALIZZA REGISTRAZIONI
# ─────────────────────────────────────────────────────────────────
def _form_rapporto(df: pd.DataFrame, riga_esistente: dict, numero_riga_foglio: int):
    """Form generico di modifica per una riga del foglio 'Risposte del
    modulo 9': un campo per ciascuna colonna del foglio, precompilato con
    i valori attuali. I campi numerici (Ore, Cr. Ore, Studi) usano un
    campo numerico, gli altri un campo di testo. Alcune colonne (es.
    'Video mostrati') non vengono mostrate nel form ma il loro valore
    esistente viene comunque preservato al salvataggio."""
    e = riga_esistente
    colonne_numeriche = {"ore", "cr. ore", "cr ore", "studi"}
    colonne_nascoste = {"video mostrati"}

    with st.form("form_rapporto", clear_on_submit=False):
        valori_inseriti = {}
        for colonna in df.columns:
            valore_attuale = e.get(colonna, "")
            chiave_norm = colonna.strip().lower()
            if chiave_norm in colonne_nascoste:
                valori_inseriti[colonna] = valore_attuale
                continue
            if "hai servito" in chiave_norm:
                opzioni = list(OPZIONI_HAI_SERVITO)
                if valore_attuale and valore_attuale not in opzioni:
                    opzioni = [valore_attuale] + opzioni
                indice = opzioni.index(valore_attuale) if valore_attuale in opzioni else 0
                valori_inseriti[colonna] = st.selectbox(colonna, opzioni, index=indice)
            elif chiave_norm in colonne_numeriche:
                try:
                    default_num = float(str(valore_attuale).replace(",", ".")) if valore_attuale else 0.0
                except ValueError:
                    default_num = 0.0
                valori_inseriti[colonna] = st.number_input(colonna, value=default_num, step=1.0)
            else:
                valori_inseriti[colonna] = st.text_input(colonna, value=str(valore_attuale))

        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            invia = st.form_submit_button("✔ Salva", type="primary", use_container_width=True)
        with col_btn2:
            annulla = st.form_submit_button("Annulla", use_container_width=True)

    if annulla:
        st.session_state.rapporto_modifica = None
        st.rerun()

    if invia:
        valori_finali = {
            colonna: (str(v) if not isinstance(v, str) else v)
            for colonna, v in valori_inseriti.items()
        }
        ok, err = salva_riga_foglio(workbook, NOME_FOGLIO_RISPOSTE, RIGA_INTESTAZIONE_RISPOSTE,
                                     valori_finali, riga_da_aggiornare=numero_riga_foglio)
        if ok:
            st.cache_data.clear()
            st.session_state.rapporto_modifica = None
            st.success("✔ Salvato correttamente.")
            st.rerun()
        else:
            st.error(err)


def mostra_registrazioni():
    if st.button("🏠 Torna alla Home", key="home_da_registrazioni", type="primary", use_container_width=True):
        vai_a("home")
        st.rerun()
    st.title("Rapporti consegnati")
    st.caption(f"Dati letti dal foglio «{NOME_FOGLIO_RISPOSTE}» (intestazione riga {RIGA_INTESTAZIONE_RISPOSTE}).")

    if not collegato:
        st.warning("⚠️  Nessun foglio dati collegato.")
        return

    if "rapporto_modifica" not in st.session_state:
        st.session_state.rapporto_modifica = None

    df, err = leggi_foglio_come_df(workbook, NOME_FOGLIO_RISPOSTE, RIGA_INTESTAZIONE_RISPOSTE)

    if err:
        st.error(err)
        return

    if df.empty:
        st.info("Il foglio è collegato correttamente ma non contiene ancora righe di dati.")
        return

    # ── Form di modifica di una riga selezionata ────────────────────────
    if st.session_state.rapporto_modifica is not None:
        idx = st.session_state.rapporto_modifica
        riga = df.iloc[idx].to_dict()
        numero_riga_foglio = RIGA_INTESTAZIONE_RISPOSTE + 1 + idx
        st.subheader(f"✏️ Modifica rapporto: {riga.get('Cognome e Nome', '')}")
        _form_rapporto(df, riga, numero_riga_foglio)
        return

    contenitore_azioni = st.container()

    ricerca = st.text_input("🔍 Cerca in tutte le colonne",
                             placeholder="Digita per filtrare le righe…")

    df_mostrato = df.reset_index(drop=True)
    if ricerca:
        maschera = df_mostrato.apply(
            lambda riga: riga.astype(str).str.contains(ricerca, case=False, na=False).any(),
            axis=1,
        )
        df_mostrato = df_mostrato[maschera]

    st.caption(f"{len(df_mostrato)} righe su {len(df)} totali. Tocca una riga per selezionarla, poi «Modifica».")

    df_visualizzato = df_mostrato.drop(columns=["Video mostrati"], errors="ignore")

    config_colonne = {}
    for nome_colonna in ["Mese", "Ore", "Cr. Ore", "Studi"]:
        if nome_colonna in df_visualizzato.columns:
            config_colonne[nome_colonna] = st.column_config.TextColumn(nome_colonna, width="small")
    for col in df_visualizzato.columns:
        if "osservazioni" in col.lower():
            config_colonne[col] = st.column_config.TextColumn(col, width="large")

    evento = st.dataframe(
        df_visualizzato,
        use_container_width=True,
        hide_index=True,
        height=560,
        on_select="rerun",
        selection_mode="single-row",
        column_config=config_colonne,
    )

    righe_selezionate = evento.selection.rows if evento and evento.selection else []
    idx_originale = None
    nome_sel = None
    if righe_selezionate:
        posizione = righe_selezionate[0]
        idx_originale = df_mostrato.index[posizione]
        if "Cognome e Nome" in df_mostrato.columns:
            nome_sel = df_mostrato.loc[idx_originale, "Cognome e Nome"]

    with contenitore_azioni:
        c_modifica, c_aggiorna = st.columns([4, 1])
        with c_modifica:
            etichetta = f"✏️ Modifica «{nome_sel}»" if nome_sel else "✏️ Modifica"
            if st.button(etichetta, disabled=idx_originale is None, type="primary",
                         use_container_width=True):
                st.session_state.rapporto_modifica = idx_originale
                st.rerun()
        with c_aggiorna:
            if st.button("🔄", use_container_width=True):
                st.cache_data.clear()
                st.rerun()


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
    if st.button("🏠 Torna alla Home", key="home_da_anagrafiche", type="primary", use_container_width=True):
        vai_a("home")
        st.rerun()
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
                            ["Cognome e Nome", "Data Nascita", "Data Battesimo", "Sesso", "Incarico", "Tipo",
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
        "Data Battesimo": "Battesimo",
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
            "Battesimo": st.column_config.TextColumn(width="small"),
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
# PAGINA: STORICO PROCLAMATORI
# ─────────────────────────────────────────────────────────────────
def _form_modifica_rapporto_tutti(dati_selezione: dict):
    """Form di modifica di una riga del foglio 'Tutti' (colonne C..J),
    aperto cliccando una riga dentro la tabella espansa di un Proclamatore
    in Storico rapporti consegnati."""
    nome = dati_selezione["nome"]
    mese_leggibile = dati_selezione["mese_leggibile"]
    riga_foglio = dati_selezione["riga_foglio"]
    grezza = dati_selezione["grezza"]  # [C, D, E, F, G, H, I, J]

    st.title("Modifica rapporto")
    st.caption(f"{nome} — {mese_leggibile} (foglio «{NOME_FOGLIO_TUTTI}», riga {riga_foglio})")

    with st.form("form_modifica_tutti", clear_on_submit=False):
        mese_anno = st.text_input("Mese/Anno (formato AAAA-MM)", value=grezza[0])

        opzioni_tipo = list(OPZIONI_HAI_SERVITO)
        valore_tipo = grezza[1]
        if valore_tipo and valore_tipo not in opzioni_tipo:
            opzioni_tipo = [valore_tipo] + opzioni_tipo
        indice_tipo = opzioni_tipo.index(valore_tipo) if valore_tipo in opzioni_tipo else 0
        tipo_servizio = st.selectbox("Ha servito come", opzioni_tipo, index=indice_tipo)

        opzioni_ministero = ["Si", "No"]
        valore_ministero = grezza[2] if grezza[2] in opzioni_ministero else "No"
        ministero = st.selectbox("Ha partecipato al ministero", opzioni_ministero,
                                  index=opzioni_ministero.index(valore_ministero))

        ore = st.number_input("Ore", value=a_float_it(grezza[4]), step=1.0)
        cred_ore = st.number_input("Cred. Ore", value=a_float_it(grezza[5]), step=1.0)
        studi = st.text_input("Studi Biblici", value=grezza[6])
        osservazioni = st.text_area("Osservazioni", value=grezza[7])

        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            invia = st.form_submit_button("✔ Salva", type="primary", use_container_width=True)
        with col_btn2:
            annulla = st.form_submit_button("Annulla", use_container_width=True)

    if annulla:
        st.session_state.storico_modifica = None
        st.rerun()

    if invia:
        nuova_grezza = list(grezza)
        nuova_grezza[0] = mese_anno.strip()
        nuova_grezza[1] = tipo_servizio
        nuova_grezza[2] = ministero
        nuova_grezza[4] = formatta_numero_it(ore)
        nuova_grezza[5] = formatta_numero_it(cred_ore)
        nuova_grezza[6] = studi.strip()
        nuova_grezza[7] = osservazioni.strip()
        # nuova_grezza[3] (colonna F) resta invariata, non la gestiamo nel form

        ok, err = salva_riga_tutti(workbook, riga_foglio, nuova_grezza)
        if ok:
            st.cache_data.clear()
            st.session_state.storico_modifica = None
            st.success("✔ Salvato correttamente.")
            st.rerun()
        else:
            st.error(err)


def mostra_storico_proclamatori():
    if st.button("🏠 Torna alla Home", key="home_da_storico", type="primary", use_container_width=True):
        vai_a("home")
        st.rerun()
    st.title("Storico rapporti consegnati")
    st.caption(f"Rapporti storici letti dal foglio «{NOME_FOGLIO_TUTTI}» "
               f"(intestazione riga {RIGA_INTESTAZIONE_TUTTI}).")

    if not collegato:
        st.warning("⚠️  Nessun foglio dati collegato.")
        return

    if "storico_modifica" not in st.session_state:
        st.session_state.storico_modifica = None

    if st.session_state.storico_modifica is not None:
        _form_modifica_rapporto_tutti(st.session_state.storico_modifica)
        return

    df_anagrafica, err_anagrafica = leggi_foglio_come_df(
        workbook, NOME_FOGLIO_ANAGRAFICA, RIGA_INTESTAZIONE_ANAGRAFICA)
    if err_anagrafica:
        st.error(err_anagrafica)
        return

    df_tutti, err_tutti = leggi_foglio_tutti(workbook)
    if err_tutti:
        st.error(err_tutti)
        return

    # ── Elenco anni teocratici disponibili, calcolati dai dati presenti ──
    anni_presenti = sorted({
        a for a in (anno_teocratico_di(m) for m in df_tutti["Mese/Anno"]) if a is not None
    }, reverse=True)
    if not anni_presenti:
        anno_corrente = datetime.now().year
        anni_presenti = [anno_corrente - 1, anno_corrente]

    col_anno, col_ricerca = st.columns([1, 3])
    with col_anno:
        anno_scelto = st.selectbox(
            "Anno teocratico",
            anni_presenti,
            format_func=lambda a: f"{a} – {a + 1} (set {a} → ago {a + 1})",
        )
    with col_ricerca:
        ricerca = st.text_input("🔍 Cerca per nome", placeholder="Digita per filtrare…")

    if df_anagrafica.empty or "Cognome e Nome" not in df_anagrafica.columns:
        st.info("Nessun Proclamatore trovato in Anagrafica.")
        return

    def e_inattivo(valore: str) -> bool:
        """Riconosce lo stato 'Inattivo' sia scritto come 'I' che come
        'Inattivi'/'Inattivo' per esteso — tutto ciò che inizia con 'i'."""
        return (valore or "").strip().lower().startswith("i")

    def stato_valido(valore: str) -> bool:
        """Include solo Attivi ('A'/'Attivi') e Inattivi ('I'/'Inattivi');
        esclude tutto il resto (es. 'TR'/'Trasferiti')."""
        v = (valore or "").strip().lower()
        return v.startswith("a") or v.startswith("i")

    colonna_stato = "Attivi / Inattivi" if "Attivi / Inattivi" in df_anagrafica.columns else None
    colonna_gruppo = "Gruppo" if "Gruppo" in df_anagrafica.columns else None

    stato_per_nome = {}
    gruppo_per_nome = {}
    for _, riga in df_anagrafica.iterrows():
        n = str(riga.get("Cognome e Nome", "")).strip()
        if not n:
            continue
        if colonna_stato:
            stato_per_nome[n] = riga.get(colonna_stato, "")
        if colonna_gruppo:
            gruppo_per_nome[n] = str(riga.get(colonna_gruppo, "")).strip()

    nomi = sorted(n for n in df_anagrafica["Cognome e Nome"].astype(str).str.strip().unique() if n)
    if colonna_stato:
        nomi = [n for n in nomi if stato_valido(stato_per_nome.get(n, ""))]
    if ricerca:
        nomi = [n for n in nomi if ricerca.lower() in n.lower()]

    st.caption(f"{len(nomi)} Proclamatori. Il triangolo rosso 🔺 indica i Proclamatori Inattivi.")

    # ── Raggruppamento alfabetico per Gruppo (sorvegliante) ──────────────
    gruppi = {}
    for n in nomi:
        g = gruppo_per_nome.get(n, "") or "(Senza gruppo)"
        gruppi.setdefault(g, []).append(n)
    for g in gruppi:
        gruppi[g].sort()

    def _riga_proclamatore(nome: str):
        inattivo = e_inattivo(stato_per_nome.get(nome, ""))
        indicatore = "🔺 " if inattivo else ""
        etichetta = f"{indicatore}{nome}"

        with st.expander(etichetta):
            righe_persona = df_tutti[df_tutti["Nome"].str.strip().str.lower() == nome.strip().lower()]
            righe_persona = righe_persona[
                righe_persona["Mese/Anno"].apply(anno_teocratico_di) == anno_scelto
            ]
            colonne_tabella = ["Anno di servizio", "Ha partecipato al ministero", "Studi Biblici",
                                "Pioniere ausiliario", "Ore", "Cred. Ore", "Osservazioni"]
            if righe_persona.empty:
                st.caption("Nessun rapporto trovato per l'anno teocratico selezionato.")
            else:
                righe_persona = righe_persona.sort_values("Mese/Anno")
                totale_ore = sum(a_float_it(v) for v in righe_persona["Ore"])
                totale_cred = sum(a_float_it(v) for v in righe_persona["Cred. Ore"])
                riga_totale = pd.DataFrame([{
                    "Anno di servizio": "Totale",
                    "Ha partecipato al ministero": None,
                    "Studi Biblici": "",
                    "Pioniere ausiliario": None,
                    "Ore": formatta_numero_it(totale_ore),
                    "Cred. Ore": formatta_numero_it(totale_cred),
                    "Osservazioni": "",
                }])
                tabella_completa = pd.concat(
                    [righe_persona[colonne_tabella], riga_totale[colonne_tabella]], ignore_index=True
                )
                evento_tabella = st.dataframe(
                    tabella_completa,
                    hide_index=True,
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode="single-row",
                    key=f"storico_tabella_{nome}",
                    column_config={
                        "Anno di servizio": st.column_config.TextColumn(width="small"),
                        "Ha partecipato al ministero": st.column_config.CheckboxColumn(
                            "Ha partecipato al ministero", width="small", disabled=True),
                        "Studi Biblici": st.column_config.TextColumn(width="small"),
                        "Pioniere ausiliario": st.column_config.CheckboxColumn(
                            "Pioniere ausiliario", width="small", disabled=True),
                        "Ore": st.column_config.TextColumn(width="small"),
                        "Cred. Ore": st.column_config.TextColumn(width="small"),
                        "Osservazioni": st.column_config.TextColumn(width="large"),
                    },
                )

                righe_sel = evento_tabella.selection.rows if evento_tabella and evento_tabella.selection else []
                if righe_sel:
                    posizione = righe_sel[0]
                    # La riga "Totale" (ultima) non è modificabile
                    if posizione < len(righe_persona):
                        idx_originale = righe_persona.index[posizione]
                        riga_dati = df_tutti.loc[idx_originale]
                        mese_leggibile = riga_dati["Anno di servizio"]
                        if st.button(f"✏️ Modifica «{mese_leggibile}»", key=f"storico_modifica_btn_{nome}_{posizione}"):
                            st.session_state.storico_modifica = {
                                "nome": nome,
                                "mese_leggibile": mese_leggibile,
                                "riga_foglio": int(riga_dati["RigaFoglio"]),
                                "grezza": list(riga_dati["_grezza"]),
                            }
                            st.rerun()

    for gruppo in sorted(gruppi.keys()):
        st.markdown(f"#### 👤 {gruppo}")
        for nome in gruppi[gruppo]:
            _riga_proclamatore(nome)
        st.divider()


# ─────────────────────────────────────────────────────────────────
# ROUTING — navigazione solo tramite le card della Home
# ─────────────────────────────────────────────────────────────────
if st.session_state.pagina == "registrazioni":
    mostra_registrazioni()
elif st.session_state.pagina == "anagrafiche":
    mostra_anagrafiche()
elif st.session_state.pagina == "storico":
    mostra_storico_proclamatori()
else:
    mostra_home()