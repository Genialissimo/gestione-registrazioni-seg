"""
app.py
Gestione Registrazioni SEG - Web App (Streamlit + Google Sheets)
Collegamento a un unico foglio Google + Home + Visualizza registrazioni.
Altri form (Nuova registrazione, Anagrafiche) verranno aggiunti in seguito.
"""

from datetime import datetime
import io
import os
import zipfile

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas

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
OPZIONI_ATTIVI_INATTIVI = ["A", "I", "TR"]
ETICHETTE_ATTIVI_INATTIVI = {"A": "Attivo", "I": "Inattivo", "TR": "Trasferito"}


def categoria_stato_proclamatore(valore: str) -> str:
    """Riconosce la categoria di stato da 'Attivi / Inattivi', sia scritto
    come sigla ('A'/'I'/'TR') che per esteso. Ritorna 'A' come valore di
    default se non riconosciuto."""
    v = (valore or "").strip().lower()
    if v.startswith("i"):
        return "I"
    if v.startswith("t"):
        return "TR"
    return "A"

# ── Modulo S-21 (scheda di registrazione del proclamatore) ──────────
# Il file del modello va messo nella stessa cartella di app.py.
PERCORSO_MODULO_S21 = os.path.join(os.path.dirname(__file__), "S-21_s-Mlt_I.pdf")
S21_PAGE_W, S21_PAGE_H = 595.2, 841.9
S21_OFFSET_PANNELLO = 421.0  # distanza verticale tra il pannello alto e quello basso

S21_ORDINE_MESI = ["Settembre", "Ottobre", "Novembre", "Dicembre", "Gennaio", "Febbraio",
                    "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto"]

# Coordinate (top, bottom) di ciascuna riga mensile, misurate sui quadratini
# reali del modello (non sulle etichette) per un allineamento preciso.
S21_RIGHE = {
    "Settembre": (158.44, 171.09), "Ottobre": (176.17, 188.83), "Novembre": (193.90, 206.56),
    "Dicembre": (211.64, 224.29), "Gennaio": (229.37, 242.02), "Febbraio": (247.10, 259.75),
    "Marzo": (264.83, 277.48), "Aprile": (282.56, 295.21), "Maggio": (300.29, 312.94),
    "Giugno": (318.02, 330.68), "Luglio": (335.75, 348.41), "Agosto": (353.48, 366.14),
}
S21_TOTALE_RIGA = (371.9, 389.5)

# Colonne della tabella mensile: (x0, x1) dei quadratini o del corpo colonna.
S21_COL_MINISTERO = (130.28, 143.01)
S21_COL_AUSILIARIO = (271.52, 284.25)
S21_COL_STUDI = (172.0, 242.6)
S21_COL_ORE = (313.2, 383.8)
S21_COL_OSSERVAZIONI_X = 388.5  # allineato a sinistra, con piccolo margine dal bordo colonna (383.8)

# Quadratini della testata (sesso, classe spirituale, incarico/tipo), misurati
# sui glifi reali del modello: (x0, x1, top, bottom).
S21_BOX_SESSO_M = (384.18, 394.4, 53.12, 63.28)
S21_BOX_SESSO_F = (485.58, 495.8, 53.12, 63.28)
S21_BOX_ALTRE_PECORE = (384.18, 394.4, 67.65, 77.8)
S21_BOX_UNTO = (485.58, 495.8, 67.65, 77.8)
S21_BOX_ANZIANO = (17.03, 27.25, 82.08, 92.24)
S21_BOX_SERVITORE = (80.58, 90.8, 82.08, 92.24)
S21_BOX_PIONIERE_REGOLARE = (216.05, 226.27, 82.08, 92.24)
S21_BOX_PIONIERE_SPECIALE = (327.5, 337.72, 82.08, 92.24)
S21_BOX_MISSIONARIO = (436.37, 446.59, 82.08, 92.24)

# Font (aumentati di 1-2 pt rispetto alla versione precedente).
S21_COL_ANNO_SERVIZIO = (17.5, 101.4)  # colonna "Anno di servizio"
S21_ANNO_LABEL_TOP = 145.5  # tra l'intestazione di colonna e la prima riga (Settembre)
S21_SPOSTAMENTO_RIGHE = 2.0  # piccolo spostamento verso il basso per le righe della tabella mensile
S21_SPOSTAMENTO_TESTATA = 1.8  # spostamento verso il basso per le X di Sesso/Classe/Incarico

S21_FONT_VALORI = 10.5      # nome, date
S21_FONT_CHECK_HEADER = 9.5  # X nei quadratini della testata
S21_FONT_TABELLA = 9.5      # testo/numeri nella tabella mensile
S21_FONT_CHECK_TABELLA = 10.5  # X nei quadratini della tabella mensile
S21_FONT_ETA = 8.5          # età calcolata accanto alle date, in rosso

S21_COLORE_ROSSO = (0.827, 0.125, 0.125)  # stesso rosso #D32F2F usato nel form web
S21_COLORE_NERO = (0, 0, 0)



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
            "Tipo Servizio": tipo_servizio,
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


def _s21_y_da_top(top: float, offset: float = 0.0, alza: float = 8.0) -> float:
    """Converte una coordinata 'top' (distanza dall'alto pagina, come la
    ritorna pdfplumber) nella coordinata Y usata da reportlab (che parte
    dal basso pagina). 'offset' sposta il calcolo sul pannello basso."""
    return S21_PAGE_H - (top + offset + alza)


def _s21_y_da_bottom(bottom: float, offset: float = 0.0, alza: float = 1.5) -> float:
    return S21_PAGE_H - (bottom + offset - alza)


def _s21_centro_box(c: rl_canvas.Canvas, box: tuple, offset: float, testo: str = "X",
                     font_name: str = "Helvetica-Bold", font_size: float = 10.0, sposta: float = 0.0):
    """Disegna 'testo' (di norma una 'X') perfettamente centrato — sia in
    orizzontale che in verticale — dentro un quadratino le cui coordinate
    reali (x0, x1, top, bottom) sono state misurate sul modello PDF.
    'sposta' aggiunge un piccolo spostamento verticale (positivo = più in
    basso), usato per le righe della tabella mensile."""
    x0, x1, top, bottom = box
    fattore_altezza_maiuscole = 0.717  # approssimazione per Helvetica-Bold
    largo_testo = c.stringWidth(testo, font_name, font_size)
    x = (x0 + x1) / 2 - largo_testo / 2
    centro_verticale_top = (top + bottom) / 2 + offset
    baseline_top = centro_verticale_top + (font_size * fattore_altezza_maiuscole) / 2 + sposta
    y = S21_PAGE_H - baseline_top
    c.setFont(font_name, font_size)
    c.drawString(x, y, testo)


def _s21_testo_centrato_colonna(c: rl_canvas.Canvas, testo: str, col: tuple, top: float, bottom: float,
                                 offset: float, font_name: str = "Helvetica", font_size: float = 9.5,
                                 sposta: float = 0.0):
    """Scrive 'testo' centrato orizzontalmente in una colonna (x0, x1) e
    centrato verticalmente in una riga (top, bottom). 'sposta' aggiunge un
    piccolo spostamento verticale (positivo = più in basso)."""
    x0, x1 = col
    largo_testo = c.stringWidth(testo, font_name, font_size)
    x = (x0 + x1) / 2 - largo_testo / 2
    c.setFont(font_name, font_size)
    c.drawString(x, _s21_y_da_top((top + bottom) / 2, offset, alza=font_size * 0.36 + sposta), testo)



def _s21_righe_anno_per_nome(df_tutti: pd.DataFrame, nome: str, anno_teocratico) -> dict:
    """Ritorna un dizionario {nome_mese: {ha_partecipato, pioniere_ausiliario,
    studi, ore, cred_ore, osservazioni}} per un Proclamatore e un anno
    teocratico dati, pescando dal DataFrame prodotto da leggi_foglio_tutti().
    Ritorna un dizionario vuoto se l'anno è None o non ci sono rapporti."""
    if anno_teocratico is None or df_tutti.empty:
        return {}
    righe_persona = df_tutti[df_tutti["Nome"].str.strip().str.lower() == nome.strip().lower()]
    righe_persona = righe_persona[righe_persona["Mese/Anno"].apply(anno_teocratico_di) == anno_teocratico]

    risultato = {}
    for _, r in righe_persona.iterrows():
        risultato[r["Anno di servizio"]] = {
            "ha_partecipato": bool(r["Ha partecipato al ministero"]),
            "pioniere_ausiliario": bool(r["Pioniere ausiliario"]),
            "studi": r["Studi Biblici"],
            "ore": r["Ore"],
            "cred_ore": r.get("Cred. Ore", ""),
            "osservazioni": r["Osservazioni"],
        }
    return risultato


def _s21_righe_anno_aggregate(df_tutti: pd.DataFrame, nomi: list, anno_teocratico,
                               etichetta_conteggio: str = "proclamatori") -> dict:
    """Come _s21_righe_anno_per_nome, ma aggrega più Proclamatori insieme:
    Ore/Studi/Cred. Ore sono la somma del mese. Le caselle Ministero e
    Pioniere ausiliario vengono marcate con una X se quel mese almeno una
    persona della categoria ha partecipato/fatto l'ausiliario. Nelle
    Osservazioni compare un solo numero: quante persone DI QUESTA CATEGORIA
    hanno un rapporto quel mese (es. '8 pionieri regolari') — mai dati di
    altre categorie mescolati. Usata per le cartoline di riepilogo (Tutti i
    proclamatori, Pionieri Regolari, ecc.)."""
    if anno_teocratico is None or df_tutti.empty or not nomi:
        return {}
    nomi_lower = {n.strip().lower() for n in nomi if n and n.strip()}
    if not nomi_lower:
        return {}
    righe = df_tutti[df_tutti["Nome"].str.strip().str.lower().isin(nomi_lower)]
    righe = righe[righe["Mese/Anno"].apply(anno_teocratico_di) == anno_teocratico]
    if righe.empty:
        return {}

    risultato = {}
    for mese in S21_ORDINE_MESI:
        righe_mese = righe[righe["Anno di servizio"] == mese]
        if righe_mese.empty:
            continue
        conteggio_ministero = int(righe_mese["Ha partecipato al ministero"].astype(bool).sum())
        conteggio_ausiliario = int(righe_mese["Pioniere ausiliario"].astype(bool).sum())
        studi_tot = sum(a_float_it(v) for v in righe_mese["Studi Biblici"])
        ore_tot = sum(a_float_it(v) for v in righe_mese["Ore"])
        cred_tot = sum(a_float_it(v) for v in righe_mese["Cred. Ore"]) if "Cred. Ore" in righe_mese.columns else 0.0

        risultato[mese] = {
            "ha_partecipato": conteggio_ministero > 0,
            "pioniere_ausiliario": conteggio_ausiliario > 0,
            "studi": formatta_numero_it(studi_tot) if studi_tot else "",
            "ore": formatta_numero_it(ore_tot) if ore_tot else "",
            "cred_ore": formatta_numero_it(cred_tot) if cred_tot else "",
            "osservazioni": f"{conteggio_ministero} {etichetta_conteggio}" if conteggio_ministero else "",
        }
    return risultato


def _s21_righe_anno_aggregate_per_tipo(df_tutti: pd.DataFrame, nomi_pool: list, anno_teocratico,
                                        parola_chiave_tipo: str, etichetta_conteggio: str) -> dict:
    """Aggrega Ore/Studi/Cred. Ore mese per mese SOLO per le righe del
    foglio Tutti il cui 'Tipo di servizio' di quel mese (colonna D) contiene
    'parola_chiave_tipo' (es. 'pioniere regolare', 'pioniere speciale',
    'missionario', 'pioniere ausiliario') — non una lista fissa presa dal
    campo Tipo in Anagrafica. Così, ad esempio, chi un mese risulta
    Pioniere Ausiliario in Tutti finisce SOLO nel riepilogo Ausiliari
    quel mese, anche se è normalmente registrato come Pioniere Regolare.
    'nomi_pool' limita la ricerca (di norma tutti gli Attivi). Nelle
    Osservazioni viene scritto solo il conteggio trovato quel mese, con
    'etichetta_conteggio' (es. '11 pionieri regolari')."""
    if anno_teocratico is None or df_tutti.empty or not nomi_pool:
        return {}
    nomi_lower = {n.strip().lower() for n in nomi_pool if n and n.strip()}
    if not nomi_lower:
        return {}
    righe = df_tutti[df_tutti["Nome"].str.strip().str.lower().isin(nomi_lower)]
    righe = righe[righe["Mese/Anno"].apply(anno_teocratico_di) == anno_teocratico]
    righe = righe[righe["Tipo Servizio"].str.lower().str.contains(parola_chiave_tipo, na=False, regex=True)]
    if righe.empty:
        return {}

    risultato = {}
    for mese in S21_ORDINE_MESI:
        righe_mese = righe[righe["Anno di servizio"] == mese]
        if righe_mese.empty:
            continue
        conteggio = len(righe_mese)
        conteggio_ministero = int(righe_mese["Ha partecipato al ministero"].astype(bool).sum())
        conteggio_ausiliario = int(righe_mese["Pioniere ausiliario"].astype(bool).sum())
        studi_tot = sum(a_float_it(v) for v in righe_mese["Studi Biblici"])
        ore_tot = sum(a_float_it(v) for v in righe_mese["Ore"])
        cred_tot = sum(a_float_it(v) for v in righe_mese["Cred. Ore"]) if "Cred. Ore" in righe_mese.columns else 0.0

        risultato[mese] = {
            "ha_partecipato": conteggio_ministero > 0,
            "pioniere_ausiliario": conteggio_ausiliario > 0,
            "studi": formatta_numero_it(studi_tot) if studi_tot else "",
            "ore": formatta_numero_it(ore_tot) if ore_tot else "",
            "cred_ore": formatta_numero_it(cred_tot) if cred_tot else "",
            "osservazioni": f"{conteggio} {etichetta_conteggio}" if conteggio else "",
        }
    return risultato


def _s21_dati_riepilogo(titolo: str) -> dict:
    """Dati di testata per una cartolina di riepilogo aggregata: solo il
    titolo nel campo 'Nome e cognome', nessuna data/sesso/incarico (quindi
    nessuna casella della testata verrà spuntata)."""
    return {
        "nome": titolo,
        "data_nascita": "",
        "data_battesimo": "",
        "sesso": "",
        "incarico": "",
        "tipo": "",
        "classe_spirituale": "",
    }


def _s21_nota_inattivo_dal(riga_anagrafica: dict) -> str:
    """Se il Proclamatore è Inattivo, ritorna il testo 'Inattivo dal
    gg/mm/aaaa' da scrivere nella prima riga di Osservazioni. Ritorna
    stringa vuota se non è Inattivo o se la data non è compilata."""
    stato = categoria_stato_proclamatore(riga_anagrafica.get("Attivi / Inattivi", ""))
    if stato != "I":
        return ""
    data_inattivo = (riga_anagrafica.get("Inattivo dal") or riga_anagrafica.get("Dal") or "").strip()
    if not data_inattivo:
        return ""
    return f"Inattivo dal {data_inattivo}"


def _s21_con_nota_prima_riga(righe_anno: dict, nota: str) -> dict:
    """Ritorna una copia di 'righe_anno' con 'nota' anteposta alle
    Osservazioni della prima riga (Settembre) — creandola se il mese non ha
    ancora dati, così la nota compare anche su una cartolina altrimenti
    completamente vuota."""
    if not nota:
        return righe_anno
    righe_anno = dict(righe_anno)
    prima_riga = dict(righe_anno.get("Settembre", {}))
    esistente = str(prima_riga.get("osservazioni") or "").strip()
    prima_riga["osservazioni"] = f"{nota} — {esistente}" if esistente else nota
    righe_anno["Settembre"] = prima_riga
    return righe_anno


def anni_teocratici_per_menu(df_tutti: pd.DataFrame) -> list:
    """Ritorna l'elenco degli anni teocratici (int) da proporre in un menu a
    tendina: solo gli anni per cui esistono già rapporti nel foglio 'Tutti'.
    Un anno nuovo (es. 2026-2027) compare da solo non appena arriva e viene
    registrato il primo rapporto di Settembre — non prima. Ordinati dal più
    recente al più vecchio."""
    anni = set()
    if not df_tutti.empty and "Mese/Anno" in df_tutti.columns:
        anni |= {a for a in df_tutti["Mese/Anno"].apply(anno_teocratico_di) if a is not None}
    if not anni:
        oggi = datetime.now()
        anni = {oggi.year - 1, oggi.year}
    return sorted(anni, reverse=True)


def _s21_disegna_pannello(c: rl_canvas.Canvas, offset: float, dati: dict, righe_anno: dict,
                           anno_teocratico=None):
    """Disegna un pannello completo (dati anagrafici + tabella mensile) sul
    canvas di overlay. 'offset' è 0 per il pannello alto, S21_OFFSET_PANNELLO
    per quello basso. 'anno_teocratico' (es. 2025) viene mostrato come
    intervallo (es. '2025-2026') sotto l'intestazione di colonna."""
    c.setFillColorRGB(*S21_COLORE_NERO)
    c.setFont("Helvetica", S21_FONT_VALORI)
    c.drawString(116, _s21_y_da_bottom(51.5, offset), dati.get("nome", ""))

    data_nascita_str = dati.get("data_nascita", "")
    c.drawString(104, _s21_y_da_bottom(65.9, offset), data_nascita_str)
    eta = calcola_eta_dettagliata(data_nascita_str)
    if eta:
        c.setFillColorRGB(*S21_COLORE_ROSSO)
        c.setFont("Helvetica-Bold", S21_FONT_ETA)
        larghezza_data = c.stringWidth(data_nascita_str, "Helvetica", S21_FONT_VALORI)
        c.drawString(104 + larghezza_data + 6, _s21_y_da_bottom(65.9, offset), f"({eta})")
        c.setFillColorRGB(*S21_COLORE_NERO)

    data_battesimo_str = dati.get("data_battesimo", "")
    c.setFont("Helvetica", S21_FONT_VALORI)
    c.drawString(125, _s21_y_da_bottom(80.4, offset), data_battesimo_str)
    eta_batt = calcola_eta_dettagliata(data_battesimo_str)
    if eta_batt:
        c.setFillColorRGB(*S21_COLORE_ROSSO)
        c.setFont("Helvetica-Bold", S21_FONT_ETA)
        larghezza_data_batt = c.stringWidth(data_battesimo_str, "Helvetica", S21_FONT_VALORI)
        c.drawString(125 + larghezza_data_batt + 6, _s21_y_da_bottom(80.4, offset), f"({eta_batt})")
        c.setFillColorRGB(*S21_COLORE_NERO)

    if dati.get("sesso") == "M":
        _s21_centro_box(c, S21_BOX_SESSO_M, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)
    elif dati.get("sesso") == "F":
        _s21_centro_box(c, S21_BOX_SESSO_F, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)

    classe_spirituale = (dati.get("classe_spirituale") or "").strip().upper()
    if classe_spirituale in ("U", "UNTO"):
        _s21_centro_box(c, S21_BOX_UNTO, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)
    elif classe_spirituale in ("AP", "A", "ALTRE PECORE"):
        _s21_centro_box(c, S21_BOX_ALTRE_PECORE, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)

    incarico = dati.get("incarico", "")
    tipo = dati.get("tipo", "")
    if incarico == "Anziano":
        _s21_centro_box(c, S21_BOX_ANZIANO, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)
    if incarico == "Servitore di ministero":
        _s21_centro_box(c, S21_BOX_SERVITORE, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)
    if tipo == "Pioniere Regolare":
        _s21_centro_box(c, S21_BOX_PIONIERE_REGOLARE, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)
    if tipo == "Pioniere speciale":
        _s21_centro_box(c, S21_BOX_PIONIERE_SPECIALE, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)
    if tipo == "Missionario sul campo":
        _s21_centro_box(c, S21_BOX_MISSIONARIO, offset, font_size=S21_FONT_CHECK_HEADER, sposta=S21_SPOSTAMENTO_TESTATA)

    c.setFillColorRGB(*S21_COLORE_NERO)

    if anno_teocratico is not None:
        etichetta_anno = f"{anno_teocratico}-{anno_teocratico + 1}"
        c.setFont("Helvetica-Bold", S21_FONT_TABELLA)
        largo_anno = c.stringWidth(etichetta_anno, "Helvetica-Bold", S21_FONT_TABELLA)
        x0_col, x1_col = S21_COL_ANNO_SERVIZIO
        x_anno = (x0_col + x1_col) / 2 - largo_anno / 2
        c.drawString(x_anno, _s21_y_da_top(S21_ANNO_LABEL_TOP, offset, alza=S21_FONT_TABELLA * 0.36),
                     etichetta_anno)

    totale_ore = 0.0
    totale_cred_ore = 0.0
    for mese in S21_ORDINE_MESI:
        riga = righe_anno.get(mese)
        if not riga:
            continue
        top, bottom = S21_RIGHE[mese]
        if riga.get("ha_partecipato") is True:
            _s21_centro_box(c, (*S21_COL_MINISTERO, top, bottom), offset,
                             font_size=S21_FONT_CHECK_TABELLA, sposta=S21_SPOSTAMENTO_RIGHE)
        if riga.get("pioniere_ausiliario") is True:
            _s21_centro_box(c, (*S21_COL_AUSILIARIO, top, bottom), offset,
                             font_size=S21_FONT_CHECK_TABELLA, sposta=S21_SPOSTAMENTO_RIGHE)

        studi_val = str(riga.get("studi") or "").strip()
        if studi_val and studi_val != "0":
            _s21_testo_centrato_colonna(c, studi_val, S21_COL_STUDI, top, bottom, offset,
                                         font_size=S21_FONT_TABELLA, sposta=S21_SPOSTAMENTO_RIGHE)

        ore_val = str(riga.get("ore") or "").strip()
        cred_ore_val = str(riga.get("cred_ore") or "").strip()
        if ore_val and ore_val != "0":
            _s21_testo_centrato_colonna(c, ore_val, S21_COL_ORE, top, bottom, offset,
                                         font_size=S21_FONT_TABELLA, sposta=S21_SPOSTAMENTO_RIGHE)
            totale_ore += a_float_it(ore_val)
        if cred_ore_val and cred_ore_val != "0":
            totale_cred_ore += a_float_it(cred_ore_val)

        osservazioni_val = str(riga.get("osservazioni") or "").strip()
        if osservazioni_val:
            c.setFont("Helvetica", S21_FONT_TABELLA)
            c.drawString(S21_COL_OSSERVAZIONI_X, _s21_y_da_top((top + bottom) / 2, offset,
                                                                 alza=S21_FONT_TABELLA * 0.36 + S21_SPOSTAMENTO_RIGHE),
                         osservazioni_val[:44])

    top_tot, bottom_tot = S21_TOTALE_RIGA
    if totale_ore:
        _s21_testo_centrato_colonna(c, formatta_numero_it(totale_ore), S21_COL_ORE, top_tot, bottom_tot, offset,
                                     font_name="Helvetica-Bold", font_size=S21_FONT_TABELLA,
                                     sposta=S21_SPOSTAMENTO_RIGHE)
    if totale_cred_ore:
        totale_finale = totale_ore + totale_cred_ore
        testo_crediti = (f"{formatta_numero_it(totale_ore)} + ({formatta_numero_it(totale_cred_ore)}) "
                          f"= {formatta_numero_it(totale_finale)}")
        c.setFont("Helvetica", S21_FONT_TABELLA)
        c.drawString(S21_COL_OSSERVAZIONI_X, _s21_y_da_top((top_tot + bottom_tot) / 2, offset,
                                                             alza=S21_FONT_TABELLA * 0.36 + S21_SPOSTAMENTO_RIGHE),
                     testo_crediti)
    # La somma degli Studi biblici non viene più riportata nel totale, su richiesta.


def _s21_dati_da_riga_anagrafica(riga: dict) -> dict:
    """Converte una riga del foglio Anagrafica (dizionario) nel formato
    atteso da _s21_disegna_pannello."""
    return {
        "nome": riga.get("Cognome e Nome", ""),
        "data_nascita": riga.get("Data Nascita", ""),
        "data_battesimo": riga.get("Data Battesimo", ""),
        "sesso": (riga.get("Sesso", "") or "").strip().upper()[:1],
        "incarico": riga.get("Incarico", ""),
        "tipo": riga.get("Tipo", ""),
        "classe_spirituale": riga.get("A/U", ""),
    }


def genera_pdf_s21_singolo(riga_anagrafica: dict, df_tutti: pd.DataFrame, anno_corrente: int) -> bytes:
    """Genera il PDF del modulo S-21 per un singolo Proclamatore: pannello
    basso con l'anno teocratico scelto dall'utente (es. 2025 → '2025-2026'),
    pannello alto con l'anno precedente (es. 2024 → '2024-2025'). Ritorna i
    byte del PDF pronto da scaricare."""
    nome = riga_anagrafica.get("Cognome e Nome", "")
    dati = _s21_dati_da_riga_anagrafica(riga_anagrafica)

    anno_precedente = anno_corrente - 1

    righe_corrente = _s21_righe_anno_per_nome(df_tutti, nome, anno_corrente)
    righe_precedente = _s21_righe_anno_per_nome(df_tutti, nome, anno_precedente)

    nota_inattivo = _s21_nota_inattivo_dal(riga_anagrafica)
    righe_corrente = _s21_con_nota_prima_riga(righe_corrente, nota_inattivo)
    righe_precedente = _s21_con_nota_prima_riga(righe_precedente, nota_inattivo)

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(S21_PAGE_W, S21_PAGE_H))
    _s21_disegna_pannello(c, 0.0, dati, righe_precedente, anno_teocratico=anno_precedente)
    _s21_disegna_pannello(c, S21_OFFSET_PANNELLO, dati, righe_corrente, anno_teocratico=anno_corrente)
    c.save()
    buf.seek(0)

    overlay_reader = PdfReader(buf)
    template_reader = PdfReader(PERCORSO_MODULO_S21)
    writer = PdfWriter()

    pagina = template_reader.pages[0]
    pagina.merge_page(overlay_reader.pages[0])
    writer.add_page(pagina)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def genera_pdf_s21_multiplo(righe_anagrafica: list, df_tutti: pd.DataFrame, anno_corrente: int) -> bytes:
    """Genera un unico PDF con una pagina per ciascun Proclamatore passato
    in 'righe_anagrafica' (lista di dizionari, come per il singolo)."""
    writer = PdfWriter()
    template_reader = PdfReader(PERCORSO_MODULO_S21)

    for riga_anagrafica in righe_anagrafica:
        nome = riga_anagrafica.get("Cognome e Nome", "")
        dati = _s21_dati_da_riga_anagrafica(riga_anagrafica)

        anno_precedente = anno_corrente - 1

        righe_corrente = _s21_righe_anno_per_nome(df_tutti, nome, anno_corrente)
        righe_precedente = _s21_righe_anno_per_nome(df_tutti, nome, anno_precedente)

        nota_inattivo = _s21_nota_inattivo_dal(riga_anagrafica)
        righe_corrente = _s21_con_nota_prima_riga(righe_corrente, nota_inattivo)
        righe_precedente = _s21_con_nota_prima_riga(righe_precedente, nota_inattivo)

        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(S21_PAGE_W, S21_PAGE_H))
        _s21_disegna_pannello(c, 0.0, dati, righe_precedente, anno_teocratico=anno_precedente)
        _s21_disegna_pannello(c, S21_OFFSET_PANNELLO, dati, righe_corrente, anno_teocratico=anno_corrente)
        c.save()
        buf.seek(0)

        overlay_reader = PdfReader(buf)
        pagina = template_reader.pages[0]
        pagina_overlay = pagina
        # Serve una copia fresca della pagina modello per ogni Proclamatore,
        # altrimenti gli overlay si sovrapporrebbero tutti sulla stessa pagina.
        template_fresh = PdfReader(PERCORSO_MODULO_S21)
        pagina_overlay = template_fresh.pages[0]
        pagina_overlay.merge_page(overlay_reader.pages[0])
        writer.add_page(pagina_overlay)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def _s21_nome_file_sicuro(nome: str) -> str:
    """Ripulisce un nome da caratteri non ammessi nei nomi di file/cartella
    (Windows e Dropbox in particolare non ammettono \\ / : * ? " < > |)."""
    nome_pulito = "".join(c for c in nome if c not in '\\/:*?"<>|').strip()
    return nome_pulito or "Senza_nome"


def _s21_anno_cartella_corrente() -> str:
    """Ritorna l'anno (come stringa) da usare come cartella principale
    dell'esportazione, cioè l'anno di fine dell'anno di servizio in corso
    oggi (es. per l'anno di servizio Settembre 2025 → Agosto 2026,
    ritorna '2026')."""
    oggi = datetime.now()
    anno_teo = anno_teocratico_di(f"{oggi.year}-{oggi.month:02d}")
    return str(anno_teo + 1) if anno_teo is not None else str(oggi.year)


def _s21_cartella_per_riga(riga: dict) -> str:
    """Determina il sotto-percorso di cartella (dentro 'Anno AAAA/') per un
    Proclamatore, secondo la struttura:
    Attivi/Proclamatori/<Gruppo>, Attivi/Pionieri Regolari,
    Attivi/Pionieri Speciali, Attivi/Missionari sul campo, Inattivi,
    Trasferiti."""
    stato = categoria_stato_proclamatore(riga.get("Attivi / Inattivi", ""))
    if stato == "I":
        return "Inattivi"
    if stato == "TR":
        return "Trasferiti"

    tipo = (riga.get("Tipo") or "").strip()
    if tipo == "Pioniere Regolare":
        return "Attivi/Pionieri Regolari"
    if tipo == "Pioniere speciale":
        return "Attivi/Pionieri Speciali"
    if tipo == "Missionario sul campo":
        return "Attivi/Missionari sul campo"

    gruppo = (riga.get("Gruppo") or "").strip() or "(Senza gruppo)"
    return f"Attivi/Proclamatori/{gruppo}"


def genera_zip_s21(righe_anagrafica: list, df_tutti: pd.DataFrame, anno_corrente: int) -> bytes:
    """Genera uno ZIP con una scheda S-21 per ciascun Proclamatore in
    'righe_anagrafica', organizzate secondo la struttura di cartelle
    'Anno AAAA/Attivi/…', 'Anno AAAA/Inattivi/…' ecc., col nome del file
    uguale al Cognome e Nome. 'anno_corrente' è l'anno teocratico scelto
    dall'utente (es. 2025): la cartella principale sarà 'Anno 2026' (l'anno
    in cui l'anno di servizio 2025-2026 termina). Pronto da scompattare
    dentro Dropbox o Drive."""
    anno_cartella = anno_corrente + 1
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for riga in righe_anagrafica:
            nome = (riga.get("Cognome e Nome") or "").strip()
            if not nome:
                continue
            pdf_bytes = genera_pdf_s21_singolo(riga, df_tutti, anno_corrente)
            sotto_cartella = _s21_cartella_per_riga(riga)
            nome_file = _s21_nome_file_sicuro(nome) + ".pdf"
            percorso = f"Anno {anno_cartella}/{sotto_cartella}/{nome_file}"
            zf.writestr(percorso, pdf_bytes)
    buf.seek(0)
    return buf.getvalue()


def genera_pdf_s21_riepilogo(titolo: str, nomi: list, df_tutti: pd.DataFrame, anno_corrente: int,
                              parola_chiave_tipo: str = None, etichetta_conteggio: str = "proclamatori") -> bytes:
    """Genera una cartolina S-21 di riepilogo aggregato per una categoria
    (es. 'Tutti i proclamatori', 'Tutti i pionieri regolari'): stessa
    grafica delle cartoline individuali, ma con conteggi e somme al posto
    dei dati di una singola persona.

    Se 'parola_chiave_tipo' è specificata (es. 'pioniere regolare'), 'nomi'
    è inteso come il bacino in cui cercare (di norma tutti gli Attivi) e per
    ogni mese si contano SOLO le righe del foglio Tutti il cui Tipo di
    servizio di quel mese contiene quella parola chiave (vedi
    _s21_righe_anno_aggregate_per_tipo) — non una lista fissa da Anagrafica.
    Se invece è None, si usa la lista fissa 'nomi' per tutti i mesi senza
    filtrare per Tipo (caso non usato di norma, tenuto per flessibilità).
    'etichetta_conteggio' è la parola usata nelle Osservazioni
    (es. '8 pionieri regolari')."""
    anno_precedente = anno_corrente - 1
    dati = _s21_dati_riepilogo(titolo)
    if parola_chiave_tipo:
        righe_corrente = _s21_righe_anno_aggregate_per_tipo(df_tutti, nomi, anno_corrente,
                                                              parola_chiave_tipo, etichetta_conteggio)
        righe_precedente = _s21_righe_anno_aggregate_per_tipo(df_tutti, nomi, anno_precedente,
                                                                parola_chiave_tipo, etichetta_conteggio)
    else:
        righe_corrente = _s21_righe_anno_aggregate(df_tutti, nomi, anno_corrente, etichetta_conteggio)
        righe_precedente = _s21_righe_anno_aggregate(df_tutti, nomi, anno_precedente, etichetta_conteggio)

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(S21_PAGE_W, S21_PAGE_H))
    _s21_disegna_pannello(c, 0.0, dati, righe_precedente, anno_teocratico=anno_precedente)
    _s21_disegna_pannello(c, S21_OFFSET_PANNELLO, dati, righe_corrente, anno_teocratico=anno_corrente)
    c.save()
    buf.seek(0)

    overlay_reader = PdfReader(buf)
    template_reader = PdfReader(PERCORSO_MODULO_S21)
    writer = PdfWriter()
    pagina = template_reader.pages[0]
    pagina.merge_page(overlay_reader.pages[0])
    writer.add_page(pagina)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def genera_zip_s21_completo(df: pd.DataFrame, df_tutti: pd.DataFrame, anno_corrente: int) -> bytes:
    """Genera il pacchetto ZIP completo e ufficiale: una cartolina per
    OGNI Proclamatore Attivo o Inattivo in Anagrafica (i Trasferiti sono
    esclusi), più le quattro cartoline di riepilogo aggregato (solo sugli
    Attivi) nella cartella principale. Non dipende da nessuna selezione: è
    pensato per essere sempre coerente e completo."""
    anno_cartella = anno_corrente + 1
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1) Una cartolina per ciascun Proclamatore Attivo o Inattivo (Trasferiti esclusi).
        for _, riga in df.iterrows():
            riga_dict = riga.to_dict()
            nome = (riga_dict.get("Cognome e Nome") or "").strip()
            if not nome:
                continue
            if categoria_stato_proclamatore(riga_dict.get("Attivi / Inattivi", "")) == "TR":
                continue
            pdf_bytes = genera_pdf_s21_singolo(riga_dict, df_tutti, anno_corrente)
            sotto_cartella = _s21_cartella_per_riga(riga_dict)
            nome_file = _s21_nome_file_sicuro(nome) + ".pdf"
            zf.writestr(f"Anno {anno_cartella}/{sotto_cartella}/{nome_file}", pdf_bytes)

        # 2) Cartoline di riepilogo aggregato, solo sugli Attivi, nella cartella principale.
        if "Attivi / Inattivi" in df.columns:
            categorie = df["Attivi / Inattivi"].apply(categoria_stato_proclamatore)
            df_attivi = df[categorie == "A"]
        else:
            df_attivi = df

        nomi_tutti_attivi = [str(n).strip() for n in df_attivi["Cognome e Nome"] if str(n).strip()]

        # Le 4 categorie "dinamiche" (Regolari, Speciali, Missionari, Ausiliari) non usano
        # Tutte e 5 le categorie sono "dinamiche": non usano una lista fissa presa dal
        # Tipo in Anagrafica, ma per ognuna, mese per mese, cercano nel foglio Tutti
        # (colonna D, Tipo di servizio) solo le righe di chi QUEL MESE risulta in quella
        # categoria — cercando tra tutti gli Attivi. Nessuna sovrapposizione tra cartoline:
        # ogni riga del foglio Tutti finisce in una sola categoria per quel mese.
        riepiloghi = [
            ("Tutti i proclamatori", nomi_tutti_attivi, "Riepilogo Tutti i Proclamatori.pdf",
             "proclamatore", "proclamatori"),
            ("Tutti i pionieri regolari", nomi_tutti_attivi, "Riepilogo Pionieri Regolari.pdf",
             "pioniere regolare", "pionieri regolari"),
            ("Tutti i pionieri speciali", nomi_tutti_attivi, "Riepilogo Pionieri Speciali.pdf",
             "pioniere speciale", "pionieri speciali"),
            ("Tutti i missionari sul campo", nomi_tutti_attivi, "Riepilogo Missionari sul Campo.pdf",
             "missionario|rappresentante", "missionari sul campo"),
            ("Tutti i pionieri ausiliari", nomi_tutti_attivi, "Riepilogo Pionieri Ausiliari.pdf",
             "pioniere ausiliario", "ausiliari"),
        ]
        for titolo, nomi, nome_file, parola_chiave_tipo, etichetta in riepiloghi:
            if not nomi:
                continue
            pdf_bytes = genera_pdf_s21_riepilogo(titolo, nomi, df_tutti, anno_corrente,
                                                  parola_chiave_tipo=parola_chiave_tipo,
                                                  etichetta_conteggio=etichetta)
            zf.writestr(f"Anno {anno_cartella}/{nome_file}", pdf_bytes)

    buf.seek(0)
    return buf.getvalue()


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


def elimina_riga_foglio(_workbook, nome_foglio: str, riga_da_eliminare: int):
    """Elimina una riga (numero 1-based) da un foglio. Ritorna
    (successo: bool, errore: str|None)."""
    try:
        ws = _workbook.worksheet(nome_foglio)
        ws.delete_rows(riga_da_eliminare)
        return True, None
    except Exception as e:
        return False, f"Errore durante l'eliminazione: {e}"


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
    st.markdown("## 📒 Gestione Registrazioni SEG")
    st.title("Pannello di controllo")

    st.subheader("Sezioni")
    card_data = [
        ("📖", "Rapporti consegnati", "Visualizza e modifica i rapporti di servizio consegnati.", "registrazioni"),
        ("📚", "Storico rapporti", "Storico dei rapporti di servizio per Proclamatore.", "storico"),
        ("🗂️", "Anagrafiche", "Gestisci i dati dei Proclamatori.", "anagrafiche"),
        ("📇", "Cartoline di registrazione", "Genera le cartoline S-21 per i Proclamatori scelti.", "cartoline"),
    ]
    riga1 = st.columns(2)
    riga2 = st.columns(2)
    colonne = riga1 + riga2
    for col, (icon, titolo, desc, pagina) in zip(colonne, card_data):
        with col:
            with st.container(border=True):
                st.markdown(f"#### {icon}  {titolo}")
                st.caption(desc)
                st.button("Apri →", key=f"card_{titolo}", disabled=not collegato,
                          use_container_width=True, on_click=vai_a, args=(pagina,))

    st.caption(f"Ultimo aggiornamento pagina: {datetime.now().strftime('%d/%m/%Y %H:%M')}")


# ─────────────────────────────────────────────────────────────────
# PAGINA: RAPPORTI CONSEGNATI
# ─────────────────────────────────────────────────────────────────
def _form_rapporto(df: pd.DataFrame, riga_esistente: dict, numero_riga_foglio: int,
                    chiave: str, chiave_stato_modifica: str = None):
    """Form generico di modifica per una riga del foglio 'Risposte del
    modulo 9': un campo per ciascuna colonna del foglio, precompilato con
    i valori attuali. I campi numerici (Ore, Cr. Ore, Studi) usano un
    campo numerico, gli altri un campo di testo. Alcune colonne (es.
    'Video mostrati') non vengono mostrate nel form ma il loro valore
    esistente viene comunque preservato al salvataggio. 'chiave' deve
    essere univoca per ogni istanza del form sulla stessa pagina."""
    e = riga_esistente
    colonne_numeriche = {"ore", "cr. ore", "cr ore", "studi"}
    colonne_nascoste = {"video mostrati", "cognome e nome"}

    with st.form(f"form_rapporto_{chiave}", clear_on_submit=False):
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
                valori_inseriti[colonna] = st.selectbox(colonna, opzioni, index=indice,
                                                         key=f"campo_{colonna}_{chiave}")
            elif chiave_norm in colonne_numeriche:
                try:
                    default_num = float(str(valore_attuale).replace(",", ".")) if valore_attuale else 0.0
                except ValueError:
                    default_num = 0.0
                valori_inseriti[colonna] = st.number_input(colonna, value=default_num, step=1.0,
                                                            key=f"campo_{colonna}_{chiave}")
            else:
                valori_inseriti[colonna] = st.text_input(colonna, value=str(valore_attuale),
                                                          key=f"campo_{colonna}_{chiave}")

        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            invia = st.form_submit_button("✔ Salva", type="primary", use_container_width=True)
        with col_btn2:
            annulla = st.form_submit_button("Annulla", use_container_width=True)

    if annulla:
        if chiave_stato_modifica:
            st.session_state[chiave_stato_modifica] = None
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
            if chiave_stato_modifica:
                st.session_state[chiave_stato_modifica] = None
            st.success("✔ Salvato correttamente.")
            st.rerun()
        else:
            st.error(err)


def _form_modifica_rapporto_consegnato(dati_selezione: dict):
    """Pagina intera di modifica di un rapporto consegnato (foglio
    'Risposte del modulo 9'), aperta cliccando 'Modifica' dentro la
    griglia di un Proclamatore in Rapporti consegnati."""
    df = dati_selezione["df"]
    riga_dict = dati_selezione["riga_dict"]
    numero_riga_foglio = dati_selezione["numero_riga_foglio"]
    nome = dati_selezione["nome"]

    st.title("Modifica rapporto")
    st.caption(f"{nome} (foglio «{NOME_FOGLIO_RISPOSTE}», riga {numero_riga_foglio})")

    _form_rapporto(df, riga_dict, numero_riga_foglio, chiave=str(numero_riga_foglio),
                    chiave_stato_modifica="rapporto_modifica_globale")


def mostra_registrazioni():
    if st.button("🏠 Torna alla Home", key="home_da_registrazioni", type="primary", use_container_width=True):
        vai_a("home")
        st.rerun()

    if not collegato:
        st.title("Rapporti consegnati")
        st.warning("⚠️  Nessun foglio dati collegato.")
        return

    if "rapporto_modifica_globale" not in st.session_state:
        st.session_state.rapporto_modifica_globale = None

    if st.session_state.rapporto_modifica_globale is not None:
        _form_modifica_rapporto_consegnato(st.session_state.rapporto_modifica_globale)
        return

    st.title("Rapporti consegnati")
    st.caption(f"Dati letti dal foglio «{NOME_FOGLIO_RISPOSTE}» (intestazione riga {RIGA_INTESTAZIONE_RISPOSTE}).")

    df_anagrafica, err_anagrafica = leggi_foglio_come_df(
        workbook, NOME_FOGLIO_ANAGRAFICA, RIGA_INTESTAZIONE_ANAGRAFICA)
    if err_anagrafica:
        st.error(err_anagrafica)
        return

    df, err = leggi_foglio_come_df(workbook, NOME_FOGLIO_RISPOSTE, RIGA_INTESTAZIONE_RISPOSTE)
    if err:
        st.error(err)
        return

    if df_anagrafica.empty or "Cognome e Nome" not in df_anagrafica.columns:
        st.info("Nessun Proclamatore trovato in Anagrafica.")
        return

    ricerca = st.text_input("🔍 Cerca per nome", placeholder="Digita per filtrare…")

    def e_attivo(valore: str) -> bool:
        """Vero se lo stato è 'A' oppure 'Attivi'/'Attivo' per esteso."""
        return (valore or "").strip().lower().startswith("a")

    colonna_stato = "Attivi / Inattivi" if "Attivi / Inattivi" in df_anagrafica.columns else None
    colonna_gruppo = "Gruppo" if "Gruppo" in df_anagrafica.columns else None
    stato_per_nome = {}
    gruppo_per_nome = {}
    if colonna_stato or colonna_gruppo:
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
        nomi = [n for n in nomi if e_attivo(stato_per_nome.get(n, ""))]
    if ricerca:
        nomi = [n for n in nomi if ricerca.lower() in n.lower()]

    conteggi = {}
    if "Cognome e Nome" in df.columns:
        for nome in nomi:
            conteggi[nome] = (df["Cognome e Nome"].astype(str).str.strip().str.lower() == nome.lower()).sum()
    else:
        for nome in nomi:
            conteggi[nome] = 0

    # ── Raggruppamento alfabetico per Gruppo (sorvegliante) ──────────────
    gruppi = {}
    for n in nomi:
        g = gruppo_per_nome.get(n, "") or "(Senza gruppo)"
        gruppi.setdefault(g, []).append(n)
    for g in gruppi:
        gruppi[g].sort()

    def _riga_proclamatore_rapporto(nome: str):
        conteggio = conteggi.get(nome, 0)
        pallino = "🟢" if conteggio == 1 else "🟡" if conteggio >= 2 else "🔴"

        with st.expander(f"{pallino}  {nome}"):
            if "Cognome e Nome" not in df.columns:
                righe_persona = df.iloc[0:0]
            else:
                righe_persona = df[df["Cognome e Nome"].astype(str).str.strip().str.lower() == nome.lower()]

            if righe_persona.empty:
                st.caption("Nessun rapporto consegnato per questo mese.")
                return

            colonne_tabella = [c for c in df.columns if c.strip().lower() != "cognome e nome"]
            evento_tabella = st.dataframe(
                righe_persona[colonne_tabella],
                hide_index=True,
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row",
                key=f"tabella_rapp_{nome}",
            )

            righe_sel = evento_tabella.selection.rows if evento_tabella and evento_tabella.selection else []
            idx_originale = None
            if righe_sel:
                idx_originale = righe_persona.index[righe_sel[0]]

            chiave_conferma_elim = f"rapp_elim_{nome}"

            col_mod, col_elim = st.columns(2)
            with col_mod:
                if st.button("✏️ Modifica riga selezionata", key=f"btn_mod_{nome}",
                             disabled=idx_originale is None, use_container_width=True):
                    numero_riga_foglio = RIGA_INTESTAZIONE_RISPOSTE + 1 + idx_originale
                    st.session_state.rapporto_modifica_globale = {
                        "df": df,
                        "riga_dict": df.loc[idx_originale].to_dict(),
                        "numero_riga_foglio": numero_riga_foglio,
                        "nome": nome,
                    }
                    st.rerun()
            with col_elim:
                if st.button("🗑️ Elimina riga selezionata", key=f"btn_elim_{nome}",
                             disabled=idx_originale is None, use_container_width=True):
                    st.session_state[chiave_conferma_elim] = True
                    st.rerun()

            if st.session_state.get(chiave_conferma_elim, False) and idx_originale is not None:
                numero_riga_foglio = RIGA_INTESTAZIONE_RISPOSTE + 1 + idx_originale
                st.warning("Confermi l'eliminazione di questo rapporto? L'operazione non è reversibile.")
                col_si, col_no = st.columns(2)
                with col_si:
                    if st.button("✔ Sì, elimina", key=f"btn_conf_si_{nome}",
                                 type="primary", use_container_width=True):
                        ok, err_elim = elimina_riga_foglio(workbook, NOME_FOGLIO_RISPOSTE, numero_riga_foglio)
                        if ok:
                            st.cache_data.clear()
                            st.session_state[chiave_conferma_elim] = False
                            st.success("✔ Rapporto eliminato.")
                            st.rerun()
                        else:
                            st.error(err_elim)
                with col_no:
                    if st.button("No, annulla", key=f"btn_conf_no_{nome}", use_container_width=True):
                        st.session_state[chiave_conferma_elim] = False
                        st.rerun()

                if len(righe_persona) > 1:
                    st.divider()

    for gruppo in sorted(gruppi.keys()):
        st.markdown(f"#### 👤 {gruppo}")
        for nome in gruppi[gruppo]:
            _riga_proclamatore_rapporto(nome)
        st.divider()


# ─────────────────────────────────────────────────────────────────
# PAGINA: ANAGRAFICHE
# ─────────────────────────────────────────────────────────────────
def _form_anagrafica(df: pd.DataFrame, riga_esistente: dict = None, numero_riga_foglio: int = None,
                      chiave: str = "nuovo", modo_nuovo: bool = False, chiave_expander: str = None):
    """Disegna il form di inserimento/modifica di un Proclamatore.
    Se 'riga_esistente' è None è un inserimento nuovo, altrimenti è una
    modifica (i campi vengono precompilati con i valori attuali).
    'chiave' deve essere univoca per ogni istanza del form sulla stessa
    pagina (es. l'ID del Proclamatore), per evitare conflitti quando più
    form sono presenti contemporaneamente (un riquadro per persona)."""
    e = riga_esistente or {}

    def parse_data(s):
        try:
            return datetime.strptime(s, "%d/%m/%Y").date()
        except Exception:
            return None

    with st.form(f"form_anagrafica_{chiave}", clear_on_submit=False):
        nome_cognome = st.text_input("Cognome e Nome *", value=e.get("Cognome e Nome", ""),
                                      key=f"nome_{chiave}")

        eta_nascita = calcola_eta_dettagliata(e.get("Data Nascita", ""))
        if eta_nascita:
            st.markdown(f"**Data di nascita** &nbsp; "
                        f"<span style='color:#D32F2F'>(anni {eta_nascita})</span>",
                        unsafe_allow_html=True)
        else:
            st.markdown("**Data di nascita**")
        data_nascita = st.date_input("Data di nascita", value=parse_data(e.get("Data Nascita", "")),
                                      format="DD/MM/YYYY", min_value=datetime(1900, 1, 1),
                                      label_visibility="collapsed", key=f"data_nascita_{chiave}")

        sesso_corrente = e.get("Sesso", "")
        sesso_default = ("Maschio" if sesso_corrente.upper().startswith("M")
                          else "Femmina" if sesso_corrente.upper().startswith("F") else "Maschio")
        sesso = st.selectbox("Sesso", OPZIONI_SESSO, index=OPZIONI_SESSO.index(sesso_default),
                              key=f"sesso_{chiave}")

        eta_battesimo = calcola_eta_dettagliata(e.get("Data Battesimo", ""))
        if eta_battesimo:
            st.markdown(f"**Data del battesimo** &nbsp; "
                        f"<span style='color:#D32F2F'>(anni {eta_battesimo})</span>",
                        unsafe_allow_html=True)
        else:
            st.markdown("**Data del battesimo**")
        data_battesimo = st.date_input("Data del battesimo", value=parse_data(e.get("Data Battesimo", "")),
                                        format="DD/MM/YYYY", min_value=datetime(1900, 1, 1),
                                        label_visibility="collapsed", key=f"data_batt_{chiave}")

        incarico_corrente = e.get("Incarico", "") or "(nessuno)"
        if incarico_corrente not in OPZIONI_INCARICO:
            incarico_corrente = "(nessuno)"
        incarico = st.selectbox("Incarico", OPZIONI_INCARICO,
                                 index=OPZIONI_INCARICO.index(incarico_corrente),
                                 key=f"incarico_{chiave}")

        tipo_corrente = e.get("Tipo", "") or "Proclamatore"
        if tipo_corrente not in OPZIONI_TIPO:
            tipo_corrente = "Proclamatore"
        tipo = st.selectbox("Tipo di servizio", OPZIONI_TIPO,
                             index=OPZIONI_TIPO.index(tipo_corrente), key=f"tipo_{chiave}")
        pr_dal = None
        if tipo in ("Pioniere Regolare", "Pioniere speciale", "Missionario sul campo"):
            pr_dal = st.date_input(f"{tipo} dal", value=parse_data(e.get("PR dal", "")),
                                    format="DD/MM/YYYY", min_value=datetime(1900, 1, 1),
                                    key=f"pr_dal_{chiave}")

        opzioni_gruppo = opzioni_da_colonna(df, "Gruppo")
        gruppo_corrente = e.get("Gruppo", "")
        elenco_gruppo = opzioni_gruppo + ["➕ Nuovo…"]
        if gruppo_corrente and gruppo_corrente not in elenco_gruppo:
            elenco_gruppo = [gruppo_corrente] + elenco_gruppo
        scelta_gruppo = st.selectbox("Gruppo", elenco_gruppo or ["➕ Nuovo…"],
                                      index=(elenco_gruppo.index(gruppo_corrente)
                                             if gruppo_corrente in elenco_gruppo else 0),
                                      key=f"gruppo_{chiave}")
        if scelta_gruppo == "➕ Nuovo…":
            scelta_gruppo = st.text_input("Nome del nuovo gruppo", key=f"gruppo_nuovo_{chiave}")

        opzioni_au = opzioni_da_colonna(df, "A/U")
        au_corrente = e.get("A/U", "")
        elenco_au = opzioni_au + ["➕ Nuovo…"]
        if au_corrente and au_corrente not in elenco_au:
            elenco_au = [au_corrente] + elenco_au
        scelta_au = st.selectbox("A/U", elenco_au or ["➕ Nuovo…"],
                                  index=(elenco_au.index(au_corrente) if au_corrente in elenco_au else 0),
                                  key=f"au_{chiave}")
        if scelta_au == "➕ Nuovo…":
            scelta_au = st.text_input("Nuovo valore A/U", key=f"au_nuovo_{chiave}")

        note = st.text_area("Note", value=e.get("Note", ""), height=100, key=f"note_{chiave}")

        st.divider()
        st.caption("Promemoria regolarità (da aggiornare quando manca il rapporto mensile)")
        irregolare = st.checkbox("Irregolare", value=e.get("Irregolare", "").strip().upper() in ("X", "SI", "SÌ"),
                                  key=f"irregolare_{chiave}")
        irregolare_mesi = st.number_input("Irregolare da mesi", min_value=0, max_value=36, step=1,
                                           value=int(e.get("Irregolare da Mesi", 0) or 0),
                                           key=f"irregolare_mesi_{chiave}")
        attivi_inattivi_corrente = e.get("Attivi / Inattivi", "A") or "A"
        if attivi_inattivi_corrente not in OPZIONI_ATTIVI_INATTIVI:
            attivi_inattivi_corrente = "A"
        etichetta_stato = st.selectbox("Stato", list(ETICHETTE_ATTIVI_INATTIVI.values()),
                                        index=OPZIONI_ATTIVI_INATTIVI.index(attivi_inattivi_corrente),
                                        key=f"stato_{chiave}")
        attivi_inattivi = {v: k for k, v in ETICHETTE_ATTIVI_INATTIVI.items()}[etichetta_stato]
        dal = st.date_input("Inattivo Da", value=parse_data(e.get("Inattivo dal", "")),
                             format="DD/MM/YYYY", min_value=datetime(1900, 1, 1), key=f"dal_{chiave}")

        st.divider()

        opzioni_trasf = opzioni_da_colonna(df, "Trasf.")
        trasf_corrente = e.get("Trasf.", "")
        elenco_trasf = opzioni_trasf + ["➕ Nuovo…"]
        if trasf_corrente and trasf_corrente not in elenco_trasf:
            elenco_trasf = [trasf_corrente] + elenco_trasf
        scelta_trasf = st.selectbox("Trasf.", elenco_trasf or ["➕ Nuovo…"],
                                     index=(elenco_trasf.index(trasf_corrente)
                                            if trasf_corrente in elenco_trasf else 0),
                                     key=f"trasf_{chiave}")
        if scelta_trasf == "➕ Nuovo…":
            scelta_trasf = st.text_input("Nuovo valore Trasf.", key=f"trasf_nuovo_{chiave}")

        messaggio = st.text_input("Messaggio", value=e.get("Messaggio", ""), key=f"messaggio_{chiave}")

        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            invia = st.form_submit_button("✔ Salva", use_container_width=True, type="primary")
        with col_btn2:
            annulla = st.form_submit_button("Annulla", use_container_width=True)

    if annulla:
        if modo_nuovo:
            st.session_state.anagrafica_nuovo = False
        if chiave_expander:
            st.session_state[chiave_expander] = False
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
            "A/U": scelta_au,
            "Trasf.": scelta_trasf,
            "Messaggio": messaggio.strip(),
            "Note": note.strip(),
            "Irregolare": "X" if irregolare else "",
            "Irregolare da Mesi": str(irregolare_mesi) if irregolare else "",
            "Attivi / Inattivi": attivi_inattivi,
            "Inattivo dal": dal.strftime("%d/%m/%Y") if dal else "",
            "Anni Età": calcola_eta(data_nascita_str),
            "Anni Batt": calcola_eta(data_battesimo_str),
        }

        ok, err = salva_riga_anagrafica(workbook, valori, riga_da_aggiornare=numero_riga_foglio)
        if ok:
            st.cache_data.clear()
            if modo_nuovo:
                st.session_state.anagrafica_nuovo = False
            if chiave_expander:
                st.session_state[chiave_expander] = False
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

    df, err = leggi_foglio_come_df(workbook, NOME_FOGLIO_ANAGRAFICA, RIGA_INTESTAZIONE_ANAGRAFICA)
    if err:
        st.error(err)
        return

    # ── Form di inserimento nuovo Proclamatore ──────────────────────────
    if st.session_state.anagrafica_nuovo:
        st.subheader("➕ Nuovo Proclamatore")
        _form_anagrafica(df, chiave="nuovo", modo_nuovo=True)
        return

    # ── Elenco Proclamatori a riquadri pieghevoli ───────────────────────
    if st.button("➕ Nuovo Proclamatore", use_container_width=True, type="primary"):
        st.session_state.anagrafica_nuovo = True
        st.rerun()

    ricerca = st.text_input("🔍 Cerca per nome, gruppo, tipo…", placeholder="Digita per filtrare…")

    if df.empty:
        st.info("Il foglio è collegato correttamente ma non contiene ancora Proclamatori.")
        return

    df_mostrato = df.reset_index(drop=True)
    if ricerca:
        maschera = df_mostrato.apply(
            lambda riga: riga.astype(str).str.contains(ricerca, case=False, na=False).any(), axis=1
        )
        df_mostrato = df_mostrato[maschera]

    # ── Filtro per stato: Attivi / Inattivi / Trasferiti ────────────────
    if "Attivi / Inattivi" in df_mostrato.columns:
        categorie = df_mostrato["Attivi / Inattivi"].apply(categoria_stato_proclamatore)
    else:
        categorie = pd.Series(["A"] * len(df_mostrato), index=df_mostrato.index)

    conteggio_a = int((categorie == "A").sum())
    conteggio_i = int((categorie == "I").sum())
    conteggio_tr = int((categorie == "TR").sum())

    opzioni_stato = [
        f"🟢 Attivi ({conteggio_a})",
        f"🔺 Inattivi ({conteggio_i})",
        f"↔️ Trasferiti ({conteggio_tr})",
    ]
    scelta_stato = st.radio("Stato", opzioni_stato, index=0, horizontal=True,
                             label_visibility="collapsed", key="anagrafica_filtro_stato")
    codice_stato_scelto = {opzioni_stato[0]: "A", opzioni_stato[1]: "I", opzioni_stato[2]: "TR"}[scelta_stato]

    df_mostrato = df_mostrato[categorie == codice_stato_scelto]

    df_mostrato = df_mostrato.sort_values("Cognome e Nome") if "Cognome e Nome" in df_mostrato.columns else df_mostrato

    st.caption(f"{len(df_mostrato)} Proclamatori su {len(df)} totali. Tocca un nominativo per aprirne la scheda.")

    if "anagrafica_aperto" not in st.session_state:
        st.session_state.anagrafica_aperto = None

    st.markdown(
        """
        <style>
        div[data-testid="stButton"] > button[kind="secondary"] {
            justify-content: flex-start !important;
            text-align: left !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    for idx, riga in df_mostrato.iterrows():
        nome = riga.get("Cognome e Nome", "(senza nome)")
        numero_riga_foglio = RIGA_INTESTAZIONE_ANAGRAFICA + 1 + idx
        chiave_persona = str(riga.get("ID") or idx)
        aperto = (st.session_state.anagrafica_aperto == chiave_persona)
        freccia = "▼" if aperto else "▶"

        if st.button(f"{freccia}  {nome}", key=f"btn_anagrafica_{chiave_persona}", use_container_width=True):
            st.session_state.anagrafica_aperto = None if aperto else chiave_persona
            st.rerun()

        if aperto:
            with st.container(border=True):
                _form_anagrafica(df, riga_esistente=riga.to_dict(), numero_riga_foglio=numero_riga_foglio,
                                  chiave=chiave_persona, modo_nuovo=False,
                                  chiave_expander="anagrafica_aperto")


# ─────────────────────────────────────────────────────────────────
# PAGINA: CARTOLINE DI REGISTRAZIONE (S-21)
# ─────────────────────────────────────────────────────────────────
def mostra_cartoline_registrazione():
    if st.button("🏠 Torna alla Home", key="home_da_cartoline", type="primary", use_container_width=True):
        vai_a("home")
        st.rerun()
    st.title("📇 Cartoline di registrazione")
    st.caption("Scegli l'anno di servizio e i Proclamatori per cui generare la scheda S-21. "
               "Il triangolo rosso 🔺 indica i Proclamatori Inattivi.")

    if not collegato:
        st.warning("⚠️  Nessun foglio dati collegato.")
        return

    if not os.path.exists(PERCORSO_MODULO_S21):
        st.error("Modulo S-21 non trovato: metti il file «S-21_s-Mlt_I.pdf» nella stessa cartella di app.py.")
        return

    df, err = leggi_foglio_come_df(workbook, NOME_FOGLIO_ANAGRAFICA, RIGA_INTESTAZIONE_ANAGRAFICA)
    if err:
        st.error(err)
        return
    if df.empty or "Cognome e Nome" not in df.columns:
        st.info("Nessun Proclamatore trovato in Anagrafica.")
        return

    df_tutti, err_tutti = leggi_foglio_tutti(workbook)
    if err_tutti:
        st.error(err_tutti)
        return

    # ── Scelta dell'anno teocratico, come in Storico rapporti consegnati ──
    anni_presenti = anni_teocratici_per_menu(df_tutti)
    anno_scelto = st.selectbox(
        "Seleziona anno teocratico",
        anni_presenti,
        format_func=lambda a: f"{a} – {a + 1} (set {a} → ago {a + 1})",
    )
    st.caption(f"La prima cartolina riporterà **{anno_scelto - 1}-{anno_scelto}**, "
               f"la seconda **{anno_scelto}-{anno_scelto + 1}**.")

    st.divider()
    st.markdown("#### 🗂️ Pacchetto completo")
    st.caption("Genera in un colpo solo la cartolina di **ogni** Proclamatore Attivo o Inattivo "
               "più le 5 cartoline di riepilogo (Tutti i proclamatori, Pionieri Regolari, Pionieri Speciali, "
               "Missionari sul campo, Pionieri Ausiliari). Non dipende dalla selezione qui sotto: è il "
               "pacchetto ufficiale, sempre completo e coerente.")
    if st.button("🗂️ Crea tutti i PDF delle registrazioni", type="primary", use_container_width=True):
        with st.spinner("Genero il pacchetto completo…"):
            zip_completo = genera_zip_s21_completo(df, df_tutti, anno_scelto)
        st.session_state.cartoline_pacchetto_completo = zip_completo

    if st.session_state.get("cartoline_pacchetto_completo"):
        st.download_button(
            "⬇️ Scarica il pacchetto completo (ZIP)",
            data=st.session_state.cartoline_pacchetto_completo,
            file_name=f"Registrazioni_Complete_{anno_scelto + 1}.zip",
            mime="application/zip",
            key="download_pacchetto_completo",
            use_container_width=True,
        )
    st.divider()
    st.markdown("#### ✏️ Rigenera solo una selezione")
    st.caption("Utile per rigenerare rapidamente una o poche cartoline dopo una correzione. "
               "Non sostituisce il pacchetto completo qui sopra: se rigeneri qui, ricordati di "
               "aggiornare anche il pacchetto ufficiale, altrimenti lo ZIP già distribuito resta disallineato.")

    df_lista = df.reset_index(drop=True)
    if "Attivi / Inattivi" in df_lista.columns:
        categorie = df_lista["Attivi / Inattivi"].apply(categoria_stato_proclamatore)
        df_lista = df_lista[categorie.isin(["A", "I"])]
        stato_per_nome = dict(zip(df_lista["Cognome e Nome"].astype(str).str.strip(),
                                   categorie[df_lista.index]))
    else:
        stato_per_nome = {}
    df_lista = df_lista[df_lista["Cognome e Nome"].astype(str).str.strip() != ""]

    ricerca = st.text_input("🔍 Cerca per nome", placeholder="Digita per filtrare…")
    df_mostrato = df_lista
    if ricerca:
        df_mostrato = df_mostrato[df_mostrato["Cognome e Nome"].astype(str).str.contains(ricerca, case=False, na=False)]
    df_mostrato = df_mostrato.sort_values("Cognome e Nome")

    nomi_tutti = [str(n).strip() for n in df_lista["Cognome e Nome"] if str(n).strip()]
    nomi_visibili = [str(n).strip() for n in df_mostrato["Cognome e Nome"] if str(n).strip()]

    def _chiave_cb(nome: str) -> str:
        return f"cb_cartolina_{nome}"

    col_tutti, col_nessuno = st.columns(2)
    with col_tutti:
        if st.button("☑️ Seleziona tutti (visibili)", use_container_width=True):
            for nome in nomi_visibili:
                st.session_state[_chiave_cb(nome)] = True
            st.rerun()
    with col_nessuno:
        if st.button("⬜ Deseleziona tutti", use_container_width=True):
            for nome in nomi_tutti:
                st.session_state[_chiave_cb(nome)] = False
            st.rerun()

    selezionati = [nome for nome in nomi_tutti if st.session_state.get(_chiave_cb(nome), False)]
    st.caption(f"{len(selezionati)} selezionati su {len(df_lista)} Proclamatori totali (Attivi + Inattivi).")
    st.divider()

    for nome in nomi_visibili:
        etichetta = f"🔺 {nome}" if stato_per_nome.get(nome) == "I" else nome
        st.checkbox(etichetta, key=_chiave_cb(nome))

    st.divider()

    selezionati = [nome for nome in nomi_tutti if st.session_state.get(_chiave_cb(nome), False)]
    n_sel = len(selezionati)
    if st.button(f"📄 Genera {n_sel} cartolin{'a' if n_sel == 1 else 'e'} selezionat{'a' if n_sel == 1 else 'e'}",
                 type="primary", use_container_width=True, disabled=n_sel == 0):
        righe_sel = [r.to_dict() for _, r in df.iterrows()
                     if str(r.get("Cognome e Nome", "")).strip() in selezionati]
        with st.spinner("Genero le cartoline…"):
            if len(righe_sel) == 1:
                pdf_bytes = genera_pdf_s21_singolo(righe_sel[0], df_tutti, anno_scelto)
                st.session_state.cartoline_pronto = ("pdf", pdf_bytes, righe_sel[0].get("Cognome e Nome", ""))
            else:
                zip_bytes = genera_zip_s21(righe_sel, df_tutti, anno_scelto)
                st.session_state.cartoline_pronto = ("zip", zip_bytes, anno_scelto)

    pronto = st.session_state.get("cartoline_pronto")
    if pronto:
        tipo, dati_file, extra = pronto
        if tipo == "pdf":
            st.download_button("⬇️ Scarica PDF", data=dati_file,
                                file_name=f"{_s21_nome_file_sicuro(extra)}.pdf",
                                mime="application/pdf", key="download_cartolina_pdf", use_container_width=True)
        else:
            st.download_button("⬇️ Scarica ZIP", data=dati_file,
                                file_name=f"Schede_S21_{extra + 1}.zip",
                                mime="application/zip", key="download_cartolina_zip", use_container_width=True)


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

    # ── Elenco anni teocratici disponibili (dati presenti + anno corrente/successivo) ──
    anni_presenti = anni_teocratici_per_menu(df_tutti)

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
elif st.session_state.pagina == "cartoline":
    mostra_cartoline_registrazione()
else:
    mostra_home()