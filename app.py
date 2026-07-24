# ─────────────────────────────────────────────────────────────────
# RIEPILOGO ATTIVITÀ (report libero per sorveglianti di gruppo/categoria)
# ─────────────────────────────────────────────────────────────────
CATEGORIE_RIEPILOGO_ATTIVITA = {
    "Tutti": None,
    "Solo Proclamatori": "proclamatore",
    "Solo Pionieri Ausiliari": "pioniere ausiliario",
    "Solo Pionieri Regolari": "pioniere regolare",
    "Solo Pionieri Speciali": "pioniere speciale",
    "Solo Missionari sul campo": "missionario|rappresentante",
}


def _riepilogo_ultimo_mese_con_dati(df_tutti: pd.DataFrame):
    """Ritorna (anno, mese) del mese più recente per cui esiste almeno un
    rapporto nel foglio Tutti, o None se il foglio è vuoto."""
    if df_tutti.empty or "Mese/Anno" not in df_tutti.columns:
        return None
    validi = []
    for m in df_tutti["Mese/Anno"].dropna().unique():
        try:
            a, mm = str(m).split("-")
            validi.append((int(a), int(mm)))
        except Exception:
            continue
    if not validi:
        return None
    return max(validi)


def _riepilogo_finestra_ultimi_n_mesi(anno_fine: int, mese_fine: int, n: int = 6) -> set:
    """Ritorna l'insieme (anno, mese) degli ultimi 'n' mesi, contando a
    ritroso da (anno_fine, mese_fine) incluso."""
    risultato = set()
    a, m = anno_fine, mese_fine
    for _ in range(n):
        risultato.add((a, m))
        m -= 1
        if m == 0:
            m = 12
            a -= 1
    return risultato


ETICHETTE_SINTETICO_CATEGORIA = {
    "Tutti": "Tutti i proclamatori",
    "Solo Proclamatori": "Proclamatori",
    "Solo Pionieri Ausiliari": "Pionieri Ausiliari",
    "Solo Pionieri Regolari": "Pionieri Regolari",
    "Solo Pionieri Speciali": "Pionieri Speciali",
    "Solo Missionari sul campo": "Missionari sul campo",
}


def _riepilogo_filtra_dati(df_tutti: pd.DataFrame, df_anagrafica: pd.DataFrame, periodo: str,
                            gruppo_scelto: str, categoria: str) -> pd.DataFrame:
    """Filtra il foglio Tutti per periodo, categoria (Tipo di servizio di
    quel mese) ed eventualmente Gruppo (assegnazione attuale in Anagrafica
    — è l'unico filtro che dipende da Anagrafica, perché il Gruppo non
    esiste nel foglio Tutti). Usata sia dalla vista Dettagliato che da
    quella Sintetico del Riepilogo attività."""
    df = df_tutti.copy()
    if df.empty:
        return df

    if periodo == "6 mesi":
        ultimo = _riepilogo_ultimo_mese_con_dati(df_tutti)
        if not ultimo:
            return df.iloc[0:0]
        finestra = _riepilogo_finestra_ultimi_n_mesi(ultimo[0], ultimo[1], 6)

        def _dentro_finestra(mese_anno):
            try:
                a, m = str(mese_anno).split("-")
                return (int(a), int(m)) in finestra
            except Exception:
                return False

        df = df[df["Mese/Anno"].apply(_dentro_finestra)]

    parola_categoria = CATEGORIE_RIEPILOGO_ATTIVITA.get(categoria)
    if parola_categoria:
        df = df[df["Tipo Servizio"].str.lower().str.contains(parola_categoria, na=False, regex=True)]

    if gruppo_scelto and gruppo_scelto != "Tutti i gruppi" and "Gruppo" in df_anagrafica.columns:
        nomi_gruppo = set(
            df_anagrafica.loc[df_anagrafica["Gruppo"].astype(str).str.strip() == gruppo_scelto, "Cognome e Nome"]
            .astype(str).str.strip()
        )
        df = df[df["Nome"].str.strip().isin(nomi_gruppo)]

    return df


def _riepilogo_costruisci_blocchi_dettagliato(df_filtrato: pd.DataFrame) -> list:
    """Un blocco per ciascun Proclamatore presente in 'df_filtrato' (già
    filtrato per periodo/categoria/gruppo). Ogni blocco: {'nome', 'righe',
    'totale_ore', 'totale_crediti', 'totale_studi', 'media_ore',
    'media_crediti', 'media_studi'}."""
    if df_filtrato.empty:
        return []

    blocchi = []
    for nome, gruppo_df in df_filtrato.groupby("Nome"):
        gruppo_df = gruppo_df.sort_values("Mese/Anno")
        righe = []
        tot_ore = tot_cred = tot_studi = 0.0
        for _, r in gruppo_df.iterrows():
            ore_val = a_float_it(r.get("Ore", ""))
            cred_val = a_float_it(r.get("Cred. Ore", ""))
            studi_val = a_float_it(r.get("Studi Biblici", ""))
            tot_ore += ore_val
            tot_cred += cred_val
            tot_studi += studi_val
            righe.append({
                "mese_anno": r.get("Mese/Anno", ""),
                "tipo": r.get("Tipo Servizio", "") or "",
                "ministero": "Sì" if r.get("Ha partecipato al ministero") else "No",
                "ore": formatta_numero_it(ore_val) if ore_val else "",
                "crediti": formatta_numero_it(cred_val) if cred_val else "",
                "studi": formatta_numero_it(studi_val) if studi_val else "",
                "note": r.get("Osservazioni", "") or "",
            })
        n = len(righe)
        blocchi.append({
            "nome": nome,
            "righe": righe,
            "totale_ore": tot_ore,
            "totale_crediti": tot_cred,
            "totale_studi": tot_studi,
            "media_ore": tot_ore / n if n else 0.0,
            "media_crediti": tot_cred / n if n else 0.0,
            "media_studi": tot_studi / n if n else 0.0,
        })
    blocchi.sort(key=lambda b: b["nome"])
    return blocchi


def _riepilogo_costruisci_blocco_sintetico(df_filtrato: pd.DataFrame, categoria: str) -> list:
    """Un UNICO blocco, intestato con il nome della categoria (es. 'Pionieri
    Ausiliari') invece che con un nome di persona: raggruppa 'df_filtrato'
    (già filtrato per periodo/categoria/gruppo) per mese, sommando
    Ore/Crediti/Studi e mettendo nelle Note il conteggio di quante persone
    quel mese risultano in quella categoria. Totale e Media in fondo, come
    nella vista Dettagliato."""
    if df_filtrato.empty:
        return []

    etichetta = ETICHETTE_SINTETICO_CATEGORIA.get(categoria, categoria)
    righe = []
    tot_ore = tot_cred = tot_studi = 0.0
    for mese_anno, gruppo_mese in df_filtrato.groupby("Mese/Anno"):
        ore_val = sum(a_float_it(v) for v in gruppo_mese.get("Ore", []))
        cred_val = sum(a_float_it(v) for v in gruppo_mese.get("Cred. Ore", []))
        studi_val = sum(a_float_it(v) for v in gruppo_mese.get("Studi Biblici", []))
        conteggio = len(gruppo_mese)
        tot_ore += ore_val
        tot_cred += cred_val
        tot_studi += studi_val
        righe.append({
            "mese_anno": mese_anno,
            "tipo": etichetta,
            "ministero": "",
            "ore": formatta_numero_it(ore_val) if ore_val else "",
            "crediti": formatta_numero_it(cred_val) if cred_val else "",
            "studi": formatta_numero_it(studi_val) if studi_val else "",
            "note": f"{conteggio} {etichetta.lower()}" if conteggio else "",
        })

    def _chiave_ordinamento(riga):
        try:
            a, m = str(riga["mese_anno"]).split("-")
            return (int(a), int(m))
        except Exception:
            return (0, 0)

    righe.sort(key=_chiave_ordinamento)
    n_mesi = len(righe)
    return [{
        "nome": etichetta,
        "righe": righe,
        "totale_ore": tot_ore,
        "totale_crediti": tot_cred,
        "totale_studi": tot_studi,
        "media_ore": tot_ore / n_mesi if n_mesi else 0.0,
        "media_crediti": tot_cred / n_mesi if n_mesi else 0.0,
        "media_studi": tot_studi / n_mesi if n_mesi else 0.0,
    }]


def _riepilogo_totali_generali_per_categoria(df_periodo_gruppo: pd.DataFrame) -> list:
    """Usata quando Tipo='Sintetico' e Categoria='Tutti': niente righe
    mensili, solo il gran totale (e la media) per ciascuna delle categorie
    (Proclamatori, Pionieri Regolari, Pionieri Speciali, Missionari sul
    campo, Pionieri Ausiliari) trovate in 'df_periodo_gruppo' — già
    filtrato per periodo ed eventuale Gruppo, ma NON per categoria. Salta
    le categorie senza nessuna riga. Ritorna una lista di dict:
    {'categoria', 'totale_ore', 'totale_crediti', 'totale_studi',
    'media_ore', 'media_crediti', 'media_studi', 'conteggio'}."""
    if df_periodo_gruppo.empty:
        return []
 
    risultati = []
    for chiave, parola in CATEGORIE_RIEPILOGO_ATTIVITA.items():
        if chiave == "Tutti" or not parola:
            continue
        df_cat = df_periodo_gruppo[df_periodo_gruppo["Tipo Servizio"].str.lower().str.contains(
            parola, na=False, regex=True)]
        if df_cat.empty:
            continue
        tot_ore = sum(a_float_it(v) for v in df_cat.get("Ore", []))
        tot_cred = sum(a_float_it(v) for v in df_cat.get("Cred. Ore", []))
        tot_studi = sum(a_float_it(v) for v in df_cat.get("Studi Biblici", []))
        n = len(df_cat)
        risultati.append({
            "categoria": ETICHETTE_SINTETICO_CATEGORIA.get(chiave, chiave),
            "totale_ore": tot_ore,
            "totale_crediti": tot_cred,
            "totale_studi": tot_studi,
            "media_ore": tot_ore / n if n else 0.0,
            "media_crediti": tot_cred / n if n else 0.0,
            "media_studi": tot_studi / n if n else 0.0,
            "conteggio": n,
        })
    return risultati


def _riepilogo_costruisci_blocchi_comparati(df_periodo: pd.DataFrame, df_anagrafica: pd.DataFrame) -> list:
    """Nuova funzione: Usata per la vista 'Tutti i gruppi (comparati)'.
    Esegue il merge con Anagrafica per recuperare il Gruppo di appartenenza, 
    quindi raggruppa per Categoria e per Gruppo, calcolando i totali."""
    if df_periodo.empty or df_anagrafica.empty:
        return []
        
    df = df_periodo.merge(df_anagrafica[["Nome", "Gruppo"]], on="Nome", how="left")
    df["Gruppo"] = df["Gruppo"].fillna("Nessun Gruppo")
    
    risultati = []
    
    for chiave, parola in CATEGORIE_RIEPILOGO_ATTIVITA.items():
        if chiave == "Tutti" or not parola:
            continue
            
        df_cat = df[df["Tipo Servizio"].str.lower().str.contains(parola, na=False, regex=True)]
        if df_cat.empty:
            continue
            
        gruppi_stats = []
        for gruppo_nome, gruppo_df in df_cat.groupby("Gruppo"):
            tot_ore = sum(a_float_it(v) for v in gruppo_df.get("Ore", []))
            tot_cred = sum(a_float_it(v) for v in gruppo_df.get("Cred. Ore", []))
            tot_studi = sum(a_float_it(v) for v in gruppo_df.get("Studi Biblici", []))
            conteggio = len(gruppo_df)
            
            gruppi_stats.append({
                "gruppo": gruppo_nome,
                "totale_ore": tot_ore,
                "totale_crediti": tot_cred,
                "totale_studi": tot_studi,
                "media_ore": tot_ore / conteggio if conteggio else 0.0,
                "media_crediti": tot_cred / conteggio if conteggio else 0.0,
                "media_studi": tot_studi / conteggio if conteggio else 0.0,
            })
            
        gruppi_stats.sort(key=lambda x: str(x["gruppo"]))
        
        risultati.append({
            "categoria": ETICHETTE_SINTETICO_CATEGORIA.get(chiave, chiave),
            "gruppi": gruppi_stats
        })
        
    return risultati
 
 
def genera_pdf_riepilogo_attivita(blocchi: list, etichetta_periodo: str, etichetta_categoria: str,
                                   etichetta_gruppo: str = None, etichetta_vista: str = "Dettagliato",
                                   totali_per_categoria: list = None, blocchi_comparati: list = None) -> bytes:
    """Genera il PDF del Riepilogo attività: un documento libero (non un
    modulo prestampato) con un blocco per Proclamatore (vista Dettagliato)
    o un unico blocco per categoria (vista Sintetico) — mese, tipo di
    servizio, ministero, ore, crediti, studi, note — più Totale e Media,
    con interruzioni di pagina automatiche.
 
    Se 'totali_per_categoria' è specificato (caso Sintetico + Categoria
    'Tutti'), ignora 'blocchi' e mostra invece solo il gran totale e la
    media per ciascuna categoria, senza righe mensili.
    
    Se 'blocchi_comparati' è specificato, genera la vista con le tabelle 
    divise per Gruppo sotto ogni Categoria."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm,
                             leftMargin=1.3 * cm, rightMargin=1.3 * cm)
    stili = getSampleStyleSheet()
    elementi = []
 
    elementi.append(Paragraph("Attività dei proclamatori", stili["Title"]))
    sottotitolo = (f"Congregazione: {NOME_CONGREGAZIONE} · Periodo: {etichetta_periodo} · "
                   f"{etichetta_vista} - {etichetta_categoria}")
    if etichetta_gruppo:
        sottotitolo += f" · Gruppo: {etichetta_gruppo}"
    elementi.append(Paragraph(sottotitolo, stili["Normal"]))
    elementi.append(Spacer(1, 14))
 
    # NUOVO BLOCCO: Tutti i gruppi (comparati)
    if blocchi_comparati is not None:
        if not blocchi_comparati:
            elementi.append(Paragraph("Nessun dato trovato per i filtri selezionati.", stili["Normal"]))
        else:
            for cat_data in blocchi_comparati:
                elementi.append(Paragraph(f"<b>{cat_data['categoria']}</b>", stili["Heading3"]))
                elementi.append(Spacer(1, 8))
                
                for grp in cat_data["gruppi"]:
                    elementi.append(Paragraph(f"<b>Gruppo: {grp['gruppo']}</b>", stili["Normal"]))
                    elementi.append(Spacer(1, 4))
                    
                    dati_tabella = [
                        ["Totale Ore", "Media Ore", "Totale Studi", "Totale Crediti"],
                        [formatta_numero_it(grp['totale_ore']), 
                         formatta_numero_it(grp['media_ore']), 
                         formatta_numero_it(grp['totale_studi']), 
                         formatta_numero_it(grp['totale_crediti'])]
                    ]
                    
                    tabella = Table(dati_tabella, colWidths=[3.5*cm, 3.5*cm, 3.5*cm, 3.5*cm])
                    tabella.setStyle(TableStyle([
                        ("FONTSIZE", (0, 0), (-1, -1), 9),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B6FA8")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F2F2F2")),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                        ("ALIGN", (0,0), (-1,-1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ]))
                    blocco_pdf = [tabella, Spacer(1, 10)]
                    elementi.append(KeepTogether(blocco_pdf))
                    
                elementi.append(Spacer(1, 14))
                
        doc.build(elementi)
        buf.seek(0)
        return buf.getvalue()

    # INIZIO BLOCCO MODIFICATO
    if totali_per_categoria is not None:
        if not totali_per_categoria:
            elementi.append(Paragraph("Nessun dato trovato per i filtri selezionati.", stili["Normal"]))
        else:
            intestazione = ["Mese", "Servizio", "Ministero", "Ore", "Crediti", "Studi", "Note"]
            larghezze = [1.9 * cm, 3.0 * cm, 1.7 * cm, 1.6 * cm, 1.6 * cm, 1.6 * cm, 4.7 * cm]
            
            for cat in totali_per_categoria:
                dati_tabella = [intestazione]
                dati_tabella.append([
                    "Totale", cat["categoria"], "", 
                    formatta_numero_it(cat["totale_ore"]),
                    formatta_numero_it(cat["totale_crediti"]),
                    formatta_numero_it(cat["totale_studi"]), 
                    f"{cat['categoria']} n. {cat['conteggio']}"
                ])
                dati_tabella.append([
                    "Media", cat["categoria"], "", 
                    formatta_numero_it(cat["media_ore"]),
                    formatta_numero_it(cat["media_crediti"]),
                    formatta_numero_it(cat["media_studi"]), 
                    ""
                ])
                
                tabella = Table(dati_tabella, colWidths=larghezze)
                tabella.setStyle(TableStyle([
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B6FA8")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),                
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica-Bold"),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F2F2F2")),
                    ("GRID", (0, 1), (-1, -1), 0.4, colors.grey),
                    ("LINEBEFORE", (1, 1), (1, -1), 0.6, colors.HexColor("#F2F2F2")),
                    ("LINEAFTER", (1, 1), (1, -1), 0.6, colors.HexColor("#F2F2F2")),
                    ("ALIGN", (3, 0), (5, -1), "RIGHT"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ]))
                elementi.append(tabella)
                elementi.append(Spacer(1, 16))

        doc.build(elementi)
        buf.seek(0)
        return buf.getvalue()
    # FINE BLOCCO MODIFICATO
 
    if not blocchi:
        elementi.append(Paragraph("Nessun dato trovato per i filtri selezionati.", stili["Normal"]))
 
    intestazione = ["Mese", "Servizio", "Ministero", "Ore", "Crediti", "Studi", "Note"]
    larghezze = [1.9 * cm, 3.0 * cm, 1.7 * cm, 1.6 * cm, 1.6 * cm, 1.6 * cm, 4.7 * cm]
 
    for blocco in blocchi:
        dati_tabella = [intestazione]
        for r in blocco["righe"]:
            dati_tabella.append([r["mese_anno"], r["tipo"], r["ministero"], r["ore"], r["crediti"],
                                  r["studi"], r["note"]])
        dati_tabella.append(["Totale", "", "", formatta_numero_it(blocco["totale_ore"]),
                              formatta_numero_it(blocco["totale_crediti"]),
                              formatta_numero_it(blocco["totale_studi"]), ""])
        dati_tabella.append(["Media", "", "", formatta_numero_it(blocco["media_ore"]),
                              formatta_numero_it(blocco["media_crediti"]),
                              formatta_numero_it(blocco["media_studi"]), ""])
 
        tabella = Table(dati_tabella, colWidths=larghezze, repeatRows=1)
        tabella.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B6FA8")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, -2), (-1, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, -2), (-1, -1), colors.HexColor("#F2F2F2")),
            ("GRID", (0, -2), (-1, -1), 0.4, colors.grey),
            ("LINEBEFORE", (1, -2), (1, -1), 0.6, colors.HexColor("#F2F2F2")),
            ("LINEAFTER", (1, -2), (1, -1), 0.6, colors.HexColor("#F2F2F2")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -3), [colors.white, colors.HexColor("#DCEEF9")]),
            ("ALIGN", (3, 0), (5, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        blocco_pdf = [Paragraph(f"<b>{blocco['nome']}</b>", stili["Heading3"]), tabella, Spacer(1, 16)]
        elementi.append(KeepTogether(blocco_pdf))
 
    doc.build(elementi)
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
