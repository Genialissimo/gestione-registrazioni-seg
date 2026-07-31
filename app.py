# ─────────────────────────────────────────────────────────────────
# SEZIONE INTERFACCIA STREAMLIT: RIEPILOGO ATTIVITÀ
# ─────────────────────────────────────────────────────────────────

with st.expander("📊 Riepilogo attività", expanded=False):
    st.markdown(
        "Report libero (non la scheda S-21): un elenco con mese, "
        "tipo di servizio, ore, crediti, studi e note per ciascun "
        "Proclamatore, con totali e medie. Utile da spedire ai "
        "sorveglianti di gruppo."
    )
    
    # Selezione Periodo
    st.markdown("**Periodo**")
    periodo_scelto = st.radio(
        "Periodo", 
        ["12 mesi", "6 mesi"], 
        horizontal=True, 
        label_visibility="collapsed",
        key="radio_periodo_riepilogo"
    )
    
    # Selezione Tipo (Aggiornato con "Sintetico compara gruppi")
    st.markdown("**Tipo**")
    tipo_report_scelto = st.radio(
        "Tipo", 
        ["Dettagliato", "Sintetico", "Sintetico compara gruppi"], 
        horizontal=True, 
        label_visibility="collapsed",
        key="radio_tipo_riepilogo"
    )
    
    # Estrazione dei gruppi disponibili (se presenti in anagrafica)
    lista_gruppi = ["Tutti i gruppi"]
    if 'df_anagrafica' in locals() and not df_anagrafica.empty and "Gruppo" in df_anagrafica.columns:
        gruppi_unici = sorted(list(set(df_anagrafica["Gruppo"].dropna().astype(str).str.strip()) - {""}))
        lista_gruppi.extend(gruppi_unici)

    # Selezione Gruppo
    st.markdown("**Gruppo**")
    gruppo_scelto = st.selectbox(
        "Gruppo", 
        lista_gruppi, 
        label_visibility="collapsed",
        key="select_gruppo_riepilogo"
    )
    
    # Selezione Categoria
    st.markdown("**Categoria**")
    categoria_scelta = st.selectbox(
        "Categoria", 
        list(CATEGORIE_RIEPILOGO_ATTIVITA.keys()), 
        label_visibility="collapsed",
        key="select_categoria_riepilogo"
    )
    
    st.markdown("") # Spaziatura
    
    # Pulsante di generazione ed esportazione PDF
    if st.button("📄 Crea PDF", key="btn_crea_pdf_riepilogo"):
        # 1. Preparazione dei dati tramite la funzione di coordinamento
        blocchi, totali_cat, totali_compara = riepilogo_prepara_dati_visivi(
            df_tutti=df_tutti, 
            df_anagrafica=df_anagrafica, 
            periodo=periodo_scelto, 
            gruppo_scelto=gruppo_scelto, 
            categoria=categoria_scelta, 
            tipo_report=tipo_report_scelto
        )
        
        # 2. Generazione del PDF in formato bytes
        pdf_bytes = genera_pdf_riepilogo_attivita(
            blocchi=blocchi,
            etichetta_periodo=periodo_scelto,
            etichetta_categoria=categoria_scelta,
            etichetta_gruppo=gruppo_scelto if gruppo_scelto != "Tutti i gruppi" else None,
            etichetta_vista=tipo_report_scelto,
            totali_per_categoria=totali_cat,
            totali_compara_gruppi=totali_compara
        )
        
        # 3. Pulsante per scaricare il file PDF generato
        st.download_button(
            label="📥 Scarica PDF Riepilogo Attività",
            data=pdf_bytes,
            file_name=f"Riepilogo_Attivita_{tipo_report_scelto.replace(' ', '_')}.pdf",
            mime="application/pdf",
            key="download_pdf_riepilogo"
        )
