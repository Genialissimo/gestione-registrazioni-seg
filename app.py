with st.expander("📊 Riepilogo attività"):
            st.caption("Report libero (non la scheda S-21): un elenco con mese, tipo di servizio, ore, "
                       "crediti, studi e note per ciascun Proclamatore, con totali e medie. Utile da "
                       "spedire ai sorveglianti di gruppo.")

            periodo_scelto = st.radio("Periodo", ["Tutto lo storico", "6 mesi"], horizontal=True,
                                      key="riepilogo_periodo")

            tipo_vista = st.radio("Tipo", ["Dettagliato", "Sintetico", "Tutti i grupos (comparati)" if "Tutti i grupos..." else "Tutti i gruppi (comparati)"], horizontal=True,
                                   key="riepilogo_tipo_vista",
                                   help="Dettagliato: un blocco per ciascun Proclamatore. "
                                        "Sintetico: un unico blocco con i totali della categoria scelta. "
                                        "Tutti i gruppi (comparati): tabella comparativa divisa per gruppo sotto ogni categoria.")

            gruppi_disponibili = ["Tutti i gruppi"]
            if "Gruppo" in df.columns:
                gruppi_disponibili += sorted({g.strip() for g in df["Gruppo"].astype(str) if g.strip()})
            gruppo_scelto = st.selectbox("Gruppo", gruppi_disponibili, key="riepilogo_gruppo")

            categoria_scelta = st.selectbox("Categoria", list(CATEGORIE_RIEPILOGO_ATTIVITA.keys()),
                                             key="riepilogo_categoria")

            if st.button("📄 Crea PDF", key="riepilogo_crea_pdf", use_container_width=True):
                with st.spinner("Genero il riepilogo…"):
                    if tipo_vista == "Tutti i gruppi (comparati)":
                        # Filtriamo solo per periodo, ignorando il gruppo e la singola categoria (li elabora tutti la funzione)
                        df_periodo = _riepilogo_filtra_dati(df_tutti, df, periodo_scelto, "Tutti i gruppi", "Tutti")
                        blocchi_comparati = _riepilogo_costruisci_blocchi_comparati(df_periodo, df)
                        trovato_qualcosa = bool(blocchi_comparati)
                        pdf_bytes = genera_pdf_riepilogo_attivita(
                            blocchi=[],
                            etichetta_periodo=periodo_scelto,
                            etichetta_categoria="Tutti",
                            etichetta_vista=tipo_vista,
                            blocchi_comparati=blocchi_comparati
                        )
                    elif tipo_vista == "Sintetico" and categoria_scelta == "Tutti":
                        df_periodo_gruppo = _riepilogo_filtra_dati(df_tutti, df, periodo_scelto,
                                                                   gruppo_scelto, "Tutti")
                        totali_categoria = _riepilogo_totali_generali_per_categoria(df_periodo_gruppo)
                        trovato_qualcosa = bool(totali_categoria)
                        pdf_bytes = genera_pdf_riepilogo_attivita(
                            [], periodo_scelto, categoria_scelta,
                            gruppo_scelto if gruppo_scelto != "Tutti i gruppi" else None,
                            etichetta_vista=tipo_vista, totali_per_categoria=totali_categoria,
                        )
                    else:
                        df_filtrato = _riepilogo_filtra_dati(df_tutti, df, periodo_scelto, gruppo_scelto,
                                                             categoria_scelta)
                        if tipo_vista == "Sintetico":
                            blocchi = _riepilogo_costruisci_blocco_sintetico(df_filtrato, categoria_scelta)
                        else:
                            blocchi = _riepilogo_costruisci_blocchi_dettagliato(df_filtrato)
                        trovato_qualcosa = bool(blocchi)
                        pdf_bytes = genera_pdf_riepilogo_attivita(
                            blocchi, periodo_scelto, categoria_scelta,
                            gruppo_scelto if gruppo_scelto != "Tutti i gruppi" else None,
                            etichetta_vista=tipo_vista,
                        )
                if not trovato_qualcosa:
                    st.warning("Nessun dato trovato per i filtri selezionati — il PDF generato sarà vuoto.")
                st.session_state.riepilogo_pdf_pronto = pdf_bytes

            if st.session_state.get("riepilogo_pdf_pronto"):
                nome_file = "Riepilogo_Attivita"
                if gruppo_scelto != "Tutti i gruppi":
                    nome_file += f"_{_s21_nome_file_sicuro(gruppo_scelto)}"
                nome_file += f"_{tipo_vista.replace(' ', '_').replace('(', '').replace(')', '')}_{categoria_scelta.replace(' ', '_')}.pdf"
                st.download_button(
                    "⬇️ Scarica Riepilogo attività (PDF)",
                    data=st.session_state.riepilogo_pdf_pronto,
                    file_name=nome_file,
                    mime="application/pdf",
                    key="download_riepilogo_attivita",
                    use_container_width=True,
                    on_click=lambda: st.session_state.pop("riepilogo_pdf_pronto", None),
                )
