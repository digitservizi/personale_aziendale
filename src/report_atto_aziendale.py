"""
Report medici – confronto dotazione da Atto Aziendale vs personale in servizio.
"""

import pandas as pd
from openpyxl import Workbook

from src.stili_excel import (
    FILL_A, FILL_B,
    scrivi_titolo, scrivi_intestazioni, scrivi_riga_dati,
    auto_larghezza_colonne,
)
from src.caricamento_dati import (
    carica_dataframe, normalizza_colonne_personale,
    normalizza_colonne_pensionamenti,
)
from src.caricamento_xml import carica_medici_atto_aziendale


def genera_report_atto_aziendale(personale_file, pensionamenti_file,
                                  mapper_atto_aziendale, output_file,
                                  anno_analisi):
    """
    Genera un report XLSX della dirigenza medica confrontando
    la dotazione da atto aziendale con il personale effettivo.
    """
    print(f"\n{'=' * 70}")
    print("REPORT MEDICI - ATTO AZIENDALE vs PERSONALE IN SERVIZIO")
    print(f"{'=' * 70}\n")

    # Carica personale
    personale_df = carica_dataframe(personale_file)
    personale_df = normalizza_colonne_personale(personale_df)
    personale_df = personale_df[
        personale_df['DESC_NATURA'].str.upper() == "TEMPO INDETERMINATO"
    ]

    medici_df = personale_df[
        personale_df['PROFILO_RAGGRUPPATO'].str.upper() == 'DIRIGENTE MEDICO'
    ].copy()
    medici_df['DISC_UPPER'] = (
        medici_df['DESC_DISCIPLINE'].str.upper().str.strip()
    )

    # Pensionamenti
    pensionamenti_df = carica_dataframe(pensionamenti_file)
    pensionamenti_df = normalizza_colonne_pensionamenti(pensionamenti_df)

    if 'DT_CESSAZIONE' in medici_df.columns:
        medici_df = medici_df.rename(
            columns={'DT_CESSAZIONE': 'DT_CESSAZIONE_PERS'}
        )
    medici_df = pd.merge(
        medici_df,
        pensionamenti_df[['MATR.', 'DT_CESSAZIONE']],
        on='MATR.',
        how='left',
    )

    anni_pensionamento = [anno_analisi + 1, anno_analisi + 2, anno_analisi + 3]

    # Mapper atto aziendale
    mapper = carica_medici_atto_aziendale(mapper_atto_aziendale)
    print(f"Discipline da atto aziendale: {len(mapper)}")
    print(f"Dirigenti medici TI nel DB: {len(medici_df)}\n")

    # Costruisci il report
    righe_report = []
    discipline_db_mappate = set()
    totale_dotazione = 0
    totale_in_servizio = 0

    for disc in mapper:
        nome_atto = disc['nome_atto']
        dotazione = disc['dotazione']
        voci_db = disc['discipline_db']

        mask = medici_df['DISC_UPPER'].isin(voci_db)
        medici_disc = medici_df[mask]
        in_servizio = len(medici_disc)

        pens = {}
        for anno in anni_pensionamento:
            pens[anno] = medici_disc['DT_CESSAZIONE'].apply(
                lambda x, a=anno: (
                    1 if pd.to_datetime(x, errors='coerce').year == a else 0
                )
            ).sum()

        delta = in_servizio - dotazione
        discipline_db_mappate.update(voci_db)

        riga = {
            'DISCIPLINA_ATTO_AZIENDALE': nome_atto,
            'DOTAZIONE_ATTO': dotazione,
            'IN_SERVIZIO': in_servizio,
            'DELTA': delta,
        }
        for anno in anni_pensionamento:
            riga[f'PENSIONAMENTI_{anno}'] = pens[anno]

        pens_cumulati = 0
        for anno in anni_pensionamento:
            pens_cumulati += pens[anno]
            riga[f'PROIEZIONE_{anno}'] = in_servizio - pens_cumulati

        righe_report.append(riga)
        totale_dotazione += dotazione
        totale_in_servizio += in_servizio

        stato = "OK" if delta >= 0 else f"CARENZA {abs(delta)}"
        print(
            f"  {nome_atto:45s}  Atto: {dotazione:3d}  "
            f"Servizio: {in_servizio:3d}  Delta: {delta:+4d}  [{stato}]"
        )

    # Discipline non mappate
    tutte_disc_db = set(medici_df['DISC_UPPER'].dropna().unique())
    non_mappate = tutte_disc_db - discipline_db_mappate
    medici_non_mappati = medici_df[medici_df['DISC_UPPER'].isin(non_mappate)]

    if non_mappate:
        print(
            f"\n  --- Discipline fuori dall'atto aziendale "
            f"({len(medici_non_mappati)} medici) ---"
        )
        for disc_nm in sorted(non_mappate):
            n = len(medici_df[medici_df['DISC_UPPER'] == disc_nm])
            mask_nm = medici_df['DISC_UPPER'] == disc_nm
            medici_nm = medici_df[mask_nm]
            pens_nm = {}
            for anno in anni_pensionamento:
                pens_nm[anno] = medici_nm['DT_CESSAZIONE'].apply(
                    lambda x, a=anno: (
                        1 if pd.to_datetime(x, errors='coerce').year == a
                        else 0
                    )
                ).sum()

            riga = {
                'DISCIPLINA_ATTO_AZIENDALE': f'[FUORI ATTO] {disc_nm}',
                'DOTAZIONE_ATTO': '-',
                'IN_SERVIZIO': n,
                'DELTA': '-',
            }
            for anno in anni_pensionamento:
                riga[f'PENSIONAMENTI_{anno}'] = pens_nm[anno]
            pens_cum = 0
            for anno in anni_pensionamento:
                pens_cum += pens_nm[anno]
                riga[f'PROIEZIONE_{anno}'] = n - pens_cum
            righe_report.append(riga)

            print(f"  {disc_nm:45s}  Servizio: {n:3d}")

    # Totali
    print(
        f"\n  {'TOTALE DISCIPLINE DA ATTO':45s}  Atto: {totale_dotazione:3d}  "
        f"Servizio: {totale_in_servizio:3d}  "
        f"Delta: {totale_in_servizio - totale_dotazione:+4d}"
    )
    print(f"  Medici in discipline fuori atto: {len(medici_non_mappati)}")
    print(f"  Totale dirigenti medici TI: {len(medici_df)}")

    # Rinomina colonne
    def rinomina_colonne(df):
        rename = {
            'DISCIPLINA_ATTO_AZIENDALE': 'Disciplina',
            'DOTAZIONE_ATTO': 'Dotazione Atto',
            'IN_SERVIZIO': 'In Servizio',
            'DELTA': 'Delta',
        }
        for anno in anni_pensionamento:
            rename[f'PENSIONAMENTI_{anno}'] = f'Pensionamenti {anno}'
            rename[f'PROIEZIONE_{anno}'] = f'Proiezione {anno}'
        return df.rename(columns=rename)

    df_atto = pd.DataFrame(
        [r for r in righe_report
         if not str(r.get('DISCIPLINA_ATTO_AZIENDALE', '')).startswith(
             '[FUORI ATTO]'
         )]
    )
    df_fuori = pd.DataFrame(
        [r for r in righe_report
         if str(r.get('DISCIPLINA_ATTO_AZIENDALE', '')).startswith(
             '[FUORI ATTO]'
         )]
    )
    if not df_fuori.empty:
        df_fuori = df_fuori.copy()
        df_fuori['DISCIPLINA_ATTO_AZIENDALE'] = (
            df_fuori['DISCIPLINA_ATTO_AZIENDALE']
            .str.replace(r'^\[FUORI ATTO\] ', '', regex=True)
        )

    df_atto = rinomina_colonne(df_atto)
    df_fuori = rinomina_colonne(df_fuori).drop(
        columns=['Dotazione Atto', 'Delta'], errors='ignore'
    )

    # Genera XLSX
    def _scrivi_foglio_atto(wb, sheet_name, df, titolo,
                             col_gruppo='Disciplina'):
        ws = wb.create_sheet(title=sheet_name[:31])
        cols = list(df.columns)
        scrivi_titolo(ws, titolo, len(cols))
        scrivi_intestazioni(ws, cols)
        color_toggle = 0
        prev_grp = None
        for r_idx, row_data in enumerate(df.itertuples(index=False), 3):
            grp_idx = cols.index(col_gruppo) if col_gruppo in cols else 0
            current_grp = row_data[grp_idx]
            if current_grp != prev_grp:
                color_toggle = 1 - color_toggle
                prev_grp = current_grp
            scrivi_riga_dati(
                ws, r_idx, list(row_data),
                FILL_A if color_toggle else FILL_B,
            )

        auto_larghezza_colonne(ws, cols)

    wb = Workbook()
    wb.remove(wb.active)
    _scrivi_foglio_atto(
        wb, 'Requisiti Atto Aziendale', df_atto,
        'Personale Medico - Requisiti da Atto Aziendale',
    )
    _scrivi_foglio_atto(
        wb, 'Fuori Atto Aziendale', df_fuori,
        'Personale Medico - Discipline Fuori Atto Aziendale',
    )
    wb.save(output_file)

    report_df = pd.concat([df_atto, df_fuori], ignore_index=True)
    print(f"\n  Report salvato in: {output_file}")
    print(f"{'=' * 70}\n")

    return report_df
