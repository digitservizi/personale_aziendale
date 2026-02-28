"""
Report Ospedali di Comunità – confronto personale vs fabbisogno DM 77/2022.
Include foglio RIEPILOGO e fogli dettaglio per sede ODC.
"""

import os

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from src.stili_excel import (
    THIN_BORDER, FILL_A, FILL_B,
    scrivi_titolo, scrivi_intestazioni, scrivi_riga_dati,
    auto_larghezza_colonne,
)
from src.caricamento_dati import (
    carica_dataframe, normalizza_colonne_personale,
    normalizza_colonne_pensionamenti,
)
from src.caricamento_xml import carica_fabbisogno_odc_dm77


# ============================================================
# HELPER INTERNO
# ============================================================

def _scrivi_foglio_odc(ws, titolo, colonne, righe_df):
    """Scrive un foglio ODC con titolo, intestazioni e righe colorate."""
    scrivi_titolo(ws, titolo, len(colonne))
    scrivi_intestazioni(ws, colonne)

    prev_struttura = None
    fill_idx = 0
    for r_idx, (_, riga) in enumerate(righe_df.iterrows(), start=3):
        struttura = riga.get('Struttura', '')
        if struttura != prev_struttura:
            if prev_struttura is not None:
                fill_idx += 1
            prev_struttura = struttura
        fill = FILL_A if fill_idx % 2 == 0 else FILL_B
        valori = [riga[c] for c in colonne]
        scrivi_riga_dati(ws, r_idx, valori, fill)

    auto_larghezza_colonne(ws, colonne)


# ============================================================
# REPORT ODC
# ============================================================

def genera_report_odc(personale_file, pensionamenti_file, lista_odc,
                      indicatori_odc_file, output_file, anno_analisi):
    """
    Genera il report XLSX degli Ospedali di Comunità con confronto
    DM 77/2022.
    """
    print(f"\n{'=' * 70}")
    print("REPORT OSPEDALI DI COMUNITÀ - CONFRONTO DM 77/2022")
    print(f"{'=' * 70}\n")

    # Carica personale
    personale_df = carica_dataframe(personale_file)
    personale_df = normalizza_colonne_personale(personale_df)
    personale_df = personale_df[
        personale_df['DESC_NATURA'].str.upper() == "TEMPO INDETERMINATO"
    ]

    # Pensionamenti
    pensionamenti_df = carica_dataframe(pensionamenti_file)
    pensionamenti_df = normalizza_colonne_pensionamenti(pensionamenti_df)

    if 'DT_CESSAZIONE' in personale_df.columns:
        personale_df = personale_df.rename(
            columns={'DT_CESSAZIONE': 'DT_CESSAZIONE_PERS'}
        )
    personale_df = pd.merge(
        personale_df,
        pensionamenti_df[['MATR.', 'DT_CESSAZIONE']],
        on='MATR.',
        how='left',
    )

    anni_pensionamento = [anno_analisi + 1, anno_analisi + 2, anno_analisi + 3]

    # Fabbisogno DM 77
    fabbisogno_dm77 = {}
    mappa_profili_dm77 = {}
    if indicatori_odc_file and os.path.exists(indicatori_odc_file):
        fabbisogno_dm77, mappa_profili_dm77 = carica_fabbisogno_odc_dm77(
            indicatori_odc_file
        )

    righe_report = []

    for odc in lista_odc:
        nome_odc = odc['nome']
        print(f"  === {nome_odc} ===")

        for struttura in odc['strutture']:
            nome_struttura = struttura['nome']
            pattern_cdc = struttura['pattern_cdc']
            posti_letto_ord = struttura['posti_letto']['ordinari']
            note = struttura.get('note', '')

            is_udi = (
                'U.D.I' in nome_struttura.upper()
                or 'OSP.*COMUNITA' in pattern_cdc.upper()
            )

            mask = personale_df['DESC_TIPO_CDC'].str.contains(
                pattern_cdc, case=False, na=False, regex=True
            )
            personale_struttura = personale_df[mask]

            if len(personale_struttura) == 0 and posti_letto_ord == 0:
                continue

            # Rimappa i profili alle categorie DM 77 per tutte le strutture
            # (unifica le etichette: FISIOTERAPISTA → PROF SAN RIAB, ecc.)
            if mappa_profili_dm77:
                personale_struttura = personale_struttura.copy()
                personale_struttura['PROFILO_RAGGRUPPATO'] = (
                    personale_struttura['PROFILO_RAGGRUPPATO']
                    .apply(lambda p: mappa_profili_dm77.get(
                        p.strip().upper(), p
                    ))
                )

            profili = personale_struttura.groupby(
                'PROFILO_RAGGRUPPATO'
            ).agg(
                IN_SERVIZIO=('PROFILO_RAGGRUPPATO', 'size')
            ).reset_index()

            if len(profili) == 0:
                riga = {
                    'ODC': nome_odc,
                    'STRUTTURA': nome_struttura,
                    'POSTI_LETTO': posti_letto_ord,
                    'TIPO': 'DM 77' if is_udi else 'Struttura',
                    'PROFILO': '-',
                    'IN_SERVIZIO': 0,
                    'FABBISOGNO_DM77': '-',
                    'DELTA': '-',
                    'NOTE': note,
                }
                for anno in anni_pensionamento:
                    riga[f'PENSIONAMENTI_{anno}'] = 0
                    riga[f'PROIEZIONE_{anno}'] = 0
                righe_report.append(riga)
                print(
                    f"    {nome_struttura:45s}  PL: {posti_letto_ord:3d}  "
                    f"Personale: 0"
                )
                continue

            totale_struttura = 0
            totale_fabb = 0
            for _, pr in profili.iterrows():
                profilo = pr['PROFILO_RAGGRUPPATO']
                in_servizio = pr['IN_SERVIZIO']
                totale_struttura += in_servizio

                if is_udi and fabbisogno_dm77:
                    fabb = fabbisogno_dm77.get(profilo.strip().upper(), 0)
                    delta = in_servizio - fabb
                    fabb_str = fabb
                    totale_fabb += fabb
                else:
                    fabb_str = '-'
                    delta = '-'

                pers_profilo = personale_struttura[
                    personale_struttura['PROFILO_RAGGRUPPATO'] == profilo
                ]
                pens = {}
                for anno in anni_pensionamento:
                    pens[anno] = pers_profilo['DT_CESSAZIONE'].apply(
                        lambda x, a=anno: (
                            1 if pd.to_datetime(x, errors='coerce').year == a
                            else 0
                        )
                    ).sum()

                riga = {
                    'ODC': nome_odc,
                    'STRUTTURA': nome_struttura,
                    'POSTI_LETTO': posti_letto_ord,
                    'TIPO': 'DM 77' if is_udi else 'Struttura',
                    'PROFILO': profilo,
                    'IN_SERVIZIO': in_servizio,
                    'FABBISOGNO_DM77': fabb_str,
                    'DELTA': delta,
                    'NOTE': note,
                }
                for anno in anni_pensionamento:
                    riga[f'PENSIONAMENTI_{anno}'] = pens[anno]
                pens_cum = 0
                for anno in anni_pensionamento:
                    pens_cum += pens[anno]
                    riga[f'PROIEZIONE_{anno}'] = in_servizio - pens_cum
                righe_report.append(riga)

            # Profili DM 77 mancanti
            if is_udi and fabbisogno_dm77:
                profili_presenti = set(
                    profili['PROFILO_RAGGRUPPATO'].str.upper()
                )
                for profilo_dm, fabb_dm in fabbisogno_dm77.items():
                    if profilo_dm not in profili_presenti:
                        riga = {
                            'ODC': nome_odc,
                            'STRUTTURA': nome_struttura,
                            'POSTI_LETTO': posti_letto_ord,
                            'TIPO': 'DM 77',
                            'PROFILO': f'[MANCANTE] {profilo_dm}',
                            'IN_SERVIZIO': 0,
                            'FABBISOGNO_DM77': fabb_dm,
                            'DELTA': -fabb_dm,
                            'NOTE': 'Profilo previsto dal DM 77 ma assente',
                        }
                        for anno in anni_pensionamento:
                            riga[f'PENSIONAMENTI_{anno}'] = 0
                            riga[f'PROIEZIONE_{anno}'] = 0
                        righe_report.append(riga)
                        totale_fabb += fabb_dm

            stato = ''
            if is_udi and fabbisogno_dm77:
                diff = totale_struttura - totale_fabb
                stato = f'  DM77: {totale_fabb}  Delta: {diff:+d}'
            print(
                f"    {nome_struttura:45s}  PL: {posti_letto_ord:3d}  "
                f"Personale: {totale_struttura:3d}{stato}"
            )

        print()

    # ------------------------------------------------------------------
    # Genera XLSX
    # ------------------------------------------------------------------
    report_df = pd.DataFrame(righe_report)

    rename_cols = {
        'ODC': 'ODC',
        'STRUTTURA': 'Struttura',
        'POSTI_LETTO': 'Posti Letto',
        'TIPO': 'Tipo',
        'PROFILO': 'Profilo Professionale',
        'IN_SERVIZIO': 'In Servizio',
        'FABBISOGNO_DM77': 'Fabbisogno DM 77',
        'DELTA': 'Delta',
        'NOTE': 'Note',
    }
    for anno in anni_pensionamento:
        rename_cols[f'PENSIONAMENTI_{anno}'] = f'Pensionamenti {anno}'
        rename_cols[f'PROIEZIONE_{anno}'] = f'Proiezione {anno}'
    report_df = report_df.rename(columns=rename_cols)

    wb = Workbook()
    wb.remove(wb.active)
    colonne_dati = [c for c in report_df.columns if c != 'ODC']

    # ====== FOGLIO RIEPILOGO ======
    ws_riep = wb.create_sheet(title='RIEPILOGO')
    col_pens_odc = [f'Pensionamenti {a}' for a in anni_pensionamento]
    col_riep = (
        ['Sede', 'Profilo Professionale', 'In Servizio',
         'Fabbisogno DM 77', 'Delta']
        + col_pens_odc + ['Proiezione']
    )

    scrivi_titolo(ws_riep, 'Riepilogo Ospedali di Comunità', len(col_riep))
    scrivi_intestazioni(ws_riep, col_riep)

    riep_row = 3
    fill_idx_riep = 0
    prev_odc_riep = None

    for nome_odc, df_odc_riep in report_df.groupby('ODC', sort=False):
        if nome_odc != prev_odc_riep:
            if prev_odc_riep is not None:
                fill_idx_riep += 1
            prev_odc_riep = nome_odc
        fill_riep = FILL_A if fill_idx_riep % 2 == 0 else FILL_B

        for profilo, df_prof in df_odc_riep.groupby(
            'Profilo Professionale', sort=True
        ):
            in_serv = df_prof['In Servizio'].sum()
            fabb_vals = pd.to_numeric(
                df_prof['Fabbisogno DM 77'], errors='coerce'
            )
            fabb = (
                int(fabb_vals.sum()) if fabb_vals.notna().any() else '-'
            )
            delta_vals = pd.to_numeric(df_prof['Delta'], errors='coerce')
            delta = (
                int(delta_vals.sum()) if delta_vals.notna().any() else '-'
            )
            pens_vals = [int(df_prof[cp].sum()) for cp in col_pens_odc]
            proiezione = int(in_serv) - sum(pens_vals)

            valori = (
                [nome_odc, profilo, int(in_serv), fabb, delta]
                + pens_vals + [proiezione]
            )
            scrivi_riga_dati(ws_riep, riep_row, valori, fill_riep)
            riep_row += 1

    auto_larghezza_colonne(ws_riep, col_riep)

    # ====== FOGLI DETTAGLIO ======
    for nome_odc, df_sede in report_df.groupby('ODC', sort=False):
        df_sede = df_sede.drop(columns=['ODC'])
        sheet_name = str(nome_odc)[:31]
        ws = wb.create_sheet(title=sheet_name)
        _scrivi_foglio_odc(
            ws, f'Ospedali di Comunità - {nome_odc}',
            colonne_dati, df_sede,
        )

    wb.save(output_file)

    print(f"  Report salvato in: {output_file}")
    print(f"{'=' * 70}\n")

    return report_df
