"""
Gestione posti letto: lettura CSV, matching DB→personale, generazione
template.  Include anche le funzioni di utilità per identificare
ospedali e reparti.
"""

import re
import unicodedata

import pandas as pd

from src.caricamento_dati import carica_dataframe, normalizza_colonne_personale


# ============================================================
# UTILITÀ DI MATCHING
# ============================================================

def assegna_intensita(desc_sc_ssd_ss, mapping_intensita_pattern):
    """
    Assegna un'intensità di default a un reparto in base al nome,
    usando la tabella di pattern caricata dal XML.
    """
    nome = desc_sc_ssd_ss.upper()
    for pattern, intensita in mapping_intensita_pattern:
        if re.search(pattern, nome):
            return intensita
    return ''


def identifica_ospedale_personale(nome_ospedale_db, mapping_ospedali):
    """
    Dato il nome ospedale dal dump DB, restituisce la sottostringa
    da cercare in DESC_SEDE_FISICA del personale.
    Restituisce None se non è un ospedale ASREM.
    """
    nome_upper = nome_ospedale_db.upper()
    for chiave, valore in mapping_ospedali.items():
        if chiave in nome_upper:
            return valore
    return None


def trova_reparto_personale(nome_reparto_db, sede_keyword,
                            reparti_personale, mapping_reparti):
    """
    Dato il nome reparto dal dump DB e la keyword della sede,
    cerca il reparto corrispondente tra quelli del personale.

    Restituisce la tupla (DESC_SEDE_FISICA, DESC_SC_SSD_SS)
    oppure None se non trovato.
    """
    nome_upper = nome_reparto_db.strip().upper()

    # Controlla nella mappa esplicita (con e senza accenti)
    nome_no_accent = unicodedata.normalize('NFD', nome_upper)
    nome_no_accent = ''.join(
        c for c in nome_no_accent if unicodedata.category(c) != 'Mn'
    )

    if nome_upper in mapping_reparti:
        pattern = mapping_reparti[nome_upper]
    elif nome_no_accent in mapping_reparti:
        pattern = mapping_reparti[nome_no_accent]
    else:
        pattern = re.escape(nome_upper)

    if pattern is None:
        return None  # esplicitamente escluso

    reparti_sede = [
        r for r in reparti_personale
        if isinstance(r[0], str) and sede_keyword.upper() in r[0].upper()
    ]

    matches = []
    for sede_fisica, ssd in reparti_sede:
        if re.search(pattern, ssd.upper()):
            matches.append((sede_fisica, ssd))

    if len(matches) >= 1:
        return matches[0]
    return None


# ============================================================
# LETTURA CSV POSTI LETTO
# ============================================================

def leggi_posti_letto_csv(file_path):
    """
    Legge il CSV dei posti letto e restituisce un dizionario indicizzato
    per (DESC_SEDE_FISICA, DESC_SC_SSD_SS).
    """
    df = pd.read_csv(file_path, delimiter=';', encoding='UTF-8')

    colonne_numeriche = [
        'ordinari', 'dh', 'utic', 'ds',
    ]
    for col in colonne_numeriche:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    df['intensita'] = df['intensita'].fillna('').str.strip()

    posti_letto = {}
    for _, row in df.iterrows():
        key = (row['DESC_SEDE_FISICA'], row['DESC_SC_SSD_SS'])
        # Somma eventuali colonne breast_unit (retrocompatibilità CSV)
        bu_ord = int(row.get('breast_unit_ordinari', 0))
        bu_dh = int(row.get('breast_unit_dh', 0))
        posti_letto[key] = {
            'DESC_SEDE_FISICA': row['DESC_SEDE_FISICA'],
            'ordinari': int(row['ordinari']) + bu_ord,
            'dh': int(row['dh']) + bu_dh,
            'utic': row.get('utic', 0),
            'ds': row.get('ds', 0),
            'intensita': row['intensita'],
        }
    return posti_letto


def verifica_posti_letto_compilati(file_path):
    """
    Verifica se il file dei posti letto è stato compilato
    (almeno una riga con intensità valorizzata).
    """
    df = pd.read_csv(file_path, delimiter=';', encoding='UTF-8')
    intensita_compilate = df['intensita'].fillna('').str.strip()
    return (intensita_compilate != '').any()


# ============================================================
# GENERAZIONE TEMPLATE POSTI LETTO DA DB
# ============================================================

def genera_posti_letto_da_db(personale_file, reparti_db_file, output_csv,
                             mapping_ospedali, mapping_reparti,
                             mapping_intensita_pattern, lista_odc=None):
    """
    Genera il CSV posti_letto.csv incrociando:
    1. Le unità operative dal database del personale
    2. I posti letto dal dump DB reparti
    3. Le strutture degli Ospedali di Comunità (da XML)
    """
    print(f"Lettura file personale: {personale_file}")
    personale_df = carica_dataframe(personale_file)
    personale_df = normalizza_colonne_personale(personale_df)
    personale_ti = personale_df[
        personale_df['DESC_NATURA'].str.upper() == "TEMPO INDETERMINATO"
    ]

    print(f"Lettura dump DB reparti: {reparti_db_file}")
    reparti_db = pd.read_csv(reparti_db_file, delimiter=';',
                             encoding='ISO-8859-1')

    # Combinazioni uniche sede + reparto
    reparti_personale = (
        personale_ti[['DESC_SEDE_FISICA', 'DESC_SC_SSD_SS']]
        .drop_duplicates()
        .dropna(subset=['DESC_SEDE_FISICA', 'DESC_SC_SSD_SS'])
        .sort_values(['DESC_SEDE_FISICA', 'DESC_SC_SSD_SS'])
        .reset_index(drop=True)
    )
    lista_reparti = list(
        zip(reparti_personale['DESC_SEDE_FISICA'],
            reparti_personale['DESC_SC_SSD_SS'])
    )

    # Struttura per accumulare posti letto per (sede, reparto)
    posti_letto_map = {}
    for sede, ssd in lista_reparti:
        key = (sede, ssd)
        if key not in posti_letto_map:
            posti_letto_map[key] = {
                'ordinari': 0, 'dh': 0, 'ds': 0, 'utic': 0,
                'match_db': [],
            }

    # Per ogni riga del dump DB, trova il reparto del personale corrispondente
    match_log = []
    no_match_log = []

    for _, row_db in reparti_db.iterrows():
        nome_ospedale = str(row_db['nome ospedale'])
        nome_reparto = str(row_db['nome_reparto'])
        ordinari = int(row_db['posti_ordinari'])
        dh = int(row_db['posti_dh'])
        ds = int(row_db['posti_ds'])

        sede_keyword = identifica_ospedale_personale(nome_ospedale,
                                                     mapping_ospedali)
        if sede_keyword is None:
            no_match_log.append(
                f"  Ospedale non ASREM: {nome_ospedale} - {nome_reparto}"
            )
            continue

        match_result = trova_reparto_personale(
            nome_reparto, sede_keyword, lista_reparti, mapping_reparti
        )
        if match_result is None:
            no_match_log.append(
                f"  Reparto non trovato: {nome_ospedale} - {nome_reparto} "
                f"(ord={ordinari}, dh={dh}, ds={ds})"
            )
            continue

        sede_match, ssd_match = match_result
        pl_key = (sede_match, ssd_match)

        nome_upper = nome_reparto.strip().upper()
        nome_upper_no_accent = unicodedata.normalize('NFD', nome_upper)
        nome_upper_no_accent = ''.join(
            c for c in nome_upper_no_accent
            if unicodedata.category(c) != 'Mn'
        )

        if 'CORONARICA' in nome_upper or 'CORONARICA' in nome_upper_no_accent:
            posti_letto_map[pl_key]['utic'] += ordinari
            match_log.append(
                f"  {nome_ospedale:50s} | {nome_reparto:40s} -> "
                f"{sede_match} / {ssd_match} (UTIC={ordinari})"
            )
        elif 'SENOLOGIC' in nome_upper:
            posti_letto_map[pl_key]['ordinari'] += ordinari
            posti_letto_map[pl_key]['dh'] += dh + ds
            match_log.append(
                f"  {nome_ospedale:50s} | {nome_reparto:40s} -> "
                f"{sede_match} / {ssd_match} "
                f"(ord={ordinari}, dh={dh + ds})"
            )
        else:
            posti_letto_map[pl_key]['ordinari'] += ordinari
            posti_letto_map[pl_key]['dh'] += dh
            posti_letto_map[pl_key]['ds'] += ds
            match_log.append(
                f"  {nome_ospedale:50s} | {nome_reparto:40s} -> "
                f"{sede_match} / {ssd_match} "
                f"(ord={ordinari}, dh={dh}, ds={ds})"
            )
        posti_letto_map[pl_key]['match_db'].append(nome_reparto)

    # ---------------------------------------------------------------
    # OSPEDALI DI COMUNITÀ
    # ---------------------------------------------------------------
    odc_rows = []
    odc_log = []
    if lista_odc:
        for odc in lista_odc:
            nome_odc = odc['nome']
            print(f"\nElaborazione Ospedale di Comunità: {nome_odc}")
            for struttura in odc['strutture']:
                pattern_cdc = struttura['pattern_cdc']
                if not pattern_cdc:
                    continue

                mask = personale_ti['DESC_TIPO_CDC'].fillna('').str.contains(
                    pattern_cdc, case=False, regex=True
                )
                n_personale = mask.sum()
                pl = struttura['posti_letto']
                nome_struttura = struttura['nome']
                intensita = struttura['intensita']

                odc_rows.append({
                    'DESC_SEDE_FISICA': f"OdC {nome_odc}",
                    'DESC_SC_SSD_SS': nome_struttura,
                    'ordinari': pl['ordinari'],
                    'dh': pl['dh'],
                    'utic': 0,
                    'ds': pl['ds'],
                    'intensita': intensita,
                    'match_db_reparti': (
                        f'OdC (CDC: {pattern_cdc}, personale: {n_personale})'
                    ),
                })
                odc_log.append(
                    f"  {'OdC ' + nome_odc:50s} | {nome_struttura:40s} "
                    f"(ord={pl['ordinari']}, dh={pl['dh']}, ds={pl['ds']}, "
                    f"intensita={intensita}, personale={n_personale})"
                )

    # DataFrame finale
    rows = []
    for _, r in reparti_personale.iterrows():
        ssd = r['DESC_SC_SSD_SS']
        sede = r['DESC_SEDE_FISICA']
        pl_key = (sede, ssd)
        pl = posti_letto_map.get(pl_key, {})
        rows.append({
            'DESC_SEDE_FISICA': sede,
            'DESC_SC_SSD_SS': ssd,
            'ordinari': pl.get('ordinari', 0),
            'dh': pl.get('dh', 0),
            'utic': pl.get('utic', 0),
            'ds': pl.get('ds', 0),
            'intensita': assegna_intensita(ssd, mapping_intensita_pattern),
            'match_db_reparti': '; '.join(pl.get('match_db', [])),
        })
    rows.extend(odc_rows)

    result_df = pd.DataFrame(rows)
    result_df.to_csv(output_csv, index=False, sep=';', encoding='UTF-8')

    # Log
    print(f"\n{'=' * 70}")
    print(f"TEMPLATE POSTI LETTO GENERATO: {output_csv}")
    print(f"{'=' * 70}")
    print(f"Reparti ospedalieri totali: {len(reparti_personale)}")

    reparti_con_match = sum(
        1 for pl in posti_letto_map.values() if pl.get('match_db')
    )
    print(f"Reparti con posti letto assegnati: {reparti_con_match}")
    print(f"Reparti senza posti letto: "
          f"{len(reparti_personale) - reparti_con_match}")

    print(f"\n--- MATCH RIUSCITI (Ospedali) ---")
    for line in match_log:
        print(line)

    if odc_log:
        print(f"\n--- OSPEDALI DI COMUNITÀ ---")
        for line in odc_log:
            print(line)

    if no_match_log:
        print(f"\n--- SENZA MATCH (da verificare) ---")
        for line in no_match_log:
            print(line)

    print(f"\n{'=' * 70}")
    print(f"APRI IL FILE '{output_csv}' e compila/verifica:")
    print(f"  1. CONTROLLA che i posti letto siano stati assegnati correttamente")
    print(f"  2. VERIFICA la colonna 'intensita' (pre-compilata automaticamente)")
    print(f"     Valori: Intensiva, Alta, Medio/Alta, Media, "
          f"Medio/Bassa, Bassa, DH_DS")
    print(f"  3. La colonna 'match_db_reparti' mostra da dove provengono i dati")
    print(f"Poi riesegui lo script per calcolare il fabbisogno.")
    print(f"{'=' * 70}\n")

    return result_df
