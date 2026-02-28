"""
Elaborazione Personale ASREM – Entry Point.

Questo script orchestra l'intera elaborazione importando i moduli
dalla cartella src/.  Tutta la logica è nei moduli; qui c'è solo
il flusso di esecuzione.
"""

import os
import sys

os.makedirs('elaborati', exist_ok=True)

from src.config import (
    ANNO_ANALISI, DIR_ELABORATI,
    FILE_PERSONALE, FILE_PENSIONAMENTI, FILE_REPARTI_DB,
    FILE_POSTI_LETTO_CSV,
    FILE_INDICATORI, FILE_MAPPING_OSPEDALI, FILE_MAPPING_REPARTI,
    FILE_MAPPING_ODC, FILE_INDICATORI_ODC,
    FILE_MEDICI_ATTO_AZIENDALE,
    FILE_OUTPUT, FILE_DEBUG,
    FILE_OUTPUT_ATTO_AZIENDALE, FILE_OUTPUT_ODC,
    FILE_INDICATORI_AGENAS_MATERNO_INFANTILE,
    PARTI_PER_PRESIDIO,
    FILE_INDICATORI_AGENAS_RADIOLOGIA,
    FILE_INDICATORI_AGENAS_TRASFUSIONALE,
    FILE_INDICATORI_TRASFUSIONALE_SPECIALE,
    FILE_INDICATORI_AGENAS_ANATOMIA_PATOLOGICA,
    FILE_INDICATORI_AGENAS_LABORATORIO,
    FILE_INDICATORI_AGENAS_TECNICI_LABORATORIO,
    FILE_INDICATORI_AGENAS_MEDICINA_LEGALE,
    FILE_INDICATORI_AGENAS_EMERGENZA_URGENZA,
    FILE_INDICATORI_AGENAS_TERAPIA_INTENSIVA,
    FILE_INDICATORI_AGENAS_SALE_OPERATORIE,
    SALE_OPERATORIE_PER_PRESIDIO,
    LIVELLO_PRESIDIO,
    POPOLAZIONE_AREA, DETENUTI_PER_ISTITUTO,
    FILE_INDICATORI_AGENAS_SALUTE_MENTALE,
    FILE_INDICATORI_AGENAS_DIPENDENZE,
    FILE_INDICATORI_AGENAS_NPIA,
    FILE_INDICATORI_AGENAS_CARCERE,
)
from src.caricamento_dati import (
    carica_dataframe, normalizza_colonne_personale,
)
from src.caricamento_xml import (
    carica_mapping_ospedali, carica_sigle_ospedali,
    carica_mapping_reparti,
    carica_intensita_per_pattern, carica_mapping_odc,
    carica_indicatori_agenas_materno_infantile,
    carica_indicatori_agenas_radiologia,
    carica_indicatori_agenas_trasfusionale,
    carica_fabbisogno_trasfusionale_speciale,
    carica_indicatori_agenas_anatomia_patologica,
    carica_indicatori_agenas_laboratorio,
    carica_indicatori_agenas_tecnici_laboratorio,
    carica_indicatori_agenas_medicina_legale,
    carica_indicatori_agenas_emergenza_urgenza,
    carica_indicatori_agenas_salute_mentale,
    carica_indicatori_agenas_dipendenze,
    carica_indicatori_agenas_npia,
    carica_indicatori_agenas_carcere,
    carica_indicatori_agenas_terapia_intensiva,
    carica_indicatori_agenas_sale_operatorie,
)
from src.posti_letto import (
    assegna_intensita, genera_posti_letto_da_db,
    verifica_posti_letto_compilati,
)
from src.report_fabbisogno import process_data
from src.report_atto_aziendale import genera_report_atto_aziendale
from src.report_odc import genera_report_odc


# ============================================================
# MAIN
# ============================================================
if __name__ == '__main__':

    print(f"\n{'=' * 70}")
    print(f"ELABORAZIONE PERSONALE ASREM - Anno di analisi: {ANNO_ANALISI}")
    print(f"{'=' * 70}\n")

    # --- Caricamento mapping da XML ---
    print("Caricamento mapping da XML...")
    mapping_ospedali = carica_mapping_ospedali(FILE_MAPPING_OSPEDALI)
    sigle_ospedali = carica_sigle_ospedali(FILE_MAPPING_OSPEDALI)
    mapping_reparti = carica_mapping_reparti(FILE_MAPPING_REPARTI)
    mapping_intensita_pattern = carica_intensita_per_pattern(FILE_MAPPING_REPARTI)
    print(f"  Ospedali: {len(mapping_ospedali)} regole")
    print(f"  Reparti: {len(mapping_reparti)} regole")
    print(f"  Intensità: {len(mapping_intensita_pattern)} pattern")

    # Caricamento Ospedali di Comunità
    lista_odc = []
    for odc_file in FILE_MAPPING_ODC:
        if os.path.exists(odc_file):
            odc = carica_mapping_odc(odc_file)
            lista_odc.append(odc)
            print(f"  OdC {odc['nome']}: {len(odc['strutture'])} strutture")
    print()

    # STEP 1: Se il file posti_letto.csv non esiste, generalo
    if not os.path.exists(FILE_POSTI_LETTO_CSV):
        if os.path.exists(FILE_REPARTI_DB):
            print("File posti letto CSV non trovato.")
            print("Generazione automatica dal dump DB reparti "
                  "+ database personale + OdC...\n")
            genera_posti_letto_da_db(
                FILE_PERSONALE, FILE_REPARTI_DB, FILE_POSTI_LETTO_CSV,
                mapping_ospedali, mapping_reparti,
                mapping_intensita_pattern, lista_odc,
                sigle_ospedali=sigle_ospedali,
            )
        else:
            print("File posti letto CSV e dump DB reparti non trovati.")
            print("Generazione template vuoto dal database personale...\n")
            personale_df = carica_dataframe(FILE_PERSONALE)
            personale_df = normalizza_colonne_personale(personale_df)
            personale_df = personale_df[
                personale_df['DESC_NATURA'].str.upper() == "TEMPO INDETERMINATO"
            ]
            reparti = personale_df[
                ['DESC_SEDE_FISICA', 'DESC_SC_SSD_SS']
            ].drop_duplicates()
            reparti = reparti.sort_values(
                ['DESC_SEDE_FISICA', 'DESC_SC_SSD_SS']
            ).reset_index(drop=True)
            for col in ['ordinari', 'dh', 'utic', 'ds']:
                reparti[col] = 0
            reparti['intensita'] = reparti['DESC_SC_SSD_SS'].apply(
                lambda ssd: assegna_intensita(ssd, mapping_intensita_pattern)
            )
            reparti['match_db_reparti'] = ''
            reparti.to_csv(
                FILE_POSTI_LETTO_CSV, index=False, sep=';', encoding='UTF-8'
            )
            print(f"Template vuoto generato: {FILE_POSTI_LETTO_CSV}")

        print("\nCompila il file dei posti letto e riesegui lo script.")
        sys.exit(0)

    # STEP 2: Verifica che sia stato compilato
    if not verifica_posti_letto_compilati(FILE_POSTI_LETTO_CSV):
        print(f"ATTENZIONE: Il file '{FILE_POSTI_LETTO_CSV}' non è "
              "ancora stato compilato!")
        print("Apri il file e inserisci/verifica i valori dei posti letto "
              "e le intensità.")
        print("Poi riesegui lo script.\n")

        risposta = input(
            "Vuoi rigenerare il template dal dump DB? (s/n): "
        ).strip().lower()
        if risposta == 's' and os.path.exists(FILE_REPARTI_DB):
            genera_posti_letto_da_db(
                FILE_PERSONALE, FILE_REPARTI_DB, FILE_POSTI_LETTO_CSV,
                mapping_ospedali, mapping_reparti,
                mapping_intensita_pattern, lista_odc,
                sigle_ospedali=sigle_ospedali,
            )
        sys.exit(0)

    # STEP 3: Elaborazione fabbisogno
    indicatori_agenas = None
    if os.path.exists(FILE_INDICATORI_AGENAS_MATERNO_INFANTILE):
        indicatori_agenas = carica_indicatori_agenas_materno_infantile(
            FILE_INDICATORI_AGENAS_MATERNO_INFANTILE
        )
        print(f"Indicatori AGENAS materno-infantile caricati: "
              f"{len(indicatori_agenas['fasce'])} fasce")
        print(f"Parti per presidio: {PARTI_PER_PRESIDIO}")

    indicatori_radiologia = None
    if os.path.exists(FILE_INDICATORI_AGENAS_RADIOLOGIA):
        indicatori_radiologia = carica_indicatori_agenas_radiologia(
            FILE_INDICATORI_AGENAS_RADIOLOGIA
        )
        print(f"Indicatori AGENAS radiologia caricati: "
              f"{len(indicatori_radiologia['livelli'])} livelli")
        print(f"Livello presidi: {LIVELLO_PRESIDIO}")

    indicatori_trasfusionale = None
    if os.path.exists(FILE_INDICATORI_AGENAS_TRASFUSIONALE):
        indicatori_trasfusionale = carica_indicatori_agenas_trasfusionale(
            FILE_INDICATORI_AGENAS_TRASFUSIONALE
        )
        print(f"Indicatori AGENAS trasfusionale caricati: "
              f"{len(indicatori_trasfusionale['livelli'])} livelli")

    fabb_trasf_speciale = None
    if os.path.exists(FILE_INDICATORI_TRASFUSIONALE_SPECIALE):
        fabb_trasf_speciale = carica_fabbisogno_trasfusionale_speciale(
            FILE_INDICATORI_TRASFUSIONALE_SPECIALE
        )
        print(f"Fabbisogno UOC Trasfusionale (speciale) caricato: "
              f"{len(fabb_trasf_speciale['sedi'])} sedi, "
              f"{len(fabb_trasf_speciale['ordine_profili'])} profili")

    indicatori_anatomia_pat = None
    if os.path.exists(FILE_INDICATORI_AGENAS_ANATOMIA_PATOLOGICA):
        indicatori_anatomia_pat = carica_indicatori_agenas_anatomia_patologica(
            FILE_INDICATORI_AGENAS_ANATOMIA_PATOLOGICA
        )
        print(f"Indicatori AGENAS anatomia patologica caricati: "
              f"{len(indicatori_anatomia_pat['livelli'])} livelli")

    indicatori_laboratorio = None
    if os.path.exists(FILE_INDICATORI_AGENAS_LABORATORIO):
        indicatori_laboratorio = carica_indicatori_agenas_laboratorio(
            FILE_INDICATORI_AGENAS_LABORATORIO
        )
        print(f"Indicatori AGENAS laboratorio caricati: "
              f"{len(indicatori_laboratorio['livelli'])} livelli")

    indicatori_tecnici_lab = None
    if os.path.exists(FILE_INDICATORI_AGENAS_TECNICI_LABORATORIO):
        indicatori_tecnici_lab = carica_indicatori_agenas_tecnici_laboratorio(
            FILE_INDICATORI_AGENAS_TECNICI_LABORATORIO
        )
        print(f"Indicatori AGENAS tecnici laboratorio caricati: "
              f"{len(indicatori_tecnici_lab['livelli'])} livelli")

    indicatori_med_legale = None
    if os.path.exists(FILE_INDICATORI_AGENAS_MEDICINA_LEGALE):
        indicatori_med_legale = carica_indicatori_agenas_medicina_legale(
            FILE_INDICATORI_AGENAS_MEDICINA_LEGALE
        )
        print(f"Indicatori AGENAS medicina legale caricati: "
              f"{len(indicatori_med_legale['livelli'])} livelli")

    indicatori_emergenza = None
    if os.path.exists(FILE_INDICATORI_AGENAS_EMERGENZA_URGENZA):
        indicatori_emergenza = carica_indicatori_agenas_emergenza_urgenza(
            FILE_INDICATORI_AGENAS_EMERGENZA_URGENZA
        )
        print(f"Indicatori AGENAS emergenza-urgenza caricati: "
              f"{len(indicatori_emergenza['livelli'])} livelli")

    indicatori_terapia_intensiva = None
    if os.path.exists(FILE_INDICATORI_AGENAS_TERAPIA_INTENSIVA):
        indicatori_terapia_intensiva = carica_indicatori_agenas_terapia_intensiva(
            FILE_INDICATORI_AGENAS_TERAPIA_INTENSIVA
        )
        print(f"Indicatori AGENAS terapia intensiva caricati: "
              f"{len(indicatori_terapia_intensiva['standard'])} profili")

    indicatori_sale_operatorie = None
    if os.path.exists(FILE_INDICATORI_AGENAS_SALE_OPERATORIE):
        indicatori_sale_operatorie = carica_indicatori_agenas_sale_operatorie(
            FILE_INDICATORI_AGENAS_SALE_OPERATORIE
        )
        print(f"Indicatori AGENAS sale operatorie caricati: "
              f"{len(indicatori_sale_operatorie['standard'])} profili")

    # --- Indicatori territoriali (basati su popolazione) ---
    from src.calcolo_fabbisogno import calcola_fabbisogno_agenas_territoriale

    indicatori_salute_mentale = None
    fabb_salute_mentale = None
    if os.path.exists(FILE_INDICATORI_AGENAS_SALUTE_MENTALE):
        indicatori_salute_mentale = carica_indicatori_agenas_salute_mentale(
            FILE_INDICATORI_AGENAS_SALUTE_MENTALE
        )
        fabb_salute_mentale = calcola_fabbisogno_agenas_territoriale(
            indicatori_salute_mentale, POPOLAZIONE_AREA
        )

    indicatori_dipendenze = None
    fabb_dipendenze = None
    if os.path.exists(FILE_INDICATORI_AGENAS_DIPENDENZE):
        indicatori_dipendenze = carica_indicatori_agenas_dipendenze(
            FILE_INDICATORI_AGENAS_DIPENDENZE
        )
        fabb_dipendenze = calcola_fabbisogno_agenas_territoriale(
            indicatori_dipendenze, POPOLAZIONE_AREA
        )

    indicatori_npia = None
    fabb_npia = None
    if os.path.exists(FILE_INDICATORI_AGENAS_NPIA):
        indicatori_npia = carica_indicatori_agenas_npia(
            FILE_INDICATORI_AGENAS_NPIA
        )
        fabb_npia = calcola_fabbisogno_agenas_territoriale(
            indicatori_npia, POPOLAZIONE_AREA
        )

    indicatori_carcere = None
    fabb_carcere = None
    if os.path.exists(FILE_INDICATORI_AGENAS_CARCERE):
        indicatori_carcere = carica_indicatori_agenas_carcere(
            FILE_INDICATORI_AGENAS_CARCERE
        )
        fabb_carcere = calcola_fabbisogno_agenas_territoriale(
            indicatori_carcere, POPOLAZIONE_AREA, DETENUTI_PER_ISTITUTO
        )

    process_data(
        personale_file=FILE_PERSONALE,
        pensionamenti_file=FILE_PENSIONAMENTI,
        posti_letto_csv=FILE_POSTI_LETTO_CSV,
        indicators_file=FILE_INDICATORI,
        debug_file=FILE_DEBUG,
        anno_analisi=ANNO_ANALISI,
        indicatori_odc_file=FILE_INDICATORI_ODC,
        indicatori_agenas=indicatori_agenas,
        parti_per_presidio=PARTI_PER_PRESIDIO,
        indicatori_radiologia=indicatori_radiologia,
        livello_presidio=LIVELLO_PRESIDIO,
        indicatori_trasfusionale=indicatori_trasfusionale,
        fabb_trasf_speciale=fabb_trasf_speciale,
        indicatori_anatomia_pat=indicatori_anatomia_pat,
        indicatori_laboratorio=indicatori_laboratorio,
        indicatori_tecnici_lab=indicatori_tecnici_lab,
        indicatori_med_legale=indicatori_med_legale,
        indicatori_emergenza=indicatori_emergenza,
        indicatori_salute_mentale=indicatori_salute_mentale,
        fabb_salute_mentale=fabb_salute_mentale,
        indicatori_dipendenze=indicatori_dipendenze,
        fabb_dipendenze=fabb_dipendenze,
        indicatori_npia=indicatori_npia,
        fabb_npia=fabb_npia,
        indicatori_carcere=indicatori_carcere,
        fabb_carcere=fabb_carcere,
        indicatori_terapia_intensiva=indicatori_terapia_intensiva,
        indicatori_sale_operatorie=indicatori_sale_operatorie,
        lista_odc=lista_odc,
    )

    # STEP 4: Report medici – atto aziendale
    if os.path.exists(FILE_MEDICI_ATTO_AZIENDALE):
        genera_report_atto_aziendale(
            personale_file=FILE_PERSONALE,
            pensionamenti_file=FILE_PENSIONAMENTI,
            mapper_atto_aziendale=FILE_MEDICI_ATTO_AZIENDALE,
            output_file=FILE_OUTPUT_ATTO_AZIENDALE,
            anno_analisi=ANNO_ANALISI,
        )

    # STEP 5: Report OdC – DM 77
    if lista_odc:
        genera_report_odc(
            personale_file=FILE_PERSONALE,
            pensionamenti_file=FILE_PENSIONAMENTI,
            lista_odc=lista_odc,
            indicatori_odc_file=FILE_INDICATORI_ODC,
            output_file=FILE_OUTPUT_ODC,
            anno_analisi=ANNO_ANALISI,
        )
