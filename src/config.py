"""
Configurazione centralizzata – costanti di percorso e parametri.
Modifica qui i valori principali prima di eseguire l'elaborazione.
"""

# ============================================================
# PARAMETRI PRINCIPALI
# ============================================================
ANNO_ANALISI = 2026  # Anno di riferimento per l'analisi

# ============================================================
# PERCORSI FILE SORGENTE
# ============================================================
FILE_PERSONALE = 'sorgenti/DB_PERSONALE_2025.xls'
FILE_PENSIONAMENTI = 'sorgenti/DB_PENSIONAMENTI_2025.xls'
FILE_REPARTI_DB = 'sorgenti/reparto.csv'

# File posti letto CSV (generato automaticamente dal dump DB + personale)
FILE_POSTI_LETTO_CSV = 'sorgenti/posti_letto.csv'

# ============================================================
# FILE XML DI CONFIGURAZIONE
# ============================================================

# ── Atto aziendale (dotazione organica) ──────────────────────
FILE_MEDICI_ATTO_AZIENDALE = 'configurazione/atto_aziendale/medici_atto_aziendale.xml'
FILE_PROFILI_ATTO_AZIENDALE = 'configurazione/atto_aziendale/profili_atto_aziendale.xml'

# ── Mapping strutture ────────────────────────────────────────
FILE_MAPPING_OSPEDALI = 'configurazione/mapping/mapping_ospedali.xml'
FILE_MAPPING_REPARTI = 'configurazione/mapping/mapping_reparti.xml'
FILE_MAPPING_ODC = [
    'configurazione/mapping/mapping_venafro.xml',
    'configurazione/mapping/mapping_larino.xml',
]

# ── Indicatori AGENAS ────────────────────────────────────────
FILE_INDICATORI_AGENAS_MATERNO_INFANTILE = 'configurazione/indicatori_agenas/indicatori_agenas_materno_infantile.xml'
FILE_INDICATORI_AGENAS_RADIOLOGIA = 'configurazione/indicatori_agenas/indicatori_agenas_radiologia.xml'
FILE_INDICATORI_AGENAS_TRASFUSIONALE = 'configurazione/indicatori_agenas/indicatori_agenas_trasfusionale.xml'
FILE_INDICATORI_AGENAS_ANATOMIA_PATOLOGICA = 'configurazione/indicatori_agenas/indicatori_agenas_anatomia_patologica.xml'
FILE_INDICATORI_AGENAS_LABORATORIO = 'configurazione/indicatori_agenas/indicatori_agenas_laboratorio.xml'
FILE_INDICATORI_AGENAS_TECNICI_LABORATORIO = 'configurazione/indicatori_agenas/indicatori_agenas_tecnici_laboratorio.xml'
FILE_INDICATORI_AGENAS_MEDICINA_LEGALE = 'configurazione/indicatori_agenas/indicatori_agenas_medicina_legale.xml'
FILE_INDICATORI_AGENAS_EMERGENZA_URGENZA = 'configurazione/indicatori_agenas/indicatori_agenas_emergenza_urgenza.xml'
FILE_INDICATORI_AGENAS_TERAPIA_INTENSIVA = 'configurazione/indicatori_agenas/indicatori_agenas_terapia_intensiva.xml'
FILE_INDICATORI_AGENAS_SALE_OPERATORIE = 'configurazione/indicatori_agenas/indicatori_agenas_sale_operatorie.xml'

# ── Altri indicatori ─────────────────────────────────────────
FILE_INDICATORI = 'configurazione/indicatori/indicatori_medici.xml'
FILE_INDICATORI_ODC = 'configurazione/indicatori/indicatori_ospedali_comunita_dm_77.xml'
FILE_INDICATORI_TRASFUSIONALE_SPECIALE = 'configurazione/indicatori/indicatori_trasfusionale_speciale.xml'

# ============================================================
# DATI DI ATTIVITÀ PER PRESIDIO (indicatori AGENAS)
# ============================================================
# Numero di parti annui per presidio ospedaliero (ultimo anno disponibile)
PARTI_PER_PRESIDIO = {
    'CAMPOBASSO - P.O. CARDARELLI':  617,
    'ISERNIA - P.O. VENEZIALE':      260,
    'TERMOLI - P.O. SAN TIMOTEO':    368,
}

# Numero di sale operatorie attive per presidio ospedaliero (§ 8.1.2 AGENAS)
SALE_OPERATORIE_PER_PRESIDIO = {
    'CAMPOBASSO - P.O. CARDARELLI':  9,
    'ISERNIA - P.O. VENEZIALE':      2,
    'TERMOLI - P.O. SAN TIMOTEO':    5,
}

# Maggiorazione organica AGENAS per copertura turnazione, guardie
# interdivisionali, ferie, malattie e altre indisponibilità (15%)
MAGGIORAZIONE_TURNAZIONE = 0.15

# Livello dei presidi ospedalieri (per area radiologica AGENAS - Tab. 13)
# Valori ammessi: OSPEDALE_DI_BASE, PRESIDIO_I_LIVELLO, PRESIDIO_II_LIVELLO
LIVELLO_PRESIDIO = {
    'CAMPOBASSO - P.O. CARDARELLI':  'PRESIDIO_I_LIVELLO',
    'ISERNIA - P.O. VENEZIALE':      'OSPEDALE_DI_BASE',
    'TERMOLI - P.O. SAN TIMOTEO':    'OSPEDALE_DI_BASE',
}

# ============================================================
# POPOLAZIONE PER AREA DISTRETTUALE
# Fonte: dati ISTAT popolazione residente.
# ============================================================
POPOLAZIONE_AREA = {
    'CAMPOBASSO': {
        'totale':    116.489,
        'gte_18':    98_937,   # >= 18 anni
        'range_15_64': 35_817,  # 15-64 anni
        'range_1_17':  17_552,  # 1-17 anni
    },
    'ISERNIA': {
        'totale':     79_912,
        'gte_18':     67_947,
        'range_15_64': 49_594,
        'range_1_17':  11_965,
    },
    'TERMOLI': {
        'totale':     94_235,
        'gte_18':     79_487,
        'range_15_64': 59_518,
        'range_1_17':   14_748,
    },
}

# Numero di detenuti per struttura penitenziaria
DETENUTI_PER_ISTITUTO = {
    'CAMPOBASSO': 184,   # Casa Circondariale 
    'ISERNIA':     80,   # Casa Circondariale
    'TERMOLI':    164,   # Casa Circondariale Larino
}

# ============================================================
# FILE INDICATORI AGENAS – SERVIZI TERRITORIALI
# ============================================================
FILE_INDICATORI_AGENAS_SALUTE_MENTALE = 'configurazione/indicatori_agenas/indicatori_agenas_salute_mentale.xml'
FILE_INDICATORI_AGENAS_DIPENDENZE = 'configurazione/indicatori_agenas/indicatori_agenas_dipendenze.xml'
FILE_INDICATORI_AGENAS_NPIA = 'configurazione/indicatori_agenas/indicatori_agenas_npia.xml'
FILE_INDICATORI_AGENAS_CARCERE = 'configurazione/indicatori_agenas/indicatori_agenas_carcere.xml'

# ============================================================
# FILE DI OUTPUT
# ============================================================
DIR_ELABORATI = 'elaborati'

FILE_OUTPUT = f'{DIR_ELABORATI}/analisi_personale_{ANNO_ANALISI}.xlsx'
FILE_OUTPUT_ODC = f'{DIR_ELABORATI}/odc_dm77_{ANNO_ANALISI}.xlsx'
FILE_DEBUG = f'{DIR_ELABORATI}/controprova_calcoli_{ANNO_ANALISI}.xlsx'
