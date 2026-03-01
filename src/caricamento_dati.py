"""
Caricamento e normalizzazione dei DataFrame (personale, pensionamenti).
Include il mapping delle qualifiche professionali da profili_atto_aziendale.xml
(unico punto di verità per dotazione e raggruppamento qualifiche).
"""

import os
import re
import xml.etree.ElementTree as ET

import pandas as pd

from src.config import FILE_PROFILI_ATTO_AZIENDALE

# ============================================================
# REGEX per pulizia prefisso DESC_TIPO_CDC
# ============================================================
_PREFISSI_CDC = re.compile(
    r'^(OSP\.\s*\S+|DS\s+\S+|DIP\.\s*SALUTE\s+MENTALE|'
    r'DIP\.TO\s+DI\s+\S+|DIP\.TO\s+DI\s+SALUTE\s+MENTALE|'
    r'118|FORMAZIONE|DIREZIONE\s+\S+|'
    r'PREVENZIONE\s+E\s+PROTEZIONE|UOS\b\S*|UOSVD\b\S*)\s*-\s*',
    re.IGNORECASE,
)


# ============================================================
# MAPPING QUALIFICHE (da profili_atto_aziendale.xml)
# ============================================================

def carica_mapping_qualifiche(xml_file):
    """
    Carica le regole di mapping qualifiche da profili_atto_aziendale.xml.

    Ogni <profilo> contiene:
      - <nome_atto>: il valore PROFILO_RAGGRUPPATO risultante
      - <qualifiche_db>/<prefisso>: i prefissi DESC_QUALI che matchano (startswith)

    Restituisce una lista di tuple (prefisso_upper, categoria).
    L'ordine segue quello dell'XML; i prefissi più lunghi vengono
    anteposti per garantire un match più specifico.
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()
    regole = []
    for profilo in root.findall('profilo'):
        categoria = profilo.findtext('nome_atto', '').strip()
        sezione = profilo.find('qualifiche_db')
        if sezione is not None and categoria:
            for pref_elem in sezione.findall('prefisso'):
                prefisso = pref_elem.text.strip().upper() if pref_elem.text else ''
                if prefisso:
                    regole.append((prefisso, categoria))
    # Ordina per lunghezza prefisso decrescente → match più specifico prima
    regole.sort(key=lambda r: len(r[0]), reverse=True)
    return regole


def carica_discipline_overrides(xml_file):
    """
    Carica eventuali override basati su DESC_DISCIPLINE.

    Alcuni profili (es. Avvocati) hanno DESC_QUALI identica ad
    altri raggruppamenti ma si distinguono per DESC_DISCIPLINE.
    Restituisce un dict  {disciplina_upper: nome_atto}.
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()
    overrides = {}
    for profilo in root.findall('profilo'):
        categoria = profilo.findtext('nome_atto', '').strip()
        sezione = profilo.find('discipline_db')
        if sezione is not None and categoria:
            for disc_elem in sezione.findall('disciplina'):
                disc = disc_elem.text.strip().upper() if disc_elem.text else ''
                if disc:
                    overrides[disc] = categoria
    return overrides


# Carica regole all'avvio del modulo
_MAPPING_QUALI_RULES = carica_mapping_qualifiche(FILE_PROFILI_ATTO_AZIENDALE)
_DISCIPLINE_OVERRIDES = carica_discipline_overrides(FILE_PROFILI_ATTO_AZIENDALE)


def mappa_qualifica(desc_quali):
    """
    Mappa un valore DESC_QUALI al raggruppamento professionale
    corrispondente.  Restituisce il valore originale se nessuna
    regola produce un match.
    """
    if pd.isna(desc_quali):
        return 'NON SPECIFICATO'
    val = str(desc_quali).strip().upper()
    for prefisso, categoria in _MAPPING_QUALI_RULES:
        if val.startswith(prefisso):
            return categoria
    return str(desc_quali).strip()


# ============================================================
# CARICAMENTO DATAFRAME
# ============================================================

def carica_dataframe(file_path):
    """
    Carica un file dati in un DataFrame, riconoscendo automaticamente
    il formato dal nome del file (csv, xls, xlsx).
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xls', '.xlsx'):
        return pd.read_excel(file_path)
    else:
        return pd.read_csv(file_path, delimiter=';', encoding='ISO-8859-1')


# ============================================================
# PULIZIA / NORMALIZZAZIONE COLONNE
# ============================================================

def pulisci_prefisso_cdc(cdc):
    """
    Rimuove il prefisso organizzativo ridondante da DESC_TIPO_CDC.

    Esempi:
      'OSP. CARDARELLI - TERAPIA INTENSIVA - DEGENZE ORD.'
        → 'TERAPIA INTENSIVA - DEGENZE ORD.'
      'DS CB - POLIAMBULATORIO CB (VIA PETRELLA)'
        → 'POLIAMBULATORIO CB (VIA PETRELLA)'
    """
    if pd.isna(cdc) or not isinstance(cdc, str):
        return cdc
    risultato = _PREFISSI_CDC.sub('', cdc, count=1).strip()
    return risultato if risultato else cdc


def normalizza_colonne_personale(df):
    """
    Normalizza i nomi delle colonne del file personale per garantire
    compatibilità tra i vecchi CSV e i nuovi XLS.
    """
    rename_map = {}
    if 'DT_ASSUNZIONE' in df.columns and 'PRIMA_DATA_ASSUNZIONE' not in df.columns:
        rename_map['DT_ASSUNZIONE'] = 'PRIMA_DATA_ASSUNZIONE'
    if rename_map:
        df = df.rename(columns=rename_map)

    if 'DESC_NATURA' in df.columns:
        df['DESC_NATURA'] = df['DESC_NATURA'].astype(str)

    # Calcola PROFILO_RAGGRUPPATO
    if 'DESC_QUALI' in df.columns:
        df['PROFILO_RAGGRUPPATO'] = df['DESC_QUALI'].apply(mappa_qualifica)
    elif 'DESC_PROFILO_PROFESSIONALE' in df.columns:
        df['PROFILO_RAGGRUPPATO'] = df['DESC_PROFILO_PROFESSIONALE']

    # Override basati su DESC_DISCIPLINE (es. Avvocati)
    if _DISCIPLINE_OVERRIDES and 'DESC_DISCIPLINE' in df.columns:
        disc_upper = df['DESC_DISCIPLINE'].astype(str).str.strip().str.upper()
        for disc_val, categoria in _DISCIPLINE_OVERRIDES.items():
            mask = disc_upper == disc_val
            if mask.any():
                df.loc[mask, 'PROFILO_RAGGRUPPATO'] = categoria

    return df


def normalizza_colonne_pensionamenti(df):
    """
    Normalizza i nomi delle colonne del file pensionamenti per garantire
    compatibilità tra i vecchi CSV e i nuovi XLS.
    """
    rename_map = {}
    if 'IV_MATRICOLA' in df.columns and 'MATR.' not in df.columns:
        rename_map['IV_MATRICOLA'] = 'MATR.'
    if 'DT' in df.columns and 'DT_CESSAZIONE' not in df.columns:
        rename_map['DT'] = 'DT_CESSAZIONE'
    if rename_map:
        df = df.rename(columns=rename_map)
    return df
