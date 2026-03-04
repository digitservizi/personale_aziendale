"""
Caricamento delle configurazioni da file XML:
mapping ospedali, reparti, intensità, OdC, fabbisogno DM 77, atto aziendale.
"""

import os
import xml.etree.ElementTree as ET


# ============================================================
# MAPPING OSPEDALI
# ============================================================

def carica_mapping_ospedali(xml_path):
    """
    Carica il mapping ospedali da XML.
    Restituisce dict {chiave_db: chiave_personale}.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    mapping = {}
    for osp in root.findall('ospedale'):
        chiave_db = osp.findtext('chiave_db', '').strip()
        chiave_pers = osp.findtext('chiave_personale', '').strip()
        if chiave_db:
            mapping[chiave_db] = chiave_pers
    return mapping


def carica_sigle_ospedali(xml_path):
    """
    Carica la mappa chiave_personale → sigla sede dall'XML ospedali.
    Es: {'CARDARELLI': 'CB', 'SAN TIMOTEO': 'TE', ...}
    Serve per disambiguare match multipli durante la
    generazione di posti_letto.csv.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    sigle = {}
    for osp in root.findall('ospedale'):
        chiave_pers = osp.findtext('chiave_personale', '').strip()
        sigla = osp.findtext('sigla', '').strip()
        if chiave_pers and sigla:
            sigle[chiave_pers] = sigla
    return sigle


# ============================================================
# MAPPING REPARTI
# ============================================================

def carica_mapping_reparti(xml_path):
    """
    Carica il mapping reparti da XML.
    Restituisce dict {nome_db_upper: pattern_regex_o_None}.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    mapping = {}
    for rep in root.find('Reparti').findall('reparto'):
        nome_db = rep.findtext('nome_db', '').strip().upper()
        pattern = rep.findtext('pattern', '')
        if pattern is None or pattern.strip() == '':
            mapping[nome_db] = None
        else:
            mapping[nome_db] = pattern.strip()
    return mapping


def carica_intensita_per_pattern(xml_path):
    """
    Carica la tabella pattern → intensità da XML.
    Restituisce lista di tuple [(pattern_regex, intensita), ...].
    I pattern vengono valutati in ordine: il primo match vince.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    regole = []
    sezione = root.find('IntensitaPerPattern')
    if sezione is not None:
        for regola in sezione.findall('regola'):
            pattern = regola.findtext('pattern', '').strip()
            intensita = regola.findtext('intensita', '').strip()
            if pattern and intensita:
                regole.append((pattern, intensita))
    return regole


# ============================================================
# OSPEDALI DI COMUNITÀ (OdC)
# ============================================================

def carica_mapping_odc(xml_path):
    """
    Carica la configurazione di un Ospedale di Comunità da XML.
    Restituisce dict con nome OdC e lista di strutture.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    nome_odc = root.attrib.get('nome', os.path.basename(xml_path))

    strutture = []
    for s in root.findall('struttura'):
        struttura = {
            'nome': s.findtext('nome', '').strip(),
            'pattern_cdc': s.findtext('pattern_cdc', '').strip(),
            'tipo_sede': s.findtext('tipo_sede', '').strip(),
            'sede_fisica': s.findtext('sede_fisica', '').strip(),
            'intensita': s.findtext('intensita', '').strip(),
            'note': s.findtext('note', '').strip(),
            'posti_letto': {
                'ordinari': int(s.find('posti_letto').findtext('ordinari', '0')),
                'dh': int(s.find('posti_letto').findtext('dh', '0')),
                'ds': int(s.find('posti_letto').findtext('ds', '0')),
            },
        }
        strutture.append(struttura)

    return {'nome': nome_odc, 'strutture': strutture}


# ============================================================
# FABBISOGNO DM 77 / 2022
# ============================================================

def carica_fabbisogno_odc_dm77(xml_path):
    """
    Carica il fabbisogno fisso DM 77/2022 per gli Ospedali di Comunità.

    Restituisce una tupla (fabbisogno, mappa_profili):
      - fabbisogno:    dict {categoria_dm77_upper: unita_int}
      - mappa_profili: dict {profilo_raggruppato_upper: categoria_dm77_upper}
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    fabbisogno = {}
    mappa_profili = {}

    for figura in root.findall('figura'):
        profilo = figura.findtext('profilo', '').strip()
        unita = int(figura.findtext('unita', '0'))
        if not profilo:
            continue
        cat = profilo.upper()
        fabbisogno[cat] = unita

        sez = figura.find('profili_raggruppati')
        if sez is not None:
            for voce in sez.findall('voce'):
                v = (voce.text or '').strip().upper()
                if v:
                    mappa_profili[v] = cat
        else:
            mappa_profili[cat] = cat

    return fabbisogno, mappa_profili


# ============================================================
# MEDICI – ATTO AZIENDALE
# ============================================================

def carica_medici_atto_aziendale(xml_path):
    """
    Carica il mapper medici da atto aziendale.
    Restituisce lista di dict, ciascuno con:
      - nome_atto, dotazione, discipline_db
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    discipline = []
    for disc in root.findall('disciplina'):
        nome_atto = disc.findtext('nome_atto', '').strip()
        dotazione = int(disc.findtext('dotazione', '0'))
        voci_db = []
        sezione_db = disc.find('discipline_db')
        if sezione_db is not None:
            for voce in sezione_db.findall('voce'):
                v = voce.text.strip().upper() if voce.text else ''
                if v:
                    voci_db.append(v)
        discipline.append({
            'nome_atto': nome_atto,
            'dotazione': dotazione,
            'discipline_db': voci_db,
        })
    return discipline


def carica_profili_atto_aziendale(xml_path):
    """
    Carica il mapper profili professionali da atto aziendale.

    Restituisce lista di dict, ciascuno con:
      - nome_atto:      nome profilo (= valore PROFILO_RAGGRUPPATO)
      - dotazione:      dotazione organica da atto aziendale
      - qualifiche_db:  lista di prefissi DESC_QUALI (UPPER) per il match
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    profili = []
    for prof in root.findall('profilo'):
        nome_atto = prof.findtext('nome_atto', '').strip()
        dotazione = int(prof.findtext('dotazione', '0'))
        prefissi = []
        sezione = prof.find('qualifiche_db')
        if sezione is not None:
            for pref in sezione.findall('prefisso'):
                v = pref.text.strip().upper() if pref.text else ''
                if v:
                    prefissi.append(v)
        profili.append({
            'nome_atto': nome_atto,
            'dotazione': dotazione,
            'qualifiche_db': prefissi,
        })
    return profili


# ============================================================
# INDICATORI AGENAS – AREA MATERNO INFANTILE
# ============================================================

def carica_indicatori_agenas_materno_infantile(xml_path):
    """
    Carica gli indicatori AGENAS per l'area materno-infantile.

    Restituisce un dict con:
      - fasce: lista di dict {nome, min, max, profili: {nome: {min, max}}}
      - mapping_uo: lista di dict {pattern, profilo_agenas}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    area = root.find('AreaMaternoInfantile')

    fasce = []
    for fascia in area.findall('fascia'):
        f = {
            'nome': fascia.attrib['nome'],
            'min': int(fascia.attrib['min']),
            'max': int(fascia.attrib['max']),
            'profili': {},
        }
        for profilo in fascia.findall('profilo'):
            f['profili'][profilo.attrib['nome']] = {
                'min': int(profilo.attrib['min']),
                'max': int(profilo.attrib['max']),
            }
        fasce.append(f)

    mapping_uo = []
    sez_uo = area.find('mapping_unita_operative')
    if sez_uo is not None:
        for unita in sez_uo.findall('unita'):
            mapping_uo.append({
                'pattern': unita.attrib['pattern'],
                'profilo_agenas': unita.attrib.get('profilo_agenas'),
            })

    mapping_profili = []
    sez_prof = area.find('mapping_profili')
    if sez_prof is not None:
        for mappa in sez_prof.findall('mappa'):
            mapping_profili.append({
                'pattern': mappa.attrib['pattern'],
                'profilo_agenas': mappa.attrib['profilo_agenas'],
            })

    # Discipline DB (opzionale) per conteggio "fuori area"
    discipline_db = _carica_discipline_db(area)

    return {
        'fasce': fasce,
        'mapping_uo': mapping_uo,
        'mapping_profili': mapping_profili,
        'discipline_db': discipline_db,
    }


# ============================================================
# INDICATORI AGENAS – AREA RADIOLOGIA
# ============================================================

def carica_indicatori_agenas_radiologia(xml_path):
    """
    Carica gli indicatori AGENAS per l'area dei servizi di radiologia.

    Restituisce un dict con:
      - livelli: dict {nome_livello: {profilo: {min, max}}}
      - mapping_uo: lista di dict {pattern}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
    """
    return _carica_indicatori_agenas_per_livello(xml_path)


def carica_indicatori_agenas_trasfusionale(xml_path):
    """
    Carica gli indicatori AGENAS per l'area medicina trasfusionale.

    Restituisce un dict con:
      - livelli: dict {nome_livello: {profilo: {min, max}}}
      - mapping_uo: lista di dict {pattern}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
    """
    return _carica_indicatori_agenas_per_livello(xml_path)


def carica_indicatori_agenas_anatomia_patologica(xml_path):
    """
    Carica gli indicatori AGENAS per l'area anatomia patologica (Tab. 16).

    Restituisce un dict con:
      - livelli: dict {nome_livello: {profilo: {min, max}}}
      - mapping_uo: lista di dict {pattern}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
    """
    return _carica_indicatori_agenas_per_livello(xml_path)


def carica_indicatori_agenas_laboratorio(xml_path):
    """
    Carica gli indicatori AGENAS per l'area servizi di laboratorio (Tab. 14).

    Restituisce un dict con:
      - livelli: dict {nome_livello: {profilo: {min, max}}}
      - mapping_uo: lista di dict {pattern}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
      - esclusioni: lista di pattern regex da escludere
    """
    return _carica_indicatori_agenas_per_livello(xml_path)


def carica_indicatori_agenas_tecnici_laboratorio(xml_path):
    """
    Carica gli indicatori AGENAS per i tecnici di laboratorio (Tab. 17).

    Il conteggio è per ruolo su tutto il presidio (nessun filtro UO).
    Restituisce un dict con:
      - livelli: dict {nome_livello: {profilo: {min, max}}}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
    """
    return _carica_indicatori_agenas_per_livello(xml_path)


def carica_indicatori_agenas_medicina_legale(xml_path):
    """
    Carica gli indicatori AGENAS per la medicina legale (Tab. 18).

    Solo presidi di I e II livello.
    Restituisce un dict con:
      - livelli: dict {nome_livello: {profilo: {min, max}}}
      - mapping_uo: lista di dict {pattern}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
    """
    return _carica_indicatori_agenas_per_livello(xml_path)


def carica_indicatori_agenas_emergenza_urgenza(xml_path):
    """
    Carica gli indicatori AGENAS per l'area emergenza-urgenza (Tab. 20).

    Livelli: OSPEDALE_DI_BASE (PS), PRESIDIO_I_LIVELLO (DEA I),
             PRESIDIO_II_LIVELLO (DEA II).
    Restituisce un dict con:
      - livelli: dict {nome_livello: {profilo: {min, max}}}
      - mapping_uo: lista di dict {pattern}
      - mapping_profili: lista di dict {pattern, profilo_agenas}
    """
    return _carica_indicatori_agenas_per_livello(xml_path)


def carica_fabbisogno_trasfusionale_speciale(xml_path):
    """
    Carica il fabbisogno UOC Medicina Trasfusionale espresso dalla Primaria.

    Restituisce un dict con:
      - sedi: lista di dict {nome, abbreviazione, nota, richiesti: {profilo: int}}
      - mapping_uo: lista di dict {pattern}
      - mapping_profili: lista di dict {chiave, label, pattern}
      - ordine_profili: lista ordinata delle chiavi profilo
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    mapping_uo = []
    sez_uo = root.find('mapping_unita_operative')
    if sez_uo is not None:
        for unita in sez_uo.findall('unita'):
            mapping_uo.append({'pattern': unita.attrib['pattern']})

    mapping_profili = []
    ordine_profili = []
    sez_prof = root.find('mapping_profili')
    if sez_prof is not None:
        for prof in sez_prof.findall('profilo'):
            chiave = prof.attrib['chiave']
            mapping_profili.append({
                'chiave': chiave,
                'label': prof.attrib['label'],
                'pattern': prof.attrib['pattern'],
            })
            ordine_profili.append(chiave)

    sedi = []
    for sede_el in root.findall('sede'):
        richiesti = {}
        for req in sede_el.findall('richiesti'):
            richiesti[req.attrib['profilo']] = int(req.attrib['valore'])
        sedi.append({
            'nome': sede_el.attrib['nome'],
            'abbreviazione': sede_el.attrib.get('abbreviazione', ''),
            'nota': sede_el.attrib.get('nota', ''),
            'richiesti': richiesti,
        })

    return {
        'sedi': sedi,
        'mapping_uo': mapping_uo,
        'mapping_profili': mapping_profili,
        'ordine_profili': ordine_profili,
    }


def _carica_indicatori_agenas_per_livello(xml_path):
    """Caricamento generico indicatori AGENAS basati su livello presidio."""
    tree = ET.parse(xml_path)
    root = tree.getroot()

    livelli = {}
    for livello in root.findall('livello'):
        nome = livello.attrib['nome']
        profili = {}
        for profilo in livello.findall('profilo'):
            profili[profilo.attrib['nome']] = {
                'min': int(profilo.attrib['min']),
                'max': int(profilo.attrib['max']),
            }
        livelli[nome] = profili

    mapping_uo = []
    sez_uo = root.find('mapping_unita_operative')
    if sez_uo is not None:
        for unita in sez_uo.findall('unita'):
            mapping_uo.append({
                'pattern': unita.attrib['pattern'],
            })

    mapping_profili = []
    sez_prof = root.find('mapping_profili')
    if sez_prof is not None:
        for mappa in sez_prof.findall('mappa'):
            mapping_profili.append({
                'pattern': mappa.attrib['pattern'],
                'profilo_agenas': mappa.attrib['profilo_agenas'],
            })

    # Discipline DB (opzionale) per conteggio "fuori area"
    discipline_db = _carica_discipline_db(root)

    return {
        'livelli': livelli,
        'mapping_uo': mapping_uo,
        'mapping_profili': mapping_profili,
        'esclusioni': _carica_esclusioni(root),
        'discipline_db': discipline_db,
    }


def _carica_esclusioni(root):
    """Carica i pattern di esclusione (opzionali) dall'XML."""
    esclusioni = []
    sez = root.find('esclusioni')
    if sez is not None:
        for el in sez.findall('escludi'):
            esclusioni.append(el.attrib['pattern'])
    return esclusioni


def _carica_discipline_db(root):
    """Carica il mapping discipline DB → profilo AGENAS (opzionale).

    Usato per contare il personale "fuori area": dipendenti con
    disciplina coerente all'area ma assegnati a UO diverse.

    Restituisce lista di dict {pattern, profilo_agenas}
    con in più la chiave opzionale 'uo_in_area' (pattern regex)
    che indica le UO che fanno parte dell'area (escluse dal
    conteggio fuori area).
    """
    discipline = []
    uo_in_area_pattern = None
    sez = root.find('discipline_db')
    if sez is not None:
        for el in sez.findall('disciplina'):
            discipline.append({
                'pattern': el.attrib['pattern'],
                'profilo_agenas': el.attrib['profilo_agenas'],
            })
        uo_el = sez.find('uo_in_area')
        if uo_el is not None:
            uo_in_area_pattern = uo_el.attrib.get('pattern', '')
    # Aggiungiamo il pattern uo_in_area a ciascun elemento
    if uo_in_area_pattern:
        for d in discipline:
            d['uo_in_area'] = uo_in_area_pattern
    return discipline


# ============================================================
# INDICATORI TERRITORIALI (basati su popolazione)
# ============================================================

def _carica_indicatori_agenas_territoriale(xml_path):
    """Caricamento generico indicatori AGENAS territoriali (tassi per popolazione).

    Struttura XML attesa:
      <parametri>
        <base_popolazione>10000</base_popolazione>
        <fascia_popolazione>gte_18</fascia_popolazione>
        <titolo>AREA ...</titolo>
        <riferimento>...</riferimento>
      </parametri>
      <unita_operative>
        <uo pattern="CSM"/>
      </unita_operative>
      <profili>
        <profilo nome="..." tasso="1.0">           ← standard singolo
          <mappa qualifica="..."/>
        </profilo>
        <profilo nome="..." tasso_min="3.0" tasso_regime="4.0">  ← range
          <mappa qualifica="..."/>
        </profilo>
      </profili>

    Restituisce un dict con:
      - base_popolazione: int
      - fascia_popolazione: str (chiave di POPOLAZIONE_AREA)
      - titolo: str
      - riferimento: str
      - profili: lista di dict {nome, tasso_min, tasso_regime, nota, qualifiche}
      - unita_operative: lista di str (pattern regex)
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Parametri
    params = root.find('parametri')
    base_pop = int(params.findtext('base_popolazione', '10000'))
    fascia_pop = params.findtext('fascia_popolazione', 'totale')
    titolo = params.findtext('titolo', '')
    riferimento = params.findtext('riferimento', '')

    # Unità operative da filtrare
    uo_patterns = []
    sez_uo = root.find('unita_operative')
    if sez_uo is not None:
        for uo in sez_uo.findall('uo'):
            uo_patterns.append(uo.attrib['pattern'])

    # Profili con tassi
    profili = []
    sez_prof = root.find('profili')
    if sez_prof is not None:
        for prof in sez_prof.findall('profilo'):
            nome = prof.attrib['nome']
            nota = prof.attrib.get('nota', '')

            # Supporta sia tasso singolo sia range min/regime
            if 'tasso' in prof.attrib:
                t = float(prof.attrib['tasso'])
                tasso_min = t
                tasso_regime = None  # nessun limite superiore
            else:
                tasso_min = float(prof.attrib.get('tasso_min', '0'))
                tasso_regime = float(prof.attrib.get('tasso_regime', '0'))

            qualifiche = [m.attrib['qualifica']
                          for m in prof.findall('mappa')]

            profili.append({
                'nome': nome,
                'tasso_min': tasso_min,
                'tasso_regime': tasso_regime,
                'nota': nota,
                'qualifiche': qualifiche,
            })

    return {
        'base_popolazione': base_pop,
        'fascia_popolazione': fascia_pop,
        'titolo': titolo,
        'riferimento': riferimento,
        'profili': profili,
        'unita_operative': uo_patterns,
    }


def carica_indicatori_agenas_salute_mentale(xml_path):
    """Carica indicatori AGENAS per l'area Salute Mentale Adulti."""
    result = _carica_indicatori_agenas_territoriale(xml_path)
    print(f"Indicatori AGENAS salute mentale caricati: "
          f"{len(result['profili'])} profili")
    return result


def carica_indicatori_agenas_dipendenze(xml_path):
    """Carica indicatori AGENAS per l'area Dipendenze Patologiche (SerD)."""
    result = _carica_indicatori_agenas_territoriale(xml_path)
    print(f"Indicatori AGENAS dipendenze (SerD) caricati: "
          f"{len(result['profili'])} profili")
    return result


def carica_indicatori_agenas_npia(xml_path):
    """Carica indicatori AGENAS per l'area NPIA."""
    result = _carica_indicatori_agenas_territoriale(xml_path)
    print(f"Indicatori AGENAS NPIA caricati: "
          f"{len(result['profili'])} profili")
    return result


def carica_indicatori_agenas_carcere(xml_path):
    """Carica indicatori AGENAS per l'area Salute in Carcere."""
    result = _carica_indicatori_agenas_territoriale(xml_path)
    print(f"Indicatori AGENAS salute in carcere caricati: "
          f"{len(result['profili'])} profili")
    return result


# ============================================================
# CARICAMENTO INDICATORI AGENAS – TERAPIA INTENSIVA (§ 8.1.1)
# ============================================================

def carica_indicatori_agenas_terapia_intensiva(xml_path):
    """Carica indicatori AGENAS per l'area Terapia Intensiva (§ 8.1.1).

    Il file XML definisce il rapporto letti/operatore per turno
    anziché range min-max per livello presidio.

    Restituisce un dizionario con:
      - 'standard': lista di {profilo, rapporto_letti, nota}
      - 'mapping_uo': lista di {pattern, note}
      - 'mapping_profili': lista di {pattern, profilo_agenas}
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # --- Standard (rapporto letti per turno) ---
    standard = []
    el_std = root.find('standard')
    if el_std is not None:
        for prof in el_std.findall('profilo'):
            standard.append({
                'profilo': prof.get('nome'),
                'rapporto_letti': int(prof.get('rapporto_letti')),
                'nota': prof.get('nota', ''),
            })

    # --- Mapping UO ---
    mapping_uo = []
    el_muo = root.find('mapping_unita_operative')
    if el_muo is not None:
        for u in el_muo.findall('unita'):
            mapping_uo.append({
                'pattern': u.get('pattern'),
                'note': u.get('note', ''),
            })

    # --- Mapping profili ---
    mapping_profili = []
    el_mp = root.find('mapping_profili')
    if el_mp is not None:
        for m in el_mp.findall('map'):
            mapping_profili.append({
                'pattern': m.get('pattern'),
                'profilo_agenas': m.get('profilo_agenas'),
            })

    # Discipline DB (opzionale) per conteggio "fuori area"
    discipline_db = _carica_discipline_db(root)

    return {
        'standard': standard,
        'mapping_uo': mapping_uo,
        'mapping_profili': mapping_profili,
        'discipline_db': discipline_db,
    }


# ============================================================
# CARICAMENTO INDICATORI AGENAS – SALE OPERATORIE (§ 8.1.2)
# ============================================================

def carica_indicatori_agenas_sale_operatorie(xml_path):
    """Carica indicatori AGENAS per l'area Sale Operatorie (§ 8.1.2).

    Il file XML definisce il personale necessario per ogni sala
    operatoria attiva, i parametri operativi (ore copertura, giorni/anno)
    e i mapping UO/profili.

    Il numero di sale per presidio è in config.py (SALE_OPERATORIE_PER_PRESIDIO).

    Restituisce un dizionario con:
      - 'standard': lista di {profilo, personale_per_sala, nota}
      - 'parametri': {ore_copertura: int, giorni_anno: int}
      - 'mapping_uo': lista di {pattern, note}
      - 'mapping_profili': lista di {pattern, profilo_agenas}
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # --- Parametri operativi ---
    parametri = {'ore_copertura': 24, 'giorni_anno': 250}
    el_par = root.find('parametri')
    if el_par is not None:
        parametri['ore_copertura'] = int(el_par.get('ore_copertura', 24))
        parametri['giorni_anno'] = int(el_par.get('giorni_anno', 250))

    # --- Standard (personale per sala) ---
    standard = []
    el_std = root.find('standard')
    if el_std is not None:
        for prof in el_std.findall('profilo'):
            standard.append({
                'profilo': prof.get('nome'),
                'personale_per_sala': int(prof.get('personale_per_sala')),
                'nota': prof.get('nota', ''),
            })

    # --- Mapping UO ---
    mapping_uo = []
    el_muo = root.find('mapping_unita_operative')
    if el_muo is not None:
        for u in el_muo.findall('unita'):
            mapping_uo.append({
                'pattern': u.get('pattern'),
                'note': u.get('note', ''),
            })

    # --- Mapping profili ---
    mapping_profili = []
    el_mp = root.find('mapping_profili')
    if el_mp is not None:
        for m in el_mp.findall('map'):
            mapping_profili.append({
                'pattern': m.get('pattern'),
                'profilo_agenas': m.get('profilo_agenas'),
            })

    # Discipline DB (opzionale) per conteggio "fuori area"
    discipline_db = _carica_discipline_db(root)

    return {
        'standard': standard,
        'parametri': parametri,
        'mapping_uo': mapping_uo,
        'mapping_profili': mapping_profili,
        'discipline_db': discipline_db,
    }
