"""
Genera un documento PDF autonomo «Nota Metodologica».

Il documento contiene:
  1. Titolo / copertina
  2. Indicatori di calcolo per figura professionale
  3. Posti letto per presidio (una tabella per sede)
  4. Formula di calcolo e regole di arrotondamento
  5. Standard AGENAS ospedalieri (aree principali)
  6a. Area Salute Mentale  – tabella indicatori + popolazione per distretto
  6b. Area Dipendenze SerD – tabella indicatori + popolazione per distretto
  6c. Area NPIA            – tabella indicatori + popolazione per distretto
  6d. Salute in Carcere    – tabella indicatori + detenuti per istituto
  7. Dati di ingresso: presidi, parti, popolazione, detenuti
"""

from __future__ import annotations

import os
from collections import OrderedDict

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.flowables import HRFlowable

from src.config import (
    DETENUTI_PER_ISTITUTO,
    LIVELLO_PRESIDIO,
    PARTI_PER_PRESIDIO,
    POPOLAZIONE_AREA,
    ANNO_ANALISI,
    DIR_ELABORATI,
)

# ---------------------------------------------------------------------------
# Colori
# ---------------------------------------------------------------------------
CLR_HEADER   = colors.HexColor('#1F4E79')   # blu scuro titoli
CLR_SUBHDR   = colors.HexColor('#2E75B6')   # blu medio sotto-sezioni
CLR_TH_BG    = colors.HexColor('#1F4E79')   # sfondo intestazioni tabelle
CLR_TH_FG    = colors.white
CLR_ROW_A    = colors.HexColor('#DEEAF1')   # riga alternata A
CLR_ROW_B    = colors.white                 # riga alternata B
CLR_ROW_TOT  = colors.HexColor('#BDD7EE')   # riga totale
CLR_FORMULA  = colors.HexColor('#FFF2CC')   # sfondo formule
CLR_LINE     = colors.HexColor('#1F4E79')   # linea separatrice
CLR_LIGHT    = colors.HexColor('#555555')   # testo secondario/fonte
CLR_COVER_BG = colors.HexColor('#1F4E79')   # copertina

# ---------------------------------------------------------------------------
# Stili testo
# ---------------------------------------------------------------------------
_ss = getSampleStyleSheet()

ST_TITLE = ParagraphStyle(
    'NM_Title', parent=_ss['Title'],
    fontSize=22, leading=28,
    textColor=colors.white, alignment=TA_CENTER,
    spaceAfter=6,
)
ST_SUBTITLE = ParagraphStyle(
    'NM_Subtitle', parent=_ss['Normal'],
    fontSize=13, leading=17,
    textColor=colors.HexColor('#BDD7EE'), alignment=TA_CENTER,
    spaceAfter=4,
)
ST_SECTION = ParagraphStyle(
    'NM_Section', parent=_ss['Heading1'],
    fontSize=13, leading=17,
    textColor=CLR_HEADER, spaceBefore=10, spaceAfter=4,
    fontName='Helvetica-Bold',
)
ST_SUBSECT = ParagraphStyle(
    'NM_Subsect', parent=_ss['Heading2'],
    fontSize=11, leading=14,
    textColor=CLR_SUBHDR, spaceBefore=8, spaceAfter=3,
    fontName='Helvetica-Bold',
)
ST_BODY = ParagraphStyle(
    'NM_Body', parent=_ss['Normal'],
    fontSize=9, leading=13,
    textColor=colors.black, spaceAfter=4,
    alignment=TA_JUSTIFY,
)
ST_NOTE = ParagraphStyle(
    'NM_Note', parent=_ss['Normal'],
    fontSize=8, leading=11,
    textColor=CLR_LIGHT, spaceAfter=3,
    fontName='Helvetica-Oblique',
)
ST_FONTE = ParagraphStyle(
    'NM_Fonte', parent=_ss['Normal'],
    fontSize=8, leading=11,
    textColor=CLR_LIGHT, spaceAfter=2,
    fontName='Helvetica-Oblique',
)
ST_TH = ParagraphStyle(
    'NM_TH', parent=_ss['Normal'],
    fontSize=8, leading=10,
    textColor=CLR_TH_FG, alignment=TA_CENTER,
    fontName='Helvetica-Bold',
)
ST_TD = ParagraphStyle(
    'NM_TD', parent=_ss['Normal'],
    fontSize=8, leading=11,
    textColor=colors.black,
)
ST_TD_C = ParagraphStyle(
    'NM_TDC', parent=_ss['Normal'],
    fontSize=8, leading=11,
    textColor=colors.black, alignment=TA_CENTER,
)
ST_TD_BOLD = ParagraphStyle(
    'NM_TDBold', parent=_ss['Normal'],
    fontSize=8, leading=11,
    textColor=colors.black, fontName='Helvetica-Bold',
)
ST_FORMULA = ParagraphStyle(
    'NM_Formula', parent=_ss['Code'],
    fontSize=10, leading=14,
    textColor=colors.black, alignment=TA_CENTER,
    fontName='Courier-Bold', spaceAfter=4,
)

# ---------------------------------------------------------------------------
# Helper stile tabella di base
# ---------------------------------------------------------------------------
def _base_ts(n_rows: int, n_cols: int,
             header_rows: int = 1,
             alt_start: int = 1) -> list:
    """Restituisce una lista di comandi TableStyle comuni."""
    cmds = [
        # Header
        ('BACKGROUND', (0, 0), (n_cols - 1, header_rows - 1), CLR_TH_BG),
        ('TEXTCOLOR', (0, 0), (n_cols - 1, header_rows - 1), CLR_TH_FG),
        ('FONTNAME', (0, 0), (n_cols - 1, header_rows - 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (n_cols - 1, header_rows - 1), 8),
        ('ALIGN', (0, 0), (n_cols - 1, header_rows - 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#AAAAAA')),
        ('ROWBACKGROUND', (0, alt_start), (-1, n_rows - 1),
         [CLR_ROW_A, CLR_ROW_B]),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]
    return cmds


def _p(text, style=None):
    """Shortcut Paragraph."""
    return Paragraph(str(text), style or ST_TD)


def _ph(text):
    """Paragraph intestazione tabella."""
    return Paragraph(str(text), ST_TH)


# ---------------------------------------------------------------------------
# Classe principale
# ---------------------------------------------------------------------------

class NotaMetodologicaPDF:
    """Genera il PDF della nota metodologica."""

    PAGE_W, PAGE_H = landscape(A3)
    MARGIN = 15 * mm

    def __init__(self, indicators: dict, posti_letto_citta: dict,
                 anno: int | None = None,
                 output_dir: str = DIR_ELABORATI):
        self.indicators = indicators
        self.posti_letto_citta = posti_letto_citta
        self.anno = anno or ANNO_ANALISI
        self.output_dir = output_dir
        self.avail_w = self.PAGE_W - 2 * self.MARGIN

    # ------------------------------------------------------------------
    # Entry point
    # ------------------------------------------------------------------

    def genera(self) -> str:
        """Genera il PDF e restituisce il percorso del file."""
        os.makedirs(self.output_dir, exist_ok=True)
        path = os.path.join(self.output_dir,
                            f'nota_metodologica_{self.anno}.pdf')

        doc = BaseDocTemplate(
            path,
            pagesize=landscape(A3),
            leftMargin=self.MARGIN,
            rightMargin=self.MARGIN,
            topMargin=self.MARGIN,
            bottomMargin=self.MARGIN,
        )

        # Frame unico per tutte le pagine
        frame = Frame(
            self.MARGIN, self.MARGIN,
            self.PAGE_W - 2 * self.MARGIN,
            self.PAGE_H - 2 * self.MARGIN,
            id='main',
        )

        # Template standard con footer
        def _footer(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 7)
            canvas.setFillColor(CLR_LIGHT)
            y = self.MARGIN * 0.6
            canvas.drawString(self.MARGIN, y,
                              f'ASREM – Nota Metodologica {self.anno}')
            canvas.drawRightString(self.PAGE_W - self.MARGIN, y,
                                   f'Pagina {doc.page}')
            canvas.restoreState()

        # Prima pagina: copertina con sfondo blu
        def _cover_page(canvas, doc):
            canvas.saveState()
            # Sfondo copertina
            canvas.setFillColor(CLR_COVER_BG)
            canvas.rect(0, 0, self.PAGE_W, self.PAGE_H, fill=1, stroke=0)
            canvas.restoreState()

        tpl_cover = PageTemplate(id='cover', frames=[frame],
                                 onPage=_cover_page)
        tpl_normal = PageTemplate(id='normal', frames=[frame],
                                  onPage=_footer)
        doc.addPageTemplates([tpl_cover, tpl_normal])

        story = []

        # ── Pag. 1: Copertina ────────────────────────────────────────
        story += self._copertina()
        story.append(NextPageTemplate('normal'))
        story.append(PageBreak())

        # ── Pag. 2: Indicatori per figura professionale ───────────────
        story += self._sezione_indicatori()
        story.append(PageBreak())

        # ── Pag. 3: Posti letto per presidio ─────────────────────────
        story += self._sezione_posti_letto()
        story.append(PageBreak())

        # ── Pag. 4: Formula + Arrotondamenti + Standard AGENAS osp. ──
        story += self._sezione_formula()
        story += self._sezione_standard_ospedalieri()
        story.append(PageBreak())

        # ── Pag. 5: 6a Salute Mentale + 6b SerD ──────────────────────
        story += self._sezione_territoriale_sm()
        story += self._sezione_territoriale_serd()
        story.append(PageBreak())

        # ── Pag. 6: 6c NPIA + 6d Carcere ────────────────────────────
        story += self._sezione_territoriale_npia()
        story += self._sezione_territoriale_carcere()
        story.append(PageBreak())

        # ── Pag. 7: Dati di ingresso ──────────────────────────────────
        story += self._sezione_dati_ingresso()

        doc.build(story)
        return path

    # ------------------------------------------------------------------
    # Copertina
    # ------------------------------------------------------------------

    def _copertina(self) -> list:
        items = []
        items.append(Spacer(1, 40 * mm))
        items.append(Paragraph(
            f'NOTA METODOLOGICA', ST_TITLE))
        items.append(Paragraph(
            f'Criteri di calcolo del fabbisogno di personale', ST_SUBTITLE))
        items.append(Spacer(1, 6 * mm))
        items.append(Paragraph(
            f'Anno di riferimento: <b>{self.anno}</b>', ST_SUBTITLE))
        items.append(Spacer(1, 20 * mm))
        items.append(Paragraph(
            'ASREM – Azienda Sanitaria Regionale del Molise', ST_SUBTITLE))
        return items

    # ------------------------------------------------------------------
    # Sezione 1 – Indicatori per figura professionale
    # ------------------------------------------------------------------

    def _sezione_indicatori(self) -> list:
        items: list = []
        items.append(Paragraph(
            '1. Indicatori di calcolo per figura professionale',
            ST_SECTION))
        items.append(Paragraph(
            'Parametri utilizzati nella formula di calcolo del fabbisogno '
            'per ciascun profilo professionale e livello di intensità '
            'assistenziale.', ST_BODY))
        items.append(Spacer(1, 3 * mm))

        PROFILE_LABELS = {
            'DIRIGENTE_MEDICO': 'Dirigente Medico',
            'INFERMIERE': 'Infermiere',
            'OPERATORE_SOCIO_SANITARIO': 'Operatore Socio Sanitario',
            'OSTETRICA': 'Ostetrica',
            'TS_RADIOLOGIA': 'Tecnico Sanitario Radiologia',
            'TS_LABORATORIO': 'Tecnico Sanitario Laboratorio',
        }
        INTENSITY_LABELS = {
            'Intensiva': 'Intensiva',
            'Alta': 'Alta',
            'MedioAlta': 'Medio/Alta',
            'Media': 'Media',
            'MedioBassa': 'Medio/Bassa',
            'Bassa': 'Bassa',
            'DH_DS': 'DH / Day Surgery',
        }
        INTENSITY_ORDER = [
            'Intensiva', 'Alta', 'MedioAlta', 'Media',
            'MedioBassa', 'Bassa', 'DH_DS',
        ]
        IND_HEADERS = [
            'Intensità', 'Tasso Occupazione',
            'Coeff. Complessità', 'Ore per Turno', 'Ore Annue Effettive',
        ]
        col_w = [
            self.avail_w * 0.22,
            self.avail_w * 0.18,
            self.avail_w * 0.20,
            self.avail_w * 0.18,
            self.avail_w * 0.22,
        ]

        for prof_key, prof_label in PROFILE_LABELS.items():
            if prof_key not in self.indicators:
                continue
            prof_data = self.indicators[prof_key]
            rows = [[_ph(h) for h in IND_HEADERS]]
            for int_key in INTENSITY_ORDER:
                if int_key not in prof_data:
                    continue
                d = prof_data[int_key]
                rows.append([
                    _p(INTENSITY_LABELS.get(int_key, int_key)),
                    _p(f"{d['TassoOccupazione']:.0f}%", ST_TD_C),
                    _p(str(d['CoefficienteComplessita']), ST_TD_C),
                    _p(str(d['OreEffettuateTurni']), ST_TD_C),
                    _p(str(d['OreAnnueLavoroEffettivo']), ST_TD_C),
                ])
            n = len(rows)
            ts_cmds = _base_ts(n, 5)
            t = Table(rows, colWidths=col_w, repeatRows=1)
            t.setStyle(TableStyle(ts_cmds))
            items.append(KeepTogether([
                Paragraph(prof_label, ST_SUBSECT),
                t,
                Spacer(1, 4 * mm),
            ]))

        return items

    # ------------------------------------------------------------------
    # Sezione 2 – Posti letto per presidio
    # ------------------------------------------------------------------

    def _sezione_posti_letto(self) -> list:
        items: list = []
        items.append(Paragraph('2. Posti letto rilevati per presidio',
                                ST_SECTION))
        items.append(Paragraph(
            'Posti letto per sede e unità operativa con il livello di '
            'intensità assistenziale assegnato. Solo le unità con posti '
            'letto attivi generano un fabbisogno calcolato; le altre sono '
            'indicate come "Servizio privo di posti letto".', ST_BODY))
        items.append(Spacer(1, 3 * mm))

        PL_HEADERS = [
            'Reparto (SSD)', 'Ordinari', 'DH', 'UTIC', 'DS', 'Intensità',
        ]
        col_w = [
            self.avail_w * 0.42,
            self.avail_w * 0.10,
            self.avail_w * 0.10,
            self.avail_w * 0.10,
            self.avail_w * 0.10,
            self.avail_w * 0.18,
        ]

        # Raggruppa per sede – solo reparti con posti letto attivi
        sedi_pl: OrderedDict = OrderedDict()
        for (sede, ssd), pl in sorted(
                self.posti_letto_citta.items(),
                key=lambda x: (x[0][0], x[0][1])):
            tot = (int(pl.get('ordinari', 0)) + int(pl.get('dh', 0))
                   + int(pl.get('utic', 0)) + int(pl.get('ds', 0)))
            if tot > 0:
                sedi_pl.setdefault(sede, []).append((ssd, pl))

        for sede_nome, reparti in sedi_pl.items():
            t_rows = [[_ph(h) for h in PL_HEADERS]]
            for ssd, pl in reparti:
                t_rows.append([
                    _p(ssd),
                    _p(str(int(pl['ordinari'])), ST_TD_C),
                    _p(str(int(pl['dh'])), ST_TD_C),
                    _p(str(int(pl.get('utic', 0))), ST_TD_C),
                    _p(str(int(pl.get('ds', 0))), ST_TD_C),
                    _p(pl.get('intensita', ''), ST_TD_C),
                ])
            n = len(t_rows)
            ts_cmds = _base_ts(n, len(PL_HEADERS))
            t = Table(t_rows, colWidths=col_w, repeatRows=1)
            t.setStyle(TableStyle(ts_cmds))
            items.append(KeepTogether([
                Paragraph(sede_nome, ST_SUBSECT),
                t,
                Spacer(1, 4 * mm),
            ]))

        return items

    # ------------------------------------------------------------------
    # Sezione 3 – Formula e arrotondamenti
    # ------------------------------------------------------------------

    def _sezione_formula(self) -> list:
        items: list = []
        items.append(Paragraph(
            '3. Formula di calcolo del fabbisogno', ST_SECTION))
        items.append(Paragraph(
            'Il fabbisogno teorico per ciascun reparto ospedaliero è '
            'calcolato con la seguente formula:', ST_BODY))
        items.append(Spacer(1, 2 * mm))

        formula_box_data = [[Paragraph(
            'Fabbisogno = (PL_omog × Tasso_occ × Coeff_compl × Ore_turni × 365) '
            '/ Ore_annue_eff',
            ST_FORMULA,
        )]]
        ft = Table(formula_box_data,
                   colWidths=[self.avail_w * 0.8])
        ft.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), CLR_FORMULA),
            ('BOX', (0, 0), (-1, -1), 0.5, CLR_LINE),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        items.append(ft)
        items.append(Spacer(1, 3 * mm))

        DEFS = [
            ('PL_omog',
             'Posti Letto Omogeneizzati = Ordinari×1 + DH×0,5 + UTIC×1 + '
             'Breast Unit Ord.×1 + Breast Unit DH×0,5'),
            ('Tasso_occ', 'Tasso di occupazione (da tabella indicatori, in %)'),
            ('Coeff_compl', 'Coefficiente di complessità assistenziale'),
            ('Ore_turni', 'Ore effettuate per turno'),
            ('365', 'Giorni di assistenza annui'),
            ('Ore_annue_eff', 'Ore annue di lavoro effettivo per operatore'),
        ]
        def_rows = [
            [_ph('Simbolo'), _ph('Descrizione')],
        ] + [
            [_p(f'<b>{s}</b>', ST_TD_BOLD), _p(d)]
            for s, d in DEFS
        ]
        col_w_def = [self.avail_w * 0.18, self.avail_w * 0.62]
        dt = Table(def_rows, colWidths=col_w_def, repeatRows=1)
        dt.setStyle(TableStyle(_base_ts(len(def_rows), 2)))
        items.append(dt)
        items.append(Spacer(1, 4 * mm))

        # Arrotondamento
        items.append(Paragraph('4. Regole di arrotondamento', ST_SECTION))
        items.append(Paragraph(
            'Il fabbisogno calcolato è un valore decimale. Poiché il '
            'personale è indivisibile, si applicano le seguenti regole:',
            ST_BODY))
        items.append(Spacer(1, 2 * mm))

        RULES = [
            ('Fabbisogno < 1 (profilo assente)',
             'Se il profilo non ha personale in servizio nel reparto, '
             'il fabbisogno minimo è fissato a 1 unità.'),
            ('Fabbisogno < 1 (profilo presente)',
             'Arrotondamento standard: ≥ 0,5 → 1; < 0,5 → 0.'),
            ('Fabbisogno ≥ 1',
             "Arrotondamento all'intero più vicino: parte decimale "
             '≥ 0,5 → eccesso; < 0,5 → difetto.'),
        ]
        rule_rows = [[_ph('Caso'), _ph('Regola applicata')]] + [
            [_p(c, ST_TD_BOLD), _p(r)]
            for c, r in RULES
        ]
        col_w_r = [self.avail_w * 0.28, self.avail_w * 0.52]
        rt = Table(rule_rows, colWidths=col_w_r, repeatRows=1)
        rt.setStyle(TableStyle(_base_ts(len(rule_rows), 2)))
        items.append(rt)
        items.append(Spacer(1, 6 * mm))

        return items

    # ------------------------------------------------------------------
    # Sezione 5 – Standard AGENAS ospedalieri (tabella riepilogativa)
    # ------------------------------------------------------------------

    def _sezione_standard_ospedalieri(self) -> list:
        items: list = []
        items.append(Paragraph(
            '5. Standard AGENAS – Aree ospedaliere (riepilogo)',
            ST_SECTION))
        items.append(Paragraph(
            'Le tabelle AGENAS definiscono dotazioni minime e massime di '
            'personale (FTE) per area funzionale e livello di presidio '
            '(DM 70/2015). A tutti gli standard è applicata una '
            'maggiorazione organica del +15% per copertura della '
            'turnazione, guardie, ferie, malattie.', ST_BODY))
        items.append(Spacer(1, 3 * mm))

        # Tabella riepilogativa per area
        AREAS = [
            ('5a. Materno Infantile (Tab. 11)',
             'N. parti per presidio',
             '500–1.500 / 1.500–2.000 / >2.000 parti',
             'Dir.Med.Pediatria, Dir.Med.Ostetricia, Ostetriche, Infermieri, OSS'),
            ('5b. Radiologia (Tab. 13)',
             'Livello presidio',
             'Ospedale di Base / I Livello / II Livello',
             'Dir.Medici, Tecnici Rad., Infermieri, OSS'),
            ('5c. Laboratorio (Tab. 14)',
             'Livello presidio',
             'Ospedale di Base / I Livello / II Livello',
             'Dir.Sanitari, Infermieri, OSS'),
            ('5d. Trasfusionale (Tab. 15)',
             'Livello presidio',
             'I Livello / II Livello',
             'Dir.Sanitari, Infermieri, OSS'),
            ('5e. Anatomia Patologica (Tab. 16)',
             'Livello presidio',
             'Ospedale di Base / I Livello / II Livello',
             'Dir.Sanitari, Infermieri, OSS'),
            ('5f. Tecnici Laboratorio (Tab. 17)',
             'Livello presidio',
             'Ospedale di Base / I Livello / II Livello',
             'TS Lab. Biomedico'),
            ('5g. Medicina Legale (Tab. 18)',
             'Livello presidio',
             'I Livello / II Livello',
             'Dir.Medici, Infermieri, OSS'),
            ('5h. Emergenza-Urgenza (Tab. 20)',
             'Livello presidio',
             'Base (PS) / I (DEA I) / II (DEA II)',
             'Dir.Medici, Infermieri, OSS'),
            ('5i. Terapia Intensiva (§ 8.1.1)',
             'PL intensivi',
             'Formula FTE: PL/rapporto × 24×365/Ore_annue × 1,15',
             '1 Medico/8 letti, 1 Infermiere/2 letti'),
            ('5l. Sale Operatorie (§ 8.1.2)',
             'N. sale',
             'Formula FTE: N.Sale × Staff × Ore_copertura/Ore_annue × 1,15',
             '1 Anestesista, 3 Infermieri, 1 OSS per sala'),
        ]
        hdrs = ['Area', 'Parametro ingresso', 'Fasce/Formula', 'Profili inclusi']
        col_w = [
            self.avail_w * 0.20,
            self.avail_w * 0.16,
            self.avail_w * 0.30,
            self.avail_w * 0.34,
        ]
        rows = [[_ph(h) for h in hdrs]] + [
            [_p(a), _p(b), _p(c), _p(d)]
            for a, b, c, d in AREAS
        ]
        t = Table(rows, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle(_base_ts(len(rows), 4)))
        items.append(t)
        items.append(Spacer(1, 6 * mm))

        return items

    # ------------------------------------------------------------------
    # Helper comune per sezioni territoriali
    # ------------------------------------------------------------------

    def _tabella_indicatori_territoriale(
            self,
            titolo: str,
            fonte: str,
            descrizione: str,
            base_pop: int,
            fascia_label: str,
            uo_patterns: str,
            profili_tassi: list,
            is_range: bool = False,
    ) -> list:
        """Genera i flowable comuni (titolo, descrizione, tabella indicatori)."""
        items: list = []
        items.append(HRFlowable(width=self.avail_w, thickness=1,
                                color=CLR_LINE, spaceAfter=2))
        items.append(Paragraph(titolo, ST_SUBSECT))
        items.append(Paragraph(f'Fonte: {fonte}', ST_FONTE))
        items.append(Paragraph(descrizione, ST_BODY))
        items.append(Paragraph(
            f'Base popolazione: <b>{base_pop:,}</b> – '
            f'Fascia: <b>{fascia_label}</b> – '
            f'UO incluse: {uo_patterns}',
            ST_NOTE))
        items.append(Spacer(1, 2 * mm))

        if is_range:
            hdrs = ['Profilo', 'Tasso Min', 'Tasso Regime', 'Qualifiche incluse']
            col_w = [
                self.avail_w * 0.22,
                self.avail_w * 0.10,
                self.avail_w * 0.12,
                self.avail_w * 0.36,
            ]
        else:
            hdrs = ['Profilo', 'Tasso', 'Qualifiche incluse']
            col_w = [
                self.avail_w * 0.22,
                self.avail_w * 0.10,
                self.avail_w * 0.48,
            ]

        rows = [[_ph(h) for h in hdrs]]
        for pt in profili_tassi:
            if is_range:
                nome, t_min, t_max, quals = pt
                rows.append([_p(nome), _p(str(t_min), ST_TD_C),
                              _p(str(t_max), ST_TD_C), _p(quals)])
            else:
                nome, tasso, quals = pt
                rows.append([_p(nome), _p(str(tasso), ST_TD_C), _p(quals)])

        t = Table(rows, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle(_base_ts(len(rows), len(hdrs))))
        items.append(t)
        items.append(Spacer(1, 3 * mm))
        return items

    def _tabella_popolazione_distretto(
            self,
            fascia_key: str,
            fascia_label: str,
            base_pop: int,
            profili_tassi: list,
            is_range: bool = False,
    ) -> list:
        """
        Tabella con una riga per distretto che mostra:
          - Area distrettuale
          - Popolazione fascia
          - Per ogni profilo: atteso min (e atteso regime se is_range)
        """
        items: list = []
        items.append(Paragraph(
            f'Popolazione di riferimento per distretto – {fascia_label}',
            ST_NOTE))

        # Intestazioni: Area | Pop.fascia | profilo1_min [| profilo1_regime] | ...
        hdrs = [f'Area distrettuale', f'Pop. {fascia_label}']
        if is_range:
            for nome, t_min, t_max, _ in profili_tassi:
                hdrs.append(f'{nome}\nAtteso min')
                hdrs.append(f'{nome}\nAtteso regime')
        else:
            for nome, tasso, _ in profili_tassi:
                hdrs.append(f'{nome}\nAtteso')

        n_prof_cols = len(hdrs) - 2
        col_first = self.avail_w * 0.20
        col_pop = self.avail_w * 0.12
        col_rest = (self.avail_w - col_first - col_pop) / max(n_prof_cols, 1)
        col_w = [col_first, col_pop] + [col_rest] * n_prof_cols

        rows = [[_ph(h) for h in hdrs]]

        for area, dati in POPOLAZIONE_AREA.items():
            pop_raw = dati.get(fascia_key, 0)
            # gestione totale come float decimale (es. 116.489 = 116489)
            if fascia_key == 'totale' and isinstance(pop_raw, float) and pop_raw < 1000:
                pop_raw = int(pop_raw * 1000)
            pop = int(pop_raw)

            row_vals = [_p(area, ST_TD_BOLD), _p(f'{pop:,}', ST_TD_C)]
            if is_range:
                for _, t_min, t_max, _ in profili_tassi:
                    att_min = pop * t_min / base_pop
                    att_max = pop * t_max / base_pop
                    row_vals.append(_p(f'{att_min:.2f}', ST_TD_C))
                    row_vals.append(_p(f'{att_max:.2f}', ST_TD_C))
            else:
                for _, tasso, _ in profili_tassi:
                    att = pop * tasso / base_pop
                    row_vals.append(_p(f'{att:.2f}', ST_TD_C))
            rows.append(row_vals)

        # Riga TOTALE
        totale_pop = 0
        for area, dati in POPOLAZIONE_AREA.items():
            pop_raw = dati.get(fascia_key, 0)
            if fascia_key == 'totale' and isinstance(pop_raw, float) and pop_raw < 1000:
                pop_raw = int(pop_raw * 1000)
            totale_pop += int(pop_raw)

        tot_row = [_p('TOTALE ASREM', ST_TD_BOLD),
                   _p(f'{totale_pop:,}', ST_TD_C)]
        if is_range:
            for _, t_min, t_max, _ in profili_tassi:
                tot_row.append(_p(f'{totale_pop * t_min / base_pop:.2f}',
                                  ST_TD_BOLD))
                tot_row.append(_p(f'{totale_pop * t_max / base_pop:.2f}',
                                  ST_TD_BOLD))
        else:
            for _, tasso, _ in profili_tassi:
                tot_row.append(_p(f'{totale_pop * tasso / base_pop:.2f}',
                                  ST_TD_BOLD))
        rows.append(tot_row)

        n = len(rows)
        ts_cmds = _base_ts(n, len(hdrs))
        # Riga totale: sfondo distinto
        ts_cmds.append(
            ('BACKGROUND', (0, n - 1), (-1, n - 1), CLR_ROW_TOT))
        ts_cmds.append(
            ('FONTNAME', (0, n - 1), (-1, n - 1), 'Helvetica-Bold'))

        t = Table(rows, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle(ts_cmds))
        items.append(t)
        items.append(Spacer(1, 5 * mm))
        return items

    # ------------------------------------------------------------------
    # 6a. Salute Mentale
    # ------------------------------------------------------------------

    def _sezione_territoriale_sm(self) -> list:
        items: list = []
        items.append(Paragraph(
            '6. Standard AGENAS – Servizi territoriali', ST_SECTION))
        items.append(Paragraph(
            'Gli standard AGENAS per i servizi territoriali sono espressi '
            'come tassi di operatori per popolazione di riferimento. '
            'La formula è: <b>Atteso = Popolazione_fascia × Tasso / Base</b>.',
            ST_BODY))
        items.append(Spacer(1, 3 * mm))

        SM_PROFILI = [
            ('Medici Psichiatri', 1.0,
             'DIRIGENTE MEDICO'),
            ('Psicologi Psicoterapeuti', 0.5,
             'DIRIGENTE PSICOLOGO'),
            ('Prof. Sanitarie + Ass. Sociali', 5.0,
             'INFERMIERE, TRP, ASSISTENTE SOCIALE, EDUCATORE PROF., OSS'),
            ('Altro Personale', 0.2,
             'COLL. AMM.VO, ASS. AMM.VO, OPERATORE TECNICO'),
        ]
        block = (
            self._tabella_indicatori_territoriale(
                titolo='6a. Area Salute Mentale Adulti',
                fonte='DPR 1/11/1999 + AGENAS – indicatori_agenas_salute_mentale.xml',
                descrizione=(
                    'Tasso su 10.000 abitanti ≥ 18 anni. '
                    'Totale minimo: ≥ 6,7 operatori / 10.000 ab. '
                    'Il totale dei profili (1,0 + 0,5 + 5,0 + 0,2 = 6,7) '
                    'rappresenta la soglia minima complessiva.'
                ),
                base_pop=10_000,
                fascia_label='residenti ≥ 18 anni',
                uo_patterns='CSM | SPDC | SALUTE MENTALE',
                profili_tassi=SM_PROFILI,
            )
            + self._tabella_popolazione_distretto(
                fascia_key='gte_18',
                fascia_label='residenti ≥ 18 anni',
                base_pop=10_000,
                profili_tassi=SM_PROFILI,
            )
        )
        items.append(KeepTogether(block))
        return items

    # ------------------------------------------------------------------
    # 6b. Dipendenze SerD
    # ------------------------------------------------------------------

    def _sezione_territoriale_serd(self) -> list:
        items: list = []
        SERD_PROFILI = [
            ('Medico', 3.0, 4.0, 'DIRIGENTE MEDICO'),
            ('Psicologo', 3.0, 3.5, 'DIRIGENTE PSICOLOGO'),
            ('Infermiere', 4.0, 6.0, 'INFERMIERE'),
            ('Educatore / TeRP', 2.5, 3.5, 'EDUCATORE PROF., TRP'),
            ('Assistente Sociale', 2.0, 3.0, 'ASSISTENTE SOCIALE'),
            ('Amministrativo', 0.5, 1.0,
             'COLL. AMM.VO, ASS. AMM.VO, COLL. TECNICO'),
        ]
        block = (
            self._tabella_indicatori_territoriale(
                titolo='6b. Area Dipendenze Patologiche (SerD)',
                fonte='Tabella 1 AGENAS – indicatori_agenas_dipendenze.xml',
                descrizione=(
                    'Tasso su 100.000 residenti 15–64 anni. '
                    'Ogni profilo ha due valori: standard minimo e standard '
                    'a regime. Esito a 3 stati: CARENZA (sotto il min), '
                    'IN RANGE (tra min e regime), A REGIME (≥ regime).'
                ),
                base_pop=100_000,
                fascia_label='residenti 15–64 anni',
                uo_patterns='DIPENDENZ | SERD',
                profili_tassi=SERD_PROFILI,
                is_range=True,
            )
            + self._tabella_popolazione_distretto(
                fascia_key='range_15_64',
                fascia_label='residenti 15–64 anni',
                base_pop=100_000,
                profili_tassi=SERD_PROFILI,
                is_range=True,
            )
        )
        items.append(KeepTogether(block))
        return items

    # ------------------------------------------------------------------
    # 6c. NPIA
    # ------------------------------------------------------------------

    def _sezione_territoriale_npia(self) -> list:
        items: list = []
        NPIA_PROFILI = [
            ('Dirigenza Sanitaria', 6.0,
             'DIR. MEDICO NPI + DIR. PSICOLOGO'),
            ('Prof. San. + Ass. Sociali', 10.0,
             'INFERMIERE, TNPEE, EDUCATORE, TRP, FISIOTERAPISTA, '
             'LOGOPEDISTA, ASS. SOC.'),
            ('Altro Personale', 0.2,
             'COLL. AMM.VO, ASS. AMM.VO, OP. TECNICO'),
        ]
        block = (
            self._tabella_indicatori_territoriale(
                titolo='6c. Area NPIA – Neuropsichiatria Infanzia e Adolescenza',
                fonte='Documento AGENAS – indicatori_agenas_npia.xml',
                descrizione=(
                    'Tasso su 10.000 abitanti 1–17 anni. '
                    'Il profilo «Dirigenza Sanitaria» aggrega medici NPI '
                    'e psicologi psicoterapeuti.'
                ),
                base_pop=10_000,
                fascia_label='residenti 1–17 anni',
                uo_patterns='NEUROPSICH | NPIA',
                profili_tassi=NPIA_PROFILI,
            )
            + self._tabella_popolazione_distretto(
                fascia_key='range_1_17',
                fascia_label='residenti 1–17 anni',
                base_pop=10_000,
                profili_tassi=NPIA_PROFILI,
            )
        )
        items.append(KeepTogether(block))
        return items

    # ------------------------------------------------------------------
    # 6d. Salute in Carcere
    # ------------------------------------------------------------------

    def _sezione_territoriale_carcere(self) -> list:
        items: list = []
        CARC_PROFILI = [
            ('Medici SM+SerD', 2.0,
             'DIRIGENTE MEDICO (1 SM + 1 SerD)'),
            ('Psicologi SM+SerD', 2.0,
             'DIRIGENTE PSICOLOGO (1 SM + 1 SerD)'),
            ('Infermieri', 1.0,
             'INFERMIERE (SM: 1 prof.san. / 350 det.)'),
            ('Assistenti Sociali', 1.0,
             'ASSISTENTE SOCIALE (SerD: 1 / 350 det.)'),
        ]

        # Costruiamo la tabella detenuti
        hdrs_d = ['Istituto', 'N. Detenuti',
                  'Medici att. (2/350)',
                  'Psicologi att. (2/350)',
                  'Infermieri att. (1/350)',
                  'Ass. Sociali att. (1/350)']
        col_w_d = [
            self.avail_w * 0.22,
            self.avail_w * 0.12,
            self.avail_w * 0.165,
            self.avail_w * 0.165,
            self.avail_w * 0.165,
            self.avail_w * 0.165,
        ]
        rows_d = [[_ph(h) for h in hdrs_d]]
        tot_det = 0
        for istituto, ndet in DETENUTI_PER_ISTITUTO.items():
            tot_det += ndet
            rows_d.append([
                _p(istituto, ST_TD_BOLD),
                _p(str(ndet), ST_TD_C),
                _p(f'{ndet * 2 / 350:.2f}', ST_TD_C),
                _p(f'{ndet * 2 / 350:.2f}', ST_TD_C),
                _p(f'{ndet * 1 / 350:.2f}', ST_TD_C),
                _p(f'{ndet * 1 / 350:.2f}', ST_TD_C),
            ])
        rows_d.append([
            _p('TOTALE ASREM', ST_TD_BOLD),
            _p(str(tot_det), ST_TD_C),
            _p(f'{tot_det * 2 / 350:.2f}', ST_TD_BOLD),
            _p(f'{tot_det * 2 / 350:.2f}', ST_TD_BOLD),
            _p(f'{tot_det * 1 / 350:.2f}', ST_TD_BOLD),
            _p(f'{tot_det * 1 / 350:.2f}', ST_TD_BOLD),
        ])
        n = len(rows_d)
        ts_cmds = _base_ts(n, len(hdrs_d))
        ts_cmds.append(('BACKGROUND', (0, n - 1), (-1, n - 1), CLR_ROW_TOT))
        ts_cmds.append(('FONTNAME', (0, n - 1), (-1, n - 1), 'Helvetica-Bold'))
        td = Table(rows_d, colWidths=col_w_d, repeatRows=1)
        td.setStyle(TableStyle(ts_cmds))

        block = (
            self._tabella_indicatori_territoriale(
                titolo='6d. Area Salute in Carcere',
                fonte='Standard minimi AGENAS – indicatori_agenas_carcere.xml',
                descrizione=(
                    'Tasso ogni 350 detenuti (dato per istituto penitenziario). '
                    'Standard combinato Salute Mentale + SerD penitenziario.'
                ),
                base_pop=350,
                fascia_label='detenuti per istituto',
                uo_patterns='DETENUTI | CARCER | PENITEN',
                profili_tassi=CARC_PROFILI,
            )
            + [
                Paragraph('Detenuti per istituto penitenziario – base calcolo',
                          ST_NOTE),
                td,
                Spacer(1, 5 * mm),
            ]
        )
        items.append(KeepTogether(block))
        return items

    # ------------------------------------------------------------------
    # Sezione 7 – Dati di ingresso
    # ------------------------------------------------------------------

    def _sezione_dati_ingresso(self) -> list:
        items: list = []
        items.append(Paragraph(
            '7. Dati di ingresso utilizzati', ST_SECTION))
        items.append(Paragraph(
            'I dati seguenti sono configurati in src/config.py e '
            'determinano la classificazione dei presidi e il calcolo '
            'degli attesi territoriali.', ST_BODY))
        items.append(Spacer(1, 3 * mm))

        # 7a – Livello presidio
        items.append(Paragraph(
            '7a. Livello dei presidi ospedalieri (DM 70/2015)', ST_SUBSECT))
        PRES_LABELS = {
            'OSPEDALE_DI_BASE': 'Ospedale di Base',
            'PRESIDIO_I_LIVELLO': 'I Livello',
            'PRESIDIO_II_LIVELLO': 'II Livello',
        }
        rows_p = [[_ph('Presidio'), _ph('Livello')]]
        for pres, liv in LIVELLO_PRESIDIO.items():
            rows_p.append([_p(pres), _p(PRES_LABELS.get(liv, liv), ST_TD_C)])
        col_w_p = [self.avail_w * 0.55, self.avail_w * 0.25]
        tp = Table(rows_p, colWidths=col_w_p, repeatRows=1)
        tp.setStyle(TableStyle(_base_ts(len(rows_p), 2)))
        items.append(tp)
        items.append(Spacer(1, 3 * mm))

        # 7b – Parti per presidio
        items.append(Paragraph(
            '7b. Parti annui per presidio (Tab. 11 AGENAS)', ST_SUBSECT))
        rows_pt = [[_ph('Presidio'), _ph('N. Parti annui'), _ph('Fascia AGENAS')]]
        for pres, npar in PARTI_PER_PRESIDIO.items():
            if npar < 500:
                fascia = '< 500 (sotto soglia)'
            elif npar <= 1500:
                fascia = '500–1.500'
            elif npar <= 2000:
                fascia = '1.500–2.000'
            else:
                fascia = '> 2.000'
            rows_pt.append([
                _p(pres),
                _p(str(npar), ST_TD_C),
                _p(fascia, ST_TD_C),
            ])
        col_w_pt = [self.avail_w * 0.45, self.avail_w * 0.15, self.avail_w * 0.20]
        tpt = Table(rows_pt, colWidths=col_w_pt, repeatRows=1)
        tpt.setStyle(TableStyle(_base_ts(len(rows_pt), 3)))
        items.append(tpt)
        items.append(Spacer(1, 3 * mm))

        # 7c – Popolazione per area distrettuale (riepilogo completo)
        items.append(Paragraph(
            '7c. Popolazione per area distrettuale (ISTAT)', ST_SUBSECT))
        POP_HDRS = [
            'Area', 'Totale', '≥ 18 anni', '15–64 anni', '1–17 anni',
        ]
        col_w_pop = [
            self.avail_w * 0.22,
            self.avail_w * 0.14,
            self.avail_w * 0.14,
            self.avail_w * 0.14,
            self.avail_w * 0.14,
        ]
        rows_pop = [[_ph(h) for h in POP_HDRS]]
        tot_tot = tot_18 = tot_1564 = tot_117 = 0
        for area, dati in POPOLAZIONE_AREA.items():
            totale = dati.get('totale', 0)
            if isinstance(totale, float) and totale < 1000:
                totale = int(totale * 1000)
            totale = int(totale)
            g18 = int(dati.get('gte_18', 0))
            r64 = int(dati.get('range_15_64', 0))
            r17 = int(dati.get('range_1_17', 0))
            tot_tot += totale; tot_18 += g18
            tot_1564 += r64; tot_117 += r17
            rows_pop.append([
                _p(area, ST_TD_BOLD),
                _p(f'{totale:,}', ST_TD_C),
                _p(f'{g18:,}', ST_TD_C),
                _p(f'{r64:,}', ST_TD_C),
                _p(f'{r17:,}', ST_TD_C),
            ])
        rows_pop.append([
            _p('TOTALE ASREM', ST_TD_BOLD),
            _p(f'{tot_tot:,}', ST_TD_BOLD),
            _p(f'{tot_18:,}', ST_TD_BOLD),
            _p(f'{tot_1564:,}', ST_TD_BOLD),
            _p(f'{tot_117:,}', ST_TD_BOLD),
        ])
        n = len(rows_pop)
        ts_pop = _base_ts(n, 5)
        ts_pop.append(('BACKGROUND', (0, n - 1), (-1, n - 1), CLR_ROW_TOT))
        ts_pop.append(('FONTNAME', (0, n - 1), (-1, n - 1), 'Helvetica-Bold'))
        tpop = Table(rows_pop, colWidths=col_w_pop, repeatRows=1)
        tpop.setStyle(TableStyle(ts_pop))
        items.append(tpop)
        items.append(Spacer(1, 3 * mm))
        items.append(Paragraph(
            'Le fasce di popolazione sono utilizzate come parametro base '
            'nella formula: Atteso = Popolazione_fascia × Tasso / Base.',
            ST_NOTE))
        items.append(Spacer(1, 5 * mm))

        # 7d – Detenuti per istituto
        items.append(Paragraph(
            '7d. Detenuti per istituto penitenziario', ST_SUBSECT))
        rows_det = [[_ph('Istituto'), _ph('N. Detenuti'),
                     _ph('Unità base (÷ 350)')]]
        for area, ndet in DETENUTI_PER_ISTITUTO.items():
            rows_det.append([
                _p(area),
                _p(str(ndet), ST_TD_C),
                _p(f'{ndet / 350:.2f}', ST_TD_C),
            ])
        col_w_det = [self.avail_w * 0.35, self.avail_w * 0.15,
                     self.avail_w * 0.20]
        tdet = Table(rows_det, colWidths=col_w_det, repeatRows=1)
        tdet.setStyle(TableStyle(_base_ts(len(rows_det), 3)))
        items.append(tdet)
        items.append(Spacer(1, 5 * mm))

        # Fonti normative
        items.append(HRFlowable(width=self.avail_w, thickness=0.5,
                                color=CLR_LINE, spaceAfter=2))
        items.append(Paragraph('Fonti e riferimenti normativi', ST_SUBSECT))
        FONTI = [
            '• DM 70/2015 – Regolamento standard strutture ospedaliere '
            '(classificazione presidi: Base, I, II livello).',
            '• DM 77/2022 – Standard assistenza territoriale (OdC, CdC).',
            '• DPR 1 Novembre 1999 – Progetto Obiettivo Salute Mentale '
            '(6,7 operatori/10.000 ab.).',
            '• AGENAS – Standard di personale Tab. 11, 13–18, 20 e '
            'standard territoriali SerD, NPIA, Carcere.',
            '• Dati ISTAT – Popolazione residente per fasce d\'età.',
            '• Dati aziendali – Parti SDO, posti letto attivi, detenuti, '
            'dotazioni organiche segnalate dai Primari.',
        ]
        for testo in FONTI:
            items.append(Paragraph(testo, ST_BODY))

        return items
