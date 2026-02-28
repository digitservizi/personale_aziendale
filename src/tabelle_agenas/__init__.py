"""
Pacchetto tabelle_agenas – raccoglie tutte le funzioni
di scrittura tabelle AGENAS ospedaliere e territoriali.
"""

from src.tabelle_agenas.materno_infantile import _scrivi_tabella_agenas_materno_infantile
from src.tabelle_agenas.radiologia import _scrivi_tabella_agenas_radiologia
from src.tabelle_agenas.emergenza_urgenza import _scrivi_tabella_agenas_emergenza_urgenza
from src.tabelle_agenas.terapia_intensiva import _scrivi_tabella_agenas_terapia_intensiva
from src.tabelle_agenas.sale_operatorie import _scrivi_tabella_agenas_sale_operatorie
from src.tabelle_agenas.area_ti_bo import _scrivi_tabella_agenas_area_ti_bo
from src.tabelle_agenas.anatomia_patologica import _scrivi_tabella_agenas_anatomia_patologica
from src.tabelle_agenas.laboratorio import _scrivi_tabella_agenas_laboratorio
from src.tabelle_agenas.tecnici_laboratorio import _scrivi_tabella_agenas_tecnici_laboratorio
from src.tabelle_agenas.medicina_legale import _scrivi_tabella_agenas_medicina_legale
from src.tabelle_agenas.trasfusionale import _scrivi_tabella_agenas_trasfusionale, _scrivi_tabella_fabbisogno_uoc_trasfusionale
from src.tabelle_agenas.territoriali import _scrivi_tabella_agenas_territoriale

__all__ = [
    "_scrivi_tabella_agenas_materno_infantile",
    "_scrivi_tabella_agenas_radiologia",
    "_scrivi_tabella_agenas_emergenza_urgenza",
    "_scrivi_tabella_agenas_terapia_intensiva",
    "_scrivi_tabella_agenas_sale_operatorie",
    "_scrivi_tabella_agenas_area_ti_bo",
    "_scrivi_tabella_agenas_anatomia_patologica",
    "_scrivi_tabella_agenas_laboratorio",
    "_scrivi_tabella_agenas_tecnici_laboratorio",
    "_scrivi_tabella_agenas_medicina_legale",
    "_scrivi_tabella_agenas_trasfusionale",
    "_scrivi_tabella_fabbisogno_uoc_trasfusionale",
    "_scrivi_tabella_agenas_territoriale",
]
