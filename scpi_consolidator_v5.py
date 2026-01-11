#!/usr/bin/env python3
"""
SCPI Consolidator V5 - Extraction complète avec CAPEX détaillé par actif
=========================================================================
Version 5.0 avec 8 onglets :
- Dashboard : vue synthétique executive
- Indicateurs : 90+ indicateurs réglementaires
- Patrimoine : actifs immobiliers détenus en direct
- SCI : participations dans les SCI sous-jacentes
- Comptes Courants : avances en comptes courants d'associés
- CAPEX Détail : liste détaillée des travaux par actif
- CAPEX Synthèse : consolidation par SCPI
- Contrôles : tests de cohérence et validation

Conformité : AMF art. 422-227, ASPIM méthodologie 01/01/2022, ANC 2016-03
"""

import re
import os
import sys
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Dict, Tuple, Any
from datetime import datetime
from pathlib import Path
from decimal import Decimal, InvalidOperation
import warnings

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule, CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

warnings.filterwarnings('ignore')


# =============================================================================
# DATACLASSES - Structures de données
# =============================================================================

@dataclass
class IndicateursSCPI:
    """90+ indicateurs réglementaires SCPI"""
    source_pdf: str = ""
    date_extraction: str = ""
    qualite_extraction: str = ""
    nb_indicateurs_extraits: int = 0
    
    # Identification
    nom_scpi: Optional[str] = None
    societe_gestion: Optional[str] = None
    type_scpi: Optional[str] = None
    visa_amf: Optional[str] = None
    date_creation: Optional[str] = None
    
    # Capital et Parts
    capital_nominal: Optional[float] = None
    capital_statutaire_max: Optional[float] = None
    nombre_parts: Optional[int] = None
    nombre_associes: Optional[int] = None
    capitalisation: Optional[float] = None
    valeur_nominale: Optional[float] = None
    
    # Prix et Valeurs
    prix_souscription: Optional[float] = None
    prix_retrait: Optional[float] = None
    prime_emission: Optional[float] = None
    valeur_realisation: Optional[float] = None
    valeur_realisation_par_part: Optional[float] = None
    valeur_reconstitution: Optional[float] = None
    valeur_reconstitution_par_part: Optional[float] = None
    valeur_venale_patrimoine: Optional[float] = None
    valeur_ifi_resident: Optional[float] = None
    valeur_ifi_non_resident: Optional[float] = None
    ecart_prix_reconstitution: Optional[float] = None
    
    # Performance ASPIM
    taux_distribution: Optional[float] = None
    dividende_brut: Optional[float] = None
    dividende_exceptionnel: Optional[float] = None
    report_a_nouveau: Optional[float] = None
    report_a_nouveau_par_part: Optional[float] = None
    tri_5_ans: Optional[float] = None
    tri_10_ans: Optional[float] = None
    tri_15_ans: Optional[float] = None
    tri_20_ans: Optional[float] = None
    rendement_global_immobilier: Optional[float] = None
    
    # Occupation
    tof_annuel: Optional[float] = None
    tof_t1: Optional[float] = None
    tof_t2: Optional[float] = None
    tof_t3: Optional[float] = None
    tof_t4: Optional[float] = None
    top_annuel: Optional[float] = None
    surface_vacante: Optional[float] = None
    surface_vacante_pct: Optional[float] = None
    
    # Patrimoine
    nombre_actifs: Optional[int] = None
    nombre_actifs_direct: Optional[int] = None
    nombre_actifs_indirect: Optional[int] = None
    surface_totale: Optional[float] = None
    prix_moyen_m2: Optional[float] = None
    
    # Répartition typologique
    pct_bureaux: Optional[float] = None
    pct_commerces: Optional[float] = None
    pct_logistique: Optional[float] = None
    pct_activites: Optional[float] = None
    pct_sante: Optional[float] = None
    pct_residentiel: Optional[float] = None
    pct_hotellerie: Optional[float] = None
    pct_enseignement: Optional[float] = None
    pct_autres: Optional[float] = None
    
    # Répartition géographique
    pct_paris: Optional[float] = None
    pct_idf_hors_paris: Optional[float] = None
    pct_regions: Optional[float] = None
    pct_etranger: Optional[float] = None
    pct_france: Optional[float] = None
    
    # Acquisitions / Cessions
    nb_acquisitions: Optional[int] = None
    montant_acquisitions: Optional[float] = None
    nb_cessions: Optional[int] = None
    montant_cessions: Optional[float] = None
    plus_value_cessions: Optional[float] = None
    moins_value_cessions: Optional[float] = None
    
    # Collecte
    collecte_brute: Optional[float] = None
    collecte_nette: Optional[float] = None
    parts_en_attente_retrait: Optional[int] = None
    montant_attente_retrait: Optional[float] = None
    delai_jouissance: Optional[str] = None
    
    # Endettement
    ratio_endettement_ltv: Optional[float] = None
    emprunts_bancaires: Optional[float] = None
    dettes_financieres: Optional[float] = None
    tresorerie: Optional[float] = None
    endettement_net: Optional[float] = None
    
    # Résultats comptables
    loyers_quittances: Optional[float] = None
    produits_locatifs: Optional[float] = None
    produits_financiers: Optional[float] = None
    charges_immobilieres: Optional[float] = None
    resultat_immobilier: Optional[float] = None
    resultat_net: Optional[float] = None
    resultat_par_part: Optional[float] = None
    benefice_comptable: Optional[float] = None
    revenus_fonciers_par_part: Optional[float] = None
    revenus_financiers_par_part: Optional[float] = None
    
    # Frais
    commission_souscription_pct: Optional[float] = None
    commission_gestion_pct: Optional[float] = None
    commission_arbitrage_pct: Optional[float] = None
    commission_souscription_montant: Optional[float] = None
    commission_gestion_montant: Optional[float] = None
    provision_gros_entretien: Optional[float] = None
    
    # CAPEX / Travaux
    total_travaux: Optional[float] = None
    travaux_amelioration: Optional[float] = None
    travaux_gros_entretien: Optional[float] = None
    travaux_renovation: Optional[float] = None
    travaux_mise_conformite: Optional[float] = None
    travaux_environnementaux: Optional[float] = None
    capex_par_m2: Optional[float] = None
    ratio_capex_valeur_venale: Optional[float] = None


@dataclass
class ActifImmobilier:
    """Actif immobilier détenu en direct"""
    scpi_source: str = ""
    id_actif: str = ""
    nom_actif: Optional[str] = None
    adresse: Optional[str] = None
    ville: Optional[str] = None
    code_postal: Optional[str] = None
    pays: str = "France"
    region: Optional[str] = None
    type_actif: Optional[str] = None
    date_acquisition: Optional[str] = None
    surface_m2: Optional[float] = None
    nb_lots: Optional[int] = None
    quote_part_pct: float = 100.0
    valeur_venale: Optional[float] = None
    prix_acquisition: Optional[float] = None
    loyer_annuel: Optional[float] = None
    taux_occupation: Optional[float] = None
    nb_locataires: Optional[int] = None
    date_fin_bail_principale: Optional[str] = None
    # CAPEX détaillés
    travaux_total: float = 0.0
    travaux_renovation: float = 0.0
    travaux_gros_entretien: float = 0.0
    travaux_mise_conformite: float = 0.0
    travaux_environnementaux: float = 0.0
    travaux_amelioration: float = 0.0
    travaux_autres: float = 0.0
    annee_travaux: Optional[int] = None
    description_travaux: Optional[str] = None
    statut_travaux: Optional[str] = None  # Réalisé, En cours, Planifié


@dataclass
class SCISousJacente:
    """SCI ou OPCI détenue par la SCPI"""
    scpi_source: str = ""
    nom_sci: Optional[str] = None
    type_structure: Optional[str] = None
    affectation: Optional[str] = None
    date_acquisition: Optional[str] = None
    date_constitution: Optional[str] = None
    quote_part_detention: Optional[float] = None
    nombre_immeubles: Optional[int] = None
    surface_m2: Optional[float] = None
    valeur_venale_hd: Optional[float] = None
    droits: Optional[float] = None
    prix_acquisition: Optional[float] = None
    travaux: Optional[float] = None
    valeur_nette_comptable: Optional[float] = None
    valeur_estimee: Optional[float] = None
    ecart_acquisition: Optional[float] = None
    immeubles: Optional[float] = None
    dettes: Optional[float] = None


@dataclass
class CompteCourantAssocie:
    """Avance en compte courant d'associé"""
    scpi_source: str = ""
    nom_entite: Optional[str] = None
    type_entite: Optional[str] = None
    compte_courant_rattache: Optional[float] = None
    autres_creances: Optional[float] = None
    dettes_fournisseurs: Optional[float] = None
    commissions: Optional[float] = None
    dividendes_participations: Optional[float] = None
    produits_comptes_courants: Optional[float] = None


@dataclass
class CAPEXDetail:
    """Détail CAPEX par actif"""
    scpi_source: str = ""
    id_actif: str = ""
    nom_actif: Optional[str] = None
    type_actif: Optional[str] = None
    ville: Optional[str] = None
    surface_m2: Optional[float] = None
    valeur_venale: Optional[float] = None
    # Détail travaux
    travaux_total: float = 0.0
    travaux_renovation: float = 0.0
    travaux_gros_entretien: float = 0.0
    travaux_mise_conformite: float = 0.0
    travaux_environnementaux: float = 0.0
    travaux_amelioration: float = 0.0
    travaux_autres: float = 0.0
    # Ratios
    capex_par_m2: Optional[float] = None
    capex_sur_vv_pct: Optional[float] = None
    # Métadonnées
    annee_travaux: Optional[int] = None
    description_travaux: Optional[str] = None
    statut_travaux: Optional[str] = None
    source_donnee: str = "Actif Direct"


@dataclass
class CAPEXSynthese:
    """Synthèse CAPEX par SCPI"""
    scpi_source: str = ""
    # Volumes
    total_capex: float = 0.0
    capex_actifs_directs: float = 0.0
    capex_sci: float = 0.0
    # Comptages
    nb_actifs_total: int = 0
    nb_actifs_avec_capex: int = 0
    pct_actifs_avec_capex: float = 0.0
    nb_sci_total: int = 0
    nb_sci_avec_capex: int = 0
    # Moyennes
    capex_moyen_par_actif: float = 0.0
    capex_median_par_actif: float = 0.0
    capex_min: float = 0.0
    capex_max: float = 0.0
    # Ratios globaux
    capex_sur_vv_pct: Optional[float] = None
    capex_par_m2: Optional[float] = None
    capex_sur_loyers_pct: Optional[float] = None
    # Ventilation par type
    pct_renovation: float = 0.0
    pct_gros_entretien: float = 0.0
    pct_mise_conformite: float = 0.0
    pct_environnemental: float = 0.0
    pct_amelioration: float = 0.0
    pct_autres: float = 0.0
    # Surface concernée
    surface_totale_concernee: float = 0.0


@dataclass
class TestCoherence:
    """Résultat d'un test de cohérence"""
    scpi_source: str = ""
    categorie: str = ""
    test_id: str = ""
    description: str = ""
    valeur_testee: Optional[str] = None
    valeur_reference: Optional[str] = None
    ecart_pct: Optional[float] = None
    statut: str = ""
    commentaire: str = ""
    severite: str = "Info"  # Info, Warning, Error, Critical


# =============================================================================
# CLASSE PRINCIPALE - Consolidateur V5
# =============================================================================

class SCPIConsolidatorV5:
    """Consolidateur SCPI V5 avec CAPEX détaillé par actif"""
    
    def __init__(self, output_dir: str = "."):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.results = {
            'indicateurs': [],
            'actifs': [],
            'sci': [],
            'comptes_courants': [],
            'capex_detail': [],
            'capex_synthese': [],
            'controles': [],
            'erreurs': []
        }
        
        # Seuils pour tests de cohérence
        self.seuils = {
            'ecart_valeur_venale_pct': 5.0,
            'capex_max_pct_valeur': 20.0,
            'capex_normal_min_m2': 30.0,
            'capex_normal_max_m2': 250.0,
            'tof_min': 50.0,
            'tof_max': 100.0,
            'td_min': 0.0,
            'td_max': 15.0,
            'ltv_max': 50.0,
            'somme_geo_tolerance': 2.0,
            'somme_typo_tolerance': 2.0,
        }
        
        # Styles Excel
        self._init_styles()
    
    def _init_styles(self):
        """Initialise les styles Excel"""
        self.styles = {
            'header': {
                'fill': PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"),
                'font': Font(color="FFFFFF", bold=True, size=11),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
            },
            'header_capex': {
                'fill': PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid"),
                'font': Font(color="FFFFFF", bold=True, size=11),
                'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
            },
            'subheader': {
                'fill': PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid"),
                'font': Font(color="FFFFFF", bold=True, size=10),
            },
            'ok': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
            'warning': PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
            'error': PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
            'critical': PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid"),
            'total': {
                'fill': PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
                'font': Font(bold=True, size=11)
            },
            'money': '#,##0 €',
            'money_k': '#,##0 K€',
            'pct': '0.00%',
            'pct1': '0.0%',
            'number': '#,##0',
            'border': Border(
                left=Side(style='thin', color='B4B4B4'),
                right=Side(style='thin', color='B4B4B4'),
                top=Side(style='thin', color='B4B4B4'),
                bottom=Side(style='thin', color='B4B4B4')
            )
        }
    
    # =========================================================================
    # CONSOLIDATION
    # =========================================================================
    
    def consolidate(self):
        """Effectue toutes les consolidations"""
        print("🔧 Consolidation des CAPEX détaillés...")
        self._build_capex_detail()
        
        print("📊 Consolidation des synthèses CAPEX...")
        self._build_capex_synthese()
        
        print("✅ Exécution des tests de cohérence...")
        self._run_all_tests()
        
        return self.results
    
    def _build_capex_detail(self):
        """Construit le détail CAPEX par actif"""
        # Depuis les actifs directs
        for actif in self.results['actifs']:
            if actif.travaux_total > 0 or any([
                actif.travaux_renovation, actif.travaux_gros_entretien,
                actif.travaux_mise_conformite, actif.travaux_environnementaux,
                actif.travaux_amelioration, actif.travaux_autres
            ]):
                total = actif.travaux_total or sum(filter(None, [
                    actif.travaux_renovation, actif.travaux_gros_entretien,
                    actif.travaux_mise_conformite, actif.travaux_environnementaux,
                    actif.travaux_amelioration, actif.travaux_autres
                ]))
                
                detail = CAPEXDetail(
                    scpi_source=actif.scpi_source,
                    id_actif=actif.id_actif,
                    nom_actif=actif.nom_actif or actif.adresse,
                    type_actif=actif.type_actif,
                    ville=actif.ville,
                    surface_m2=actif.surface_m2,
                    valeur_venale=actif.valeur_venale,
                    travaux_total=total,
                    travaux_renovation=actif.travaux_renovation,
                    travaux_gros_entretien=actif.travaux_gros_entretien,
                    travaux_mise_conformite=actif.travaux_mise_conformite,
                    travaux_environnementaux=actif.travaux_environnementaux,
                    travaux_amelioration=actif.travaux_amelioration,
                    travaux_autres=actif.travaux_autres,
                    annee_travaux=actif.annee_travaux,
                    description_travaux=actif.description_travaux,
                    statut_travaux=actif.statut_travaux,
                    source_donnee="Actif Direct"
                )
                
                # Calcul des ratios
                if actif.surface_m2 and actif.surface_m2 > 0:
                    detail.capex_par_m2 = total / actif.surface_m2
                if actif.valeur_venale and actif.valeur_venale > 0:
                    detail.capex_sur_vv_pct = (total / actif.valeur_venale) * 100
                
                self.results['capex_detail'].append(detail)
        
        # Depuis les SCI
        for sci in self.results['sci']:
            if sci.travaux and sci.travaux > 0:
                detail = CAPEXDetail(
                    scpi_source=sci.scpi_source,
                    id_actif=f"SCI-{sci.nom_sci}",
                    nom_actif=sci.nom_sci,
                    type_actif="SCI",
                    surface_m2=sci.surface_m2,
                    valeur_venale=sci.valeur_venale_hd,
                    travaux_total=sci.travaux,
                    source_donnee="SCI"
                )
                
                if sci.surface_m2 and sci.surface_m2 > 0:
                    detail.capex_par_m2 = sci.travaux / sci.surface_m2
                if sci.valeur_venale_hd and sci.valeur_venale_hd > 0:
                    detail.capex_sur_vv_pct = (sci.travaux / sci.valeur_venale_hd) * 100
                
                self.results['capex_detail'].append(detail)
    
    def _build_capex_synthese(self):
        """Construit la synthèse CAPEX par SCPI"""
        scpi_list = set(ind.nom_scpi or ind.source_pdf for ind in self.results['indicateurs'])
        
        for scpi in scpi_list:
            ind = next((i for i in self.results['indicateurs'] if (i.nom_scpi or i.source_pdf) == scpi), None)
            if not ind:
                continue
            
            synthese = CAPEXSynthese(scpi_source=scpi)
            
            # Filtrer les données pour cette SCPI
            actifs_scpi = [a for a in self.results['actifs'] if a.scpi_source == scpi]
            sci_scpi = [s for s in self.results['sci'] if s.scpi_source == scpi]
            capex_scpi = [c for c in self.results['capex_detail'] if c.scpi_source == scpi]
            
            # Volumes
            capex_directs = [c for c in capex_scpi if c.source_donnee == "Actif Direct"]
            capex_sci = [c for c in capex_scpi if c.source_donnee == "SCI"]
            
            synthese.capex_actifs_directs = sum(c.travaux_total for c in capex_directs)
            synthese.capex_sci = sum(c.travaux_total for c in capex_sci)
            synthese.total_capex = synthese.capex_actifs_directs + synthese.capex_sci
            
            # Si indicateur total_travaux disponible et supérieur, l'utiliser
            if ind.total_travaux and ind.total_travaux > synthese.total_capex:
                synthese.total_capex = ind.total_travaux
            
            # Comptages
            synthese.nb_actifs_total = len(actifs_scpi)
            synthese.nb_actifs_avec_capex = len(capex_directs)
            if synthese.nb_actifs_total > 0:
                synthese.pct_actifs_avec_capex = (synthese.nb_actifs_avec_capex / synthese.nb_actifs_total) * 100
            
            synthese.nb_sci_total = len(sci_scpi)
            synthese.nb_sci_avec_capex = len(capex_sci)
            
            # Statistiques sur les montants
            all_capex_amounts = [c.travaux_total for c in capex_scpi if c.travaux_total > 0]
            if all_capex_amounts:
                synthese.capex_moyen_par_actif = sum(all_capex_amounts) / len(all_capex_amounts)
                sorted_amounts = sorted(all_capex_amounts)
                mid = len(sorted_amounts) // 2
                synthese.capex_median_par_actif = sorted_amounts[mid] if len(sorted_amounts) % 2 else (sorted_amounts[mid-1] + sorted_amounts[mid]) / 2
                synthese.capex_min = min(all_capex_amounts)
                synthese.capex_max = max(all_capex_amounts)
            
            # Ratios globaux
            if ind.valeur_venale_patrimoine and ind.valeur_venale_patrimoine > 0:
                synthese.capex_sur_vv_pct = (synthese.total_capex / ind.valeur_venale_patrimoine) * 100
            if ind.surface_totale and ind.surface_totale > 0:
                synthese.capex_par_m2 = synthese.total_capex / ind.surface_totale
            if ind.loyers_quittances and ind.loyers_quittances > 0:
                synthese.capex_sur_loyers_pct = (synthese.total_capex / ind.loyers_quittances) * 100
            
            # Ventilation par type de travaux
            if synthese.total_capex > 0:
                total_renovation = sum(c.travaux_renovation or 0 for c in capex_scpi)
                total_entretien = sum(c.travaux_gros_entretien or 0 for c in capex_scpi)
                total_conformite = sum(c.travaux_mise_conformite or 0 for c in capex_scpi)
                total_env = sum(c.travaux_environnementaux or 0 for c in capex_scpi)
                total_amelio = sum(c.travaux_amelioration or 0 for c in capex_scpi)
                total_autres = sum(c.travaux_autres or 0 for c in capex_scpi)
                
                synthese.pct_renovation = (total_renovation / synthese.total_capex) * 100
                synthese.pct_gros_entretien = (total_entretien / synthese.total_capex) * 100
                synthese.pct_mise_conformite = (total_conformite / synthese.total_capex) * 100
                synthese.pct_environnemental = (total_env / synthese.total_capex) * 100
                synthese.pct_amelioration = (total_amelio / synthese.total_capex) * 100
                synthese.pct_autres = (total_autres / synthese.total_capex) * 100
            
            # Surface concernée
            synthese.surface_totale_concernee = sum(c.surface_m2 or 0 for c in capex_scpi)
            
            self.results['capex_synthese'].append(synthese)
    
    # =========================================================================
    # TESTS DE COHÉRENCE
    # =========================================================================
    
    def _run_all_tests(self):
        """Exécute tous les tests de cohérence"""
        for ind in self.results['indicateurs']:
            scpi = ind.nom_scpi or ind.source_pdf
            
            self._test_capex(ind, scpi)
            self._test_valorisation(ind, scpi)
            self._test_occupation(ind, scpi)
            self._test_endettement(ind, scpi)
            self._test_repartition(ind, scpi)
            self._test_patrimoine(ind, scpi)
    
    def _test_capex(self, ind: IndicateursSCPI, scpi: str):
        """Tests sur les CAPEX"""
        synthese = next((s for s in self.results['capex_synthese'] if s.scpi_source == scpi), None)
        if not synthese:
            return
        
        # CAP-001: Ratio CAPEX / Valeur Vénale
        if synthese.capex_sur_vv_pct is not None:
            statut = "OK" if synthese.capex_sur_vv_pct < self.seuils['capex_max_pct_valeur'] else "ALERTE"
            severite = "Info" if statut == "OK" else "Warning"
            commentaire = ""
            if synthese.capex_sur_vv_pct >= self.seuils['capex_max_pct_valeur']:
                commentaire = "Ratio CAPEX élevé - vérifier la stratégie d'investissement"
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="CAPEX", test_id="CAP-001",
                description="Ratio CAPEX / Valeur Vénale",
                valeur_testee=f"{synthese.capex_sur_vv_pct:.2f}%",
                valeur_reference=f"< {self.seuils['capex_max_pct_valeur']}%",
                statut=statut, severite=severite, commentaire=commentaire
            ))
        
        # CAP-002: CAPEX par m²
        if synthese.capex_par_m2 is not None and synthese.capex_par_m2 > 0:
            in_range = self.seuils['capex_normal_min_m2'] <= synthese.capex_par_m2 <= self.seuils['capex_normal_max_m2']
            statut = "OK" if in_range else "ALERTE"
            commentaire = ""
            if synthese.capex_par_m2 < self.seuils['capex_normal_min_m2']:
                commentaire = "CAPEX faible - risque de sous-investissement"
            elif synthese.capex_par_m2 > self.seuils['capex_normal_max_m2']:
                commentaire = "CAPEX élevé - travaux majeurs en cours"
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="CAPEX", test_id="CAP-002",
                description="CAPEX par m² (benchmark 30-250 €/m²)",
                valeur_testee=f"{synthese.capex_par_m2:.2f} €/m²",
                valeur_reference=f"{self.seuils['capex_normal_min_m2']}-{self.seuils['capex_normal_max_m2']} €/m²",
                statut=statut, severite="Info" if statut == "OK" else "Warning",
                commentaire=commentaire
            ))
        
        # CAP-003: Couverture CAPEX
        if synthese.nb_actifs_total > 0:
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="CAPEX", test_id="CAP-003",
                description="Taux d'actifs avec CAPEX",
                valeur_testee=f"{synthese.pct_actifs_avec_capex:.1f}%",
                valeur_reference="Information",
                statut="INFO", severite="Info"
            ))
    
    def _test_valorisation(self, ind: IndicateursSCPI, scpi: str):
        """Tests sur la valorisation"""
        if ind.prix_souscription and ind.valeur_reconstitution_par_part:
            ecart = ((ind.prix_souscription - ind.valeur_reconstitution_par_part) 
                     / ind.valeur_reconstitution_par_part * 100)
            statut = "OK" if abs(ecart) < self.seuils['ecart_valeur_venale_pct'] else "ALERTE"
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="VALORISATION", test_id="VAL-001",
                description="Écart prix souscription vs valeur reconstitution",
                valeur_testee=f"{ind.prix_souscription:,.2f} €",
                valeur_reference=f"{ind.valeur_reconstitution_par_part:,.2f} €",
                ecart_pct=ecart, statut=statut,
                severite="Info" if statut == "OK" else "Warning"
            ))
    
    def _test_occupation(self, ind: IndicateursSCPI, scpi: str):
        """Tests sur l'occupation"""
        if ind.tof_annuel is not None:
            if ind.tof_annuel < 80:
                statut, severite = "ALERTE", "Warning"
                commentaire = "TOF inférieur à 80% - vigilance"
            elif ind.tof_annuel < self.seuils['tof_min']:
                statut, severite = "ERREUR", "Error"
                commentaire = "TOF anormalement bas"
            else:
                statut, severite = "OK", "Info"
                commentaire = ""
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="OCCUPATION", test_id="OCC-001",
                description="TOF dans plage acceptable",
                valeur_testee=f"{ind.tof_annuel:.2f}%",
                valeur_reference=f"≥ 80%",
                statut=statut, severite=severite, commentaire=commentaire
            ))
    
    def _test_endettement(self, ind: IndicateursSCPI, scpi: str):
        """Tests sur l'endettement"""
        if ind.ratio_endettement_ltv is not None:
            if ind.ratio_endettement_ltv >= self.seuils['ltv_max']:
                statut, severite = "ALERTE", "Warning"
                commentaire = "LTV à la limite réglementaire"
            elif ind.ratio_endettement_ltv >= 40:
                statut, severite = "OK", "Info"
                commentaire = "LTV élevé - surveiller"
            else:
                statut, severite = "OK", "Info"
                commentaire = ""
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="ENDETTEMENT", test_id="END-001",
                description="LTV vs seuil réglementaire (50%)",
                valeur_testee=f"{ind.ratio_endettement_ltv:.2f}%",
                valeur_reference=f"< {self.seuils['ltv_max']}%",
                statut=statut, severite=severite, commentaire=commentaire
            ))
    
    def _test_repartition(self, ind: IndicateursSCPI, scpi: str):
        """Tests sur les répartitions"""
        # Géographique
        geo = [ind.pct_paris, ind.pct_idf_hors_paris, ind.pct_regions, ind.pct_etranger]
        geo_valides = [g for g in geo if g is not None]
        if len(geo_valides) >= 2:
            total = sum(geo_valides)
            ecart = abs(100 - total)
            statut = "OK" if ecart < self.seuils['somme_geo_tolerance'] else "ALERTE"
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="REPARTITION", test_id="REP-001",
                description="Somme répartition géographique = 100%",
                valeur_testee=f"{total:.1f}%", valeur_reference="100%",
                ecart_pct=ecart, statut=statut,
                severite="Info" if statut == "OK" else "Warning"
            ))
        
        # Typologique
        typo = [ind.pct_bureaux, ind.pct_commerces, ind.pct_logistique, 
                ind.pct_activites, ind.pct_sante, ind.pct_residentiel,
                ind.pct_hotellerie, ind.pct_enseignement, ind.pct_autres]
        typo_valides = [t for t in typo if t is not None]
        if len(typo_valides) >= 2:
            total = sum(typo_valides)
            ecart = abs(100 - total)
            statut = "OK" if ecart < self.seuils['somme_typo_tolerance'] else "ALERTE"
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="REPARTITION", test_id="REP-002",
                description="Somme répartition typologique = 100%",
                valeur_testee=f"{total:.1f}%", valeur_reference="100%",
                ecart_pct=ecart, statut=statut,
                severite="Info" if statut == "OK" else "Warning"
            ))
    
    def _test_patrimoine(self, ind: IndicateursSCPI, scpi: str):
        """Tests sur le patrimoine"""
        actifs_scpi = [a for a in self.results['actifs'] if a.scpi_source == scpi]
        sci_scpi = [s for s in self.results['sci'] if s.scpi_source == scpi]
        
        if ind.nombre_actifs and len(actifs_scpi) > 0:
            nb_total_detail = len(actifs_scpi) + sum(s.nombre_immeubles or 0 for s in sci_scpi)
            ecart_pct = abs(ind.nombre_actifs - nb_total_detail) / ind.nombre_actifs * 100 if ind.nombre_actifs else 0
            
            self.results['controles'].append(TestCoherence(
                scpi_source=scpi, categorie="PATRIMOINE", test_id="PAT-001",
                description="Cohérence nb actifs déclaré vs détail",
                valeur_testee=str(ind.nombre_actifs),
                valeur_reference=f"{nb_total_detail} (détail)",
                ecart_pct=ecart_pct,
                statut="OK" if ecart_pct < 10 else "ALERTE",
                severite="Info" if ecart_pct < 10 else "Warning"
            ))
    
    # =========================================================================
    # EXPORT EXCEL
    # =========================================================================
    
    def export_excel(self, output_path: str = None) -> str:
        """Exporte vers Excel avec 8 onglets formatés"""
        if output_path is None:
            output_path = self.output_dir / "scpi_consolidation_v5.xlsx"
        else:
            output_path = Path(output_path)
        
        wb = Workbook()
        
        # 1. Dashboard
        ws_dash = wb.active
        ws_dash.title = "Dashboard"
        self._create_dashboard(ws_dash)
        
        # 2. Indicateurs
        ws_ind = wb.create_sheet("Indicateurs")
        self._create_indicateurs_sheet(ws_ind)
        
        # 3. Patrimoine
        ws_pat = wb.create_sheet("Patrimoine")
        self._create_patrimoine_sheet(ws_pat)
        
        # 4. SCI
        ws_sci = wb.create_sheet("SCI")
        self._create_sci_sheet(ws_sci)
        
        # 5. Comptes Courants
        ws_cc = wb.create_sheet("Comptes Courants")
        self._create_cc_sheet(ws_cc)
        
        # 6. CAPEX Détail
        ws_capex_det = wb.create_sheet("CAPEX Détail")
        self._create_capex_detail_sheet(ws_capex_det)
        
        # 7. CAPEX Synthèse
        ws_capex_syn = wb.create_sheet("CAPEX Synthèse")
        self._create_capex_synthese_sheet(ws_capex_syn)
        
        # 8. Contrôles
        ws_ctrl = wb.create_sheet("Contrôles")
        self._create_controles_sheet(ws_ctrl)
        
        wb.save(output_path)
        return str(output_path)
    
    def _apply_header_style(self, cell, style_type='header'):
        """Applique le style header à une cellule"""
        style = self.styles[style_type]
        cell.fill = style['fill']
        cell.font = style['font']
        cell.alignment = style.get('alignment', Alignment(horizontal='center', vertical='center'))
        cell.border = self.styles['border']
    
    def _auto_width(self, ws, min_width=8, max_width=50):
        """Ajuste automatiquement la largeur des colonnes"""
        for col_idx in range(1, ws.max_column + 1):
            max_length = 0
            column_letter = get_column_letter(col_idx)
            for row in range(1, ws.max_row + 1):
                try:
                    cell = ws.cell(row=row, column=col_idx)
                    if cell.value:
                        cell_len = len(str(cell.value))
                        if cell_len > max_length:
                            max_length = cell_len
                except:
                    pass
            ws.column_dimensions[column_letter].width = min(max(max_length + 2, min_width), max_width)
    
    def _create_dashboard(self, ws):
        """Crée le dashboard exécutif"""
        # Titre
        ws.merge_cells('A1:M1')
        ws['A1'] = "📊 TABLEAU DE BORD CONSOLIDATION SCPI"
        ws['A1'].font = Font(bold=True, size=16, color="1F4E79")
        ws['A1'].alignment = Alignment(horizontal='center')
        
        ws['A2'] = f"Date d'extraction: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws['A2'].font = Font(italic=True, color="666666")
        
        # KPIs globaux
        row = 4
        ws.merge_cells(f'A{row}:D{row}')
        ws[f'A{row}'] = "📈 KPIs GLOBAUX"
        ws[f'A{row}'].font = Font(bold=True, size=12)
        
        row += 1
        kpis = [
            ("SCPI analysées", len(self.results['indicateurs'])),
            ("Actifs directs", len(self.results['actifs'])),
            ("SCI sous-jacentes", len(self.results['sci'])),
            ("Comptes courants", len(self.results['comptes_courants'])),
        ]
        
        for i, (label, value) in enumerate(kpis):
            ws.cell(row=row, column=1+i*2, value=label).font = Font(bold=True)
            ws.cell(row=row, column=2+i*2, value=value)
        
        # Tableau récapitulatif par SCPI
        row += 3
        ws.merge_cells(f'A{row}:M{row}')
        ws[f'A{row}'] = "🏢 RÉCAPITULATIF PAR SCPI"
        ws[f'A{row}'].font = Font(bold=True, size=12)
        
        row += 1
        headers = ['SCPI', 'Société Gestion', 'Type', 'TD %', 'TOF %', 
                   'Nb Actifs', 'Capitalisation M€', 'VV Patrimoine M€',
                   'CAPEX Total €', 'CAPEX/VV %', 'CAPEX/m² €', 'Qualité', 'Alertes']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            self._apply_header_style(cell)
        
        for ind in self.results['indicateurs']:
            row += 1
            scpi = ind.nom_scpi or ind.source_pdf
            synthese = next((s for s in self.results['capex_synthese'] if s.scpi_source == scpi), None)
            nb_alertes = len([c for c in self.results['controles'] 
                             if c.scpi_source == scpi and c.statut in ['ALERTE', 'ERREUR']])
            
            values = [
                ind.nom_scpi,
                ind.societe_gestion,
                ind.type_scpi,
                ind.taux_distribution,
                ind.tof_annuel,
                ind.nombre_actifs,
                ind.capitalisation / 1_000_000 if ind.capitalisation else None,
                ind.valeur_venale_patrimoine / 1_000_000 if ind.valeur_venale_patrimoine else None,
                synthese.total_capex if synthese else None,
                synthese.capex_sur_vv_pct if synthese else None,
                synthese.capex_par_m2 if synthese else None,
                ind.qualite_extraction,
                nb_alertes
            ]
            
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.border = self.styles['border']
                
                # Formatage conditionnel
                if col == 13 and value and value > 0:
                    cell.fill = self.styles['warning']
                elif col == 4 and value:  # TD
                    cell.number_format = '0.00'
                elif col == 5 and value:  # TOF
                    cell.number_format = '0.00'
                elif col in [7, 8] and value:  # Montants M€
                    cell.number_format = '#,##0.00'
                elif col == 9 and value:  # CAPEX Total
                    cell.number_format = '#,##0'
                elif col in [10, 11] and value:  # Ratios
                    cell.number_format = '0.00'
        
        # Résumé CAPEX
        row += 3
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'] = "🔧 SYNTHÈSE CAPEX GLOBALE"
        ws[f'A{row}'].font = Font(bold=True, size=12)
        
        row += 1
        total_capex = sum(s.total_capex for s in self.results['capex_synthese'])
        total_actifs_capex = sum(s.nb_actifs_avec_capex for s in self.results['capex_synthese'])
        
        capex_stats = [
            ("Total CAPEX consolidé:", f"{total_capex:,.0f} €"),
            ("Actifs avec CAPEX:", total_actifs_capex),
            ("Fiches CAPEX détaillées:", len(self.results['capex_detail'])),
        ]
        
        for label, value in capex_stats:
            ws.cell(row=row, column=1, value=label).font = Font(bold=True)
            ws.cell(row=row, column=2, value=value)
            row += 1
        
        # Résumé contrôles
        row += 2
        ws.merge_cells(f'A{row}:D{row}')
        ws[f'A{row}'] = "✅ RÉSUMÉ CONTRÔLES"
        ws[f'A{row}'].font = Font(bold=True, size=12)
        
        row += 1
        nb_ok = len([c for c in self.results['controles'] if c.statut == "OK"])
        nb_alerte = len([c for c in self.results['controles'] if c.statut == "ALERTE"])
        nb_erreur = len([c for c in self.results['controles'] if c.statut == "ERREUR"])
        nb_info = len([c for c in self.results['controles'] if c.statut == "INFO"])
        
        ws.cell(row=row, column=1, value="✅ OK:")
        ws.cell(row=row, column=2, value=nb_ok)
        ws.cell(row=row, column=2).fill = self.styles['ok']
        
        ws.cell(row=row+1, column=1, value="⚠️ Alertes:")
        ws.cell(row=row+1, column=2, value=nb_alerte)
        ws.cell(row=row+1, column=2).fill = self.styles['warning']
        
        ws.cell(row=row+2, column=1, value="❌ Erreurs:")
        ws.cell(row=row+2, column=2, value=nb_erreur)
        ws.cell(row=row+2, column=2).fill = self.styles['error']
        
        ws.cell(row=row+3, column=1, value="ℹ️ Info:")
        ws.cell(row=row+3, column=2, value=nb_info)
        
        if self.results['controles']:
            taux = nb_ok / len(self.results['controles']) * 100
            ws.cell(row=row+4, column=1, value="📈 Taux conformité:")
            ws.cell(row=row+4, column=2, value=f"{taux:.1f}%")
            ws.cell(row=row+4, column=2).font = Font(bold=True)
        
        self._auto_width(ws)
    
    def _create_indicateurs_sheet(self, ws):
        """Crée l'onglet Indicateurs"""
        if not self.results['indicateurs']:
            ws['A1'] = "Aucun indicateur extrait"
            return
        
        data = [asdict(ind) for ind in self.results['indicateurs']]
        df = pd.DataFrame(data)
        
        for col, header in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col, value=header)
            self._apply_header_style(cell)
        
        for row_idx, row_data in enumerate(df.values, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.styles['border']
        
        self._auto_width(ws)
    
    def _create_patrimoine_sheet(self, ws):
        """Crée l'onglet Patrimoine (actifs)"""
        if not self.results['actifs']:
            ws['A1'] = "Aucun actif extrait"
            return
        
        headers = [
            'SCPI', 'ID Actif', 'Nom/Adresse', 'Ville', 'Région', 'Pays',
            'Type', 'Surface m²', 'Valeur Vénale €', 'Prix Acquisition €',
            'Loyer Annuel €', 'TOC %', 'Nb Locataires',
            'CAPEX Total €', 'CAPEX Réno €', 'CAPEX Entretien €', 
            'CAPEX Conformité €', 'CAPEX Env €', 'Statut Travaux', 'Description Travaux'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            if 'CAPEX' in header:
                self._apply_header_style(cell, 'header_capex')
            else:
                self._apply_header_style(cell)
        
        for row_idx, actif in enumerate(self.results['actifs'], 2):
            values = [
                actif.scpi_source, actif.id_actif, actif.nom_actif or actif.adresse,
                actif.ville, actif.region, actif.pays, actif.type_actif,
                actif.surface_m2, actif.valeur_venale, actif.prix_acquisition,
                actif.loyer_annuel, actif.taux_occupation, actif.nb_locataires,
                actif.travaux_total, actif.travaux_renovation, actif.travaux_gros_entretien,
                actif.travaux_mise_conformite, actif.travaux_environnementaux,
                actif.statut_travaux, actif.description_travaux
            ]
            
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.styles['border']
                
                # Format monétaire
                if col_idx in [9, 10, 11, 14, 15, 16, 17, 18] and value:
                    cell.number_format = '#,##0'
                # Surligner les lignes avec CAPEX
                if actif.travaux_total and actif.travaux_total > 0 and col_idx >= 14:
                    cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        
        self._auto_width(ws)
    
    def _create_sci_sheet(self, ws):
        """Crée l'onglet SCI"""
        if not self.results['sci']:
            ws['A1'] = "Aucune SCI extraite"
            return
        
        headers = [
            'SCPI', 'Nom SCI', 'Type Structure', 'Quote-part %',
            'Nb Immeubles', 'Surface m²', 'Valeur Vénale HD €',
            'Prix Acquisition €', 'TRAVAUX €', 'VNC €', 'Valeur Estimée €',
            'Écart Acq €', 'Dettes €'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            if header == 'TRAVAUX €':
                self._apply_header_style(cell, 'header_capex')
            else:
                self._apply_header_style(cell)
        
        for row_idx, sci in enumerate(self.results['sci'], 2):
            values = [
                sci.scpi_source, sci.nom_sci, sci.type_structure,
                sci.quote_part_detention, sci.nombre_immeubles, sci.surface_m2,
                sci.valeur_venale_hd, sci.prix_acquisition, sci.travaux,
                sci.valeur_nette_comptable, sci.valeur_estimee,
                sci.ecart_acquisition, sci.dettes
            ]
            
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.styles['border']
                
                if col_idx in [6, 7, 8, 9, 10, 11, 12, 13] and value:
                    cell.number_format = '#,##0'
                # Surligner travaux
                if col_idx == 9 and value and value > 0:
                    cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
                    cell.font = Font(bold=True)
        
        self._auto_width(ws)
    
    def _create_cc_sheet(self, ws):
        """Crée l'onglet Comptes Courants"""
        if not self.results['comptes_courants']:
            ws['A1'] = "Aucun compte courant extrait"
            return
        
        data = [asdict(cc) for cc in self.results['comptes_courants']]
        df = pd.DataFrame(data)
        
        for col, header in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col, value=header)
            self._apply_header_style(cell)
        
        for row_idx, row_data in enumerate(df.values, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.styles['border']
        
        self._auto_width(ws)
    
    def _create_capex_detail_sheet(self, ws):
        """Crée l'onglet CAPEX Détail par actif"""
        # Titre
        ws.merge_cells('A1:Q1')
        ws['A1'] = "🔧 DÉTAIL DES CAPEX PAR ACTIF"
        ws['A1'].font = Font(bold=True, size=14, color="2E7D32")
        ws['A1'].alignment = Alignment(horizontal='center')
        
        if not self.results['capex_detail']:
            ws['A3'] = "Aucun CAPEX détaillé extrait"
            return
        
        headers = [
            'SCPI', 'ID Actif', 'Nom Actif', 'Type', 'Ville', 'Source',
            'Surface m²', 'Valeur Vénale €',
            'CAPEX TOTAL €', 'Rénovation €', 'Gros Entretien €', 
            'Conformité €', 'Environnement €', 'Amélioration €', 'Autres €',
            'CAPEX/m² €', 'CAPEX/VV %', 'Statut', 'Description'
        ]
        
        row = 3
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            self._apply_header_style(cell, 'header_capex')
        
        # Tri par SCPI puis par montant décroissant
        sorted_capex = sorted(self.results['capex_detail'], 
                             key=lambda x: (x.scpi_source, -x.travaux_total))
        
        current_scpi = None
        for capex in sorted_capex:
            row += 1
            
            # Ligne de séparation entre SCPI
            if current_scpi and capex.scpi_source != current_scpi:
                row += 1
                ws.merge_cells(f'A{row}:S{row}')
                ws.cell(row=row, column=1).fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                row += 1
            
            current_scpi = capex.scpi_source
            
            values = [
                capex.scpi_source, capex.id_actif, capex.nom_actif, capex.type_actif,
                capex.ville, capex.source_donnee, capex.surface_m2, capex.valeur_venale,
                capex.travaux_total, capex.travaux_renovation, capex.travaux_gros_entretien,
                capex.travaux_mise_conformite, capex.travaux_environnementaux,
                capex.travaux_amelioration, capex.travaux_autres,
                capex.capex_par_m2, capex.capex_sur_vv_pct,
                capex.statut_travaux, capex.description_travaux
            ]
            
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx, value=value)
                cell.border = self.styles['border']
                
                # Formatage
                if col_idx in [7, 8, 9, 10, 11, 12, 13, 14, 15] and value:
                    cell.number_format = '#,##0'
                if col_idx == 16 and value:  # CAPEX/m²
                    cell.number_format = '#,##0.00'
                if col_idx == 17 and value:  # CAPEX/VV %
                    cell.number_format = '0.00'
                
                # Colorer les montants CAPEX
                if col_idx == 9 and value and value > 0:
                    cell.font = Font(bold=True, color="2E7D32")
        
        # Ligne de total
        row += 2
        ws.cell(row=row, column=1, value="TOTAL GÉNÉRAL").font = Font(bold=True, size=11)
        
        total_col = 9  # Colonne CAPEX TOTAL
        col_letter = get_column_letter(total_col)
        total_cell = ws.cell(row=row, column=total_col, 
                            value=f"=SUM({col_letter}4:{col_letter}{row-2})")
        total_cell.font = Font(bold=True, size=11)
        total_cell.fill = self.styles['total']['fill']
        total_cell.number_format = '#,##0'
        
        # Totaux par catégorie
        for i, col_idx in enumerate([10, 11, 12, 13, 14, 15], 1):
            col_letter = get_column_letter(col_idx)
            cell = ws.cell(row=row, column=col_idx, 
                          value=f"=SUM({col_letter}4:{col_letter}{row-2})")
            cell.font = Font(bold=True)
            cell.fill = self.styles['total']['fill']
            cell.number_format = '#,##0'
        
        self._auto_width(ws)
    
    def _create_capex_synthese_sheet(self, ws):
        """Crée l'onglet CAPEX Synthèse"""
        # Titre
        ws.merge_cells('A1:P1')
        ws['A1'] = "📊 SYNTHÈSE CAPEX PAR SCPI"
        ws['A1'].font = Font(bold=True, size=14, color="2E7D32")
        ws['A1'].alignment = Alignment(horizontal='center')
        
        if not self.results['capex_synthese']:
            ws['A3'] = "Aucune synthèse CAPEX"
            return
        
        # Tableau principal
        headers = [
            'SCPI',
            'CAPEX Total €', 'CAPEX Directs €', 'CAPEX SCI €',
            'Nb Actifs', 'Nb avec CAPEX', '% avec CAPEX',
            'CAPEX Moyen €', 'CAPEX Médian €', 'Min €', 'Max €',
            'CAPEX/VV %', 'CAPEX/m² €', 'CAPEX/Loyers %',
            'Surface m²'
        ]
        
        row = 3
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            self._apply_header_style(cell, 'header_capex')
        
        for synthese in self.results['capex_synthese']:
            row += 1
            values = [
                synthese.scpi_source,
                synthese.total_capex, synthese.capex_actifs_directs, synthese.capex_sci,
                synthese.nb_actifs_total, synthese.nb_actifs_avec_capex, synthese.pct_actifs_avec_capex,
                synthese.capex_moyen_par_actif, synthese.capex_median_par_actif,
                synthese.capex_min, synthese.capex_max,
                synthese.capex_sur_vv_pct, synthese.capex_par_m2, synthese.capex_sur_loyers_pct,
                synthese.surface_totale_concernee
            ]
            
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx, value=value)
                cell.border = self.styles['border']
                
                if col_idx in [2, 3, 4, 8, 9, 10, 11, 15] and value:
                    cell.number_format = '#,##0'
                if col_idx in [7, 12, 14] and value:
                    cell.number_format = '0.00'
                if col_idx == 13 and value:
                    cell.number_format = '0.00'
                
                if col_idx == 2:
                    cell.font = Font(bold=True)
        
        # Ligne totale
        row += 1
        ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
        
        for col_idx in [2, 3, 4, 5, 6, 15]:
            col_letter = get_column_letter(col_idx)
            cell = ws.cell(row=row, column=col_idx, 
                          value=f"=SUM({col_letter}4:{col_letter}{row-1})")
            cell.font = Font(bold=True)
            cell.fill = self.styles['total']['fill']
            cell.number_format = '#,##0'
        
        # Section ventilation par type
        row += 3
        ws.merge_cells(f'A{row}:G{row}')
        ws[f'A{row}'] = "📋 VENTILATION PAR TYPE DE TRAVAUX"
        ws[f'A{row}'].font = Font(bold=True, size=12)
        
        row += 1
        type_headers = ['SCPI', 'Rénovation %', 'Gros Entretien %', 'Conformité %', 
                       'Environnemental %', 'Amélioration %', 'Autres %']
        for col, header in enumerate(type_headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            self._apply_header_style(cell)
        
        for synthese in self.results['capex_synthese']:
            row += 1
            values = [
                synthese.scpi_source,
                synthese.pct_renovation, synthese.pct_gros_entretien,
                synthese.pct_mise_conformite, synthese.pct_environnemental,
                synthese.pct_amelioration, synthese.pct_autres
            ]
            
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx, value=value)
                cell.border = self.styles['border']
                if col_idx > 1 and value:
                    cell.number_format = '0.0'
        
        self._auto_width(ws)
    
    def _create_controles_sheet(self, ws):
        """Crée l'onglet Contrôles"""
        # Titre
        ws.merge_cells('A1:I1')
        ws['A1'] = "✅ RÉSULTATS DES TESTS DE COHÉRENCE"
        ws['A1'].font = Font(bold=True, size=14, color="1F4E79")
        ws['A1'].alignment = Alignment(horizontal='center')
        
        headers = ['SCPI', 'Catégorie', 'ID', 'Description', 
                   'Valeur', 'Référence', 'Écart %', 'Statut', 'Commentaire']
        
        row = 3
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            self._apply_header_style(cell)
        
        # Tri par SCPI puis catégorie
        sorted_controls = sorted(self.results['controles'], 
                                key=lambda x: (x.scpi_source, x.categorie, x.test_id))
        
        for ctrl in sorted_controls:
            row += 1
            values = [
                ctrl.scpi_source, ctrl.categorie, ctrl.test_id,
                ctrl.description, ctrl.valeur_testee, ctrl.valeur_reference,
                ctrl.ecart_pct, ctrl.statut, ctrl.commentaire
            ]
            
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx, value=value)
                cell.border = self.styles['border']
                
                # Coloration statut
                if col_idx == 8:
                    if value == "OK":
                        cell.fill = self.styles['ok']
                    elif value == "ALERTE":
                        cell.fill = self.styles['warning']
                    elif value == "ERREUR":
                        cell.fill = self.styles['error']
                    elif value == "INFO":
                        cell.fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
        
        # Résumé
        row += 3
        ws.merge_cells(f'A{row}:C{row}')
        ws[f'A{row}'] = "📈 RÉSUMÉ DES CONTRÔLES"
        ws[f'A{row}'].font = Font(bold=True, size=12)
        
        row += 1
        nb_ok = len([c for c in self.results['controles'] if c.statut == "OK"])
        nb_alerte = len([c for c in self.results['controles'] if c.statut == "ALERTE"])
        nb_erreur = len([c for c in self.results['controles'] if c.statut == "ERREUR"])
        nb_info = len([c for c in self.results['controles'] if c.statut == "INFO"])
        
        summary = [
            ("✅ Tests OK:", nb_ok, self.styles['ok']),
            ("⚠️ Alertes:", nb_alerte, self.styles['warning']),
            ("❌ Erreurs:", nb_erreur, self.styles['error']),
            ("ℹ️ Info:", nb_info, PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")),
        ]
        
        for label, count, fill in summary:
            ws.cell(row=row, column=1, value=label)
            cell = ws.cell(row=row, column=2, value=count)
            cell.fill = fill
            row += 1
        
        row += 1
        total = len(self.results['controles'])
        if total > 0:
            taux = nb_ok / total * 100
            ws.cell(row=row, column=1, value="📊 Taux de conformité:")
            ws.cell(row=row, column=2, value=f"{taux:.1f}%").font = Font(bold=True, size=12)
        
        self._auto_width(ws)
    
    # =========================================================================
    # EXPORT CSV
    # =========================================================================
    
    def export_csv(self, prefix: str = "scpi_2024"):
        """Exporte vers CSV"""
        outputs = []
        
        exports = [
            ('indicateurs', self.results['indicateurs']),
            ('patrimoine', self.results['actifs']),
            ('sci', self.results['sci']),
            ('comptes_courants', self.results['comptes_courants']),
            ('capex_detail', self.results['capex_detail']),
            ('capex_synthese', self.results['capex_synthese']),
            ('controles', self.results['controles']),
        ]
        
        for name, data in exports:
            if data:
                df = pd.DataFrame([asdict(d) for d in data])
                path = self.output_dir / f"{prefix}_{name}_v5.csv"
                df.to_csv(path, index=False, encoding='utf-8-sig')
                outputs.append(str(path))
        
        return outputs
    
    # =========================================================================
    # AFFICHAGE
    # =========================================================================
    
    def print_summary(self):
        """Affiche un résumé détaillé"""
        print(f"\n{'='*80}")
        print("CONSOLIDATION SCPI V5 - AVEC CAPEX DÉTAILLÉ PAR ACTIF")
        print(f"{'='*80}")
        
        print(f"\n📊 VOLUMES DE DONNÉES")
        print(f"   SCPI analysées:        {len(self.results['indicateurs']):>5}")
        print(f"   Actifs directs:        {len(self.results['actifs']):>5}")
        print(f"   SCI sous-jacentes:     {len(self.results['sci']):>5}")
        print(f"   Comptes courants:      {len(self.results['comptes_courants']):>5}")
        print(f"   Fiches CAPEX détail:   {len(self.results['capex_detail']):>5}")
        print(f"   Tests de cohérence:    {len(self.results['controles']):>5}")
        
        # CAPEX
        if self.results['capex_synthese']:
            total_capex = sum(s.total_capex for s in self.results['capex_synthese'])
            print(f"\n💰 CAPEX CONSOLIDÉ TOTAL: {total_capex:>15,.0f} €")
        
        # Contrôles
        if self.results['controles']:
            nb_ok = len([c for c in self.results['controles'] if c.statut == "OK"])
            nb_alerte = len([c for c in self.results['controles'] if c.statut == "ALERTE"])
            nb_erreur = len([c for c in self.results['controles'] if c.statut == "ERREUR"])
            
            print(f"\n✅ CONTRÔLES: {nb_ok} OK | {nb_alerte} Alertes | {nb_erreur} Erreurs")
            taux = nb_ok / len(self.results['controles']) * 100 if self.results['controles'] else 0
            print(f"   Taux de conformité: {taux:.1f}%")
        
        # Détail par SCPI
        print(f"\n{'─'*80}")
        print("DÉTAIL PAR SCPI")
        print(f"{'─'*80}")
        
        for ind in self.results['indicateurs']:
            scpi = ind.nom_scpi or ind.source_pdf
            synthese = next((s for s in self.results['capex_synthese'] if s.scpi_source == scpi), None)
            nb_alertes = len([c for c in self.results['controles'] 
                             if c.scpi_source == scpi and c.statut in ['ALERTE', 'ERREUR']])
            
            print(f"\n🏢 {scpi}")
            print(f"   Société: {ind.societe_gestion or 'N/A'}")
            print(f"   TD: {ind.taux_distribution or 'N/A'}% | TOF: {ind.tof_annuel or 'N/A'}%")
            
            if synthese:
                print(f"   CAPEX Total: {synthese.total_capex:,.0f} € ({synthese.nb_actifs_avec_capex} actifs)")
                if synthese.capex_sur_vv_pct:
                    print(f"   CAPEX/VV: {synthese.capex_sur_vv_pct:.2f}% | CAPEX/m²: {synthese.capex_par_m2 or 0:.2f} €")
            
            if nb_alertes > 0:
                print(f"   ⚠️ {nb_alertes} alerte(s)")


# =============================================================================
# MAIN
# =============================================================================

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="SCPI Consolidator V5")
    parser.add_argument("-o", "--output", help="Dossier de sortie", default=".")
    args = parser.parse_args()
    
    consolidator = SCPIConsolidatorV5(args.output)
    consolidator.consolidate()
    
    excel_path = consolidator.export_excel()
    csv_paths = consolidator.export_csv()
    
    consolidator.print_summary()
    
    print(f"\n📁 Fichiers générés:")
    print(f"   Excel: {excel_path}")
    for csv in csv_paths:
        print(f"   CSV: {csv}")


if __name__ == "__main__":
    main()
