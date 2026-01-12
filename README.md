# SCPI Consolidator V5

Outil d'extraction et de consolidation des données réglementaires des rapports annuels SCPI 2024.

## Structure

```
retail/
├── input/              # Rapports annuels PDF (2024)
├── output/             # Fichiers générés (Excel, CSV)
├── scpi_consolidator_v5.py  # Script principal
└── requirements.txt    # Dépendances Python
```

## Installation

```bash
pip install -r requirements.txt
```

## Utilisation

```bash
python scpi_consolidator_v5.py
```

## Indicateurs Réglementaires Extraits (2024)

### Indicateurs de Performance (ASPIM/AMF)
- **TD** - Taux de Distribution
- **TOF** - Taux d'Occupation Financier
- **TOP** - Taux d'Occupation Physique
- **TRI** - Taux de Rentabilité Interne
- **PGA** - Performance Globale Annuelle
- **RGI** - Rendement Global Immobilier

### Valeurs Patrimoniales
- Valeur de Réalisation
- Valeur de Reconstitution
- Prix de souscription / retrait
- Capitalisation

### Données Financières
- Report à Nouveau (RAN)
- Provision pour Gros Entretien (PGE)
- Taux d'endettement / LTV
- Dividende par part

### Données Locatives
- WALT (durée moyenne des baux)
- WALB (durée avant résiliation)
- Nombre d'actifs / locataires

### Patrimoine Immobilier
- Liste des actifs directs (adresse, surface, valeur vénale)
- Répartition géographique
- Répartition typologique
- SCI sous-jacentes

### CAPEX
- Travaux de rénovation
- Gros entretien
- Mise en conformité
- Travaux environnementaux

## Output

- `scpi_consolidation_v5.xlsx` - Fichier Excel complet multi-onglets
- `scpi_2024_indicateurs_v5.csv` - Indicateurs par SCPI
- `scpi_2024_patrimoine_v5.csv` - Liste des actifs immobiliers
- `scpi_2024_sci_v5.csv` - SCI détenues
- `scpi_2024_capex_synthese_v5.csv` - Synthèse CAPEX
- `scpi_2024_controles_v5.csv` - Tests de cohérence

## Conformité Réglementaire

Données conformes aux exigences:
- AMF (Autorité des Marchés Financiers)
- ASPIM (Association française des Sociétés de Placement Immobilier)
- Code Monétaire et Financier (Art. L214-86 à L214-118)
- Ordonnance n°2024-662 du 3 juillet 2024
