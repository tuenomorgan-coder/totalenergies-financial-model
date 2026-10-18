# TotalEnergies – Modèle de valorisation DCF

Projet de modélisation financière sur **TotalEnergies SE** (Euronext : TTE) : DCF à 5 ans (2026E–2030E) construit sur les comptes FY2025,
avec scénarios (Bas / Central / Haut) et table de sensibilité WACC × croissance terminale.
Projet d'entraînement type « financial statement literacy + DCF » pour la finance de marché / M&A énergie.

## Contenu

| Fichier | Rôle |
|---|---|
| `TotalEnergies_DCF.xlsx` | Modèle Excel (113 formules vivantes, 0 erreur) |
| `src/build_model.py` | Script Python (openpyxl) qui régénère le modèle |
| `data/sources.md` | Sources et limites des données |

### Onglets
- **Hypotheses** : sélecteur de scénario, WACC, g, impôt, EUR/USD, cours (à saisir), bilan FY2025, trajectoires par scénario
- **Historique** : ventes, résultats, CFFO, capex, coût de la dette 2024–2025
- **DCF** : FCFF = CFFO + charges financières après impôt – investissements nets ; valeur terminale (Gordon-Shapiro) ; passage VE → valeur par action
- **Sensibilite** : valeur par action (EUR) selon WACC (7–11 %) et g (0–2 %), avec cellule de contrôle

Code couleur : bleu = saisie, noir = formule, vert = lien inter-feuilles, jaune = hypothèse clé.

## Résultats (hypothèses illustratives, valorisation au 31/12/2025)

| Scénario | VE (Md$) | Capitaux propres (Md$) | Par action ($) | Par action (€) | Poids VT |
|---|---|---|---|---|---|
| Bas | 111,3 | 90,5 | 40,9 | 35,2 | 68 % |
| Central | 193,4 | 172,7 | 78,0 | 67,1 | 70 % |
| Haut | 234,4 | 213,7 | 96,5 | 83,0 | 70 % |

Ces valeurs ne sont **pas une recommandation d'investissement** : elles dépendent d'hypothèses non calibrées (WACC 9 %, g 1 %, EUR/USD 0,86 à actualiser).

## Utilisation

```bash
pip install -r requirements.txt
python src/build_model.py --scenario 2 --out TotalEnergies_DCF.xlsx
```
Ouvrir le fichier dans Excel, saisir le cours actuel (`Hypotheses!B10`) et le taux de change (`Hypotheses!B9`), puis changer le scénario en `Hypotheses!B5`.

## Limites et pistes d'amélioration
1. Remplacer la dette nette estimée (via gearing) par la dette nette publiée, et ajouter minoritaires / provisions de fin de vie des actifs
2. Passer d'un FCFF « cash-flow » à un modèle par segment (E&P, Integrated LNG, Integrated Power, Downstream) avec prix du Brent et production en kboe/j
3. Construire le WACC (CAPM, bêta, structure cible) au lieu de le poser
4. Ajouter un onglet de comparables (Shell, BP, Eni, Equinor, Exxon, Chevron) : EV/EBITDA, P/E, rendement FCF
5. Mettre à jour avec les résultats T3 2026 (29/10/2026) et la guidance de l'Investor Day
