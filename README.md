# Rapport Excel automatisé : 12 fichiers mensuels → 1 rapport en 1 seconde

**Problème** : chaque mois, une PME reçoit un fichier Excel de ventes par période. Les consolider à la main prend des heures : copier-coller, doublons, lignes vides, noms de magasins écrits de trois façons différentes (« Metz », « METZ », « metz »).

**Solution** : un script Python qui lit tout le dossier, nettoie les données et produit un rapport Excel prêt à présenter.

| Avant | Après |
| --- | --- |
| 12 fichiers, 2 573 lignes brutes | 1 rapport, 4 onglets mis en forme |
| Doublons et lignes vides | 174 lignes parasites supprimées automatiquement |
| ~2 h de copier-coller par mois | 1 seconde, en un clic |

## Ce que contient le rapport

- **Synthèse mensuelle** : CA et volumes par mois, avec graphique Excel natif
- **Par magasin** : tableau croisé magasin × mois, trié par CA total
- **Top produits** : classement par chiffre d’affaires
- **Données nettoyées** : la base consolidée, filtrable

En-têtes colorés, colonnes ajustées, montants au format €, volets figés.

## Lancer la démo

```bash
pip install pandas openpyxl
python generate_sample_data.py   # crée 12 fichiers de ventes fictives dans data/
python build_report.py           # écrit rapport_ventes_2026.xlsx
```

Sur vos propres fichiers : `python build_report.py chemin/du/dossier mon_rapport.xlsx`

## Adapter à votre cas

Le même principe fonctionne pour des exports de caisse, des relevés, des fichiers fournisseurs, des inscriptions ou des PDF de factures. Chaque besoin est adapté à la structure réelle de vos fichiers.

---

Réalisé par **Ayoub Ait Kaddi**, étudiant ingénieur à Télécom SudParis · Automatisation Python & data
