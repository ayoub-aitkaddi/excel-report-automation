"""Consolide tous les fichiers Excel d'un dossier en un rapport annuel mis en forme.

Usage :
    python build_report.py                 # lit ./data, écrit rapport_ventes_2026.xlsx
    python build_report.py dossier sortie.xlsx
"""
from pathlib import Path
import sys

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ENTETE = PatternFill("solid", fgColor="1F3A5F")
FONT_ENTETE = Font(bold=True, color="FFFFFF")
EUROS = '#,##0.00 "€"'


def charger(dossier: Path) -> pd.DataFrame:
    fichiers = sorted(dossier.glob("*.xlsx"))
    if not fichiers:
        sys.exit(f"Aucun fichier .xlsx trouvé dans {dossier}")
    df = pd.concat([pd.read_excel(f).assign(Fichier=f.name) for f in fichiers], ignore_index=True)
    print(f"{len(fichiers)} fichiers lus, {len(df)} lignes brutes")
    return df


def nettoyer(df: pd.DataFrame) -> pd.DataFrame:
    avant = len(df)
    df = df.dropna(how="all", subset=["Date", "Magasin", "Produit"])
    df["Magasin"] = df["Magasin"].str.strip().str.title()
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.drop_duplicates(subset=["Date", "Magasin", "Produit", "Quantité", "Prix unitaire"])
    df["CA"] = df["Quantité"] * df["Prix unitaire"]
    df["Mois"] = df["Date"].dt.to_period("M").astype(str)
    print(f"Nettoyage : {avant - len(df)} lignes vides ou en double supprimées")
    return df.drop(columns="Fichier").sort_values("Date")


def mettre_en_forme(chemin: Path) -> None:
    wb = load_workbook(chemin)
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill, cell.font = ENTETE, FONT_ENTETE
            cell.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"
        for col in ws.columns:
            largeur = max(len(str(c.value or "")) for c in col) + 3
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(largeur, 30)
            if col[0].value in ("CA", "Prix unitaire") or str(col[0].value).startswith("CA"):
                for c in col[1:]:
                    c.number_format = EUROS

    ws = wb["Synthèse mensuelle"]
    graph = BarChart()
    graph.title, graph.y_axis.title = "Chiffre d'affaires par mois", "€"
    graph.add_data(Reference(ws, min_col=2, min_row=1, max_row=ws.max_row), titles_from_data=True)
    graph.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    graph.width, graph.height = 22, 10
    ws.add_chart(graph, "E2")
    wb.save(chemin)


def main() -> None:
    dossier = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "data"
    sortie = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).parent / "rapport_ventes_2026.xlsx"

    df = nettoyer(charger(dossier))
    mensuel = df.groupby("Mois", as_index=False).agg(CA=("CA", "sum"), Ventes=("Quantité", "sum"))
    magasins = (df.pivot_table(index="Magasin", columns="Mois", values="CA", aggfunc="sum", fill_value=0)
                  .assign(**{"CA total": lambda t: t.sum(axis=1)}).sort_values("CA total", ascending=False))
    produits = (df.groupby("Produit", as_index=False).agg(Quantité=("Quantité", "sum"), CA=("CA", "sum"))
                  .sort_values("CA", ascending=False))

    with pd.ExcelWriter(sortie, engine="openpyxl") as xl:
        mensuel.to_excel(xl, "Synthèse mensuelle", index=False)
        magasins.reset_index().to_excel(xl, "Par magasin", index=False)
        produits.to_excel(xl, "Top produits", index=False)
        df.to_excel(xl, "Données nettoyées", index=False)
    mettre_en_forme(sortie)
    print(f"Rapport écrit : {sortie} (CA total {df['CA'].sum():,.2f} €)")


if __name__ == "__main__":
    main()
