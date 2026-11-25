"""Génère 12 fichiers Excel mensuels de ventes fictives, volontairement « sales »
(doublons, lignes vides, noms de magasins incohérents) pour la démonstration."""
from pathlib import Path
import random

import pandas as pd

random.seed(42)
OUT = Path(__file__).parent / "data"
OUT.mkdir(exist_ok=True)

MAGASINS = ["Metz", "Nancy", "Forbach", "Strasbourg"]
VARIANTES = {"Metz": ["Metz", "METZ", "metz "], "Nancy": ["Nancy", "nancy"],
             "Forbach": ["Forbach", "FORBACH"], "Strasbourg": ["Strasbourg", "Strasbourg "]}
PRODUITS = {"Clavier": 39.9, "Souris": 19.9, "Écran 24\"": 149.0, "Casque": 59.0,
            "Webcam": 45.0, "Clé USB 64 Go": 12.5, "Disque SSD 1 To": 79.0}

for mois in range(1, 13):
    lignes = []
    for _ in range(random.randint(180, 260)):
        produit = random.choice(list(PRODUITS))
        magasin = random.choice(MAGASINS)
        jour = random.randint(1, 28)
        lignes.append({
            "Date": f"2026-{mois:02d}-{jour:02d}",
            "Magasin": random.choice(VARIANTES[magasin]),
            "Produit": produit,
            "Quantité": random.randint(1, 5),
            "Prix unitaire": PRODUITS[produit],
        })
    df = pd.DataFrame(lignes)
    # Défauts réalistes : doublons et lignes vides
    df = pd.concat([df, df.sample(5, random_state=mois)], ignore_index=True)
    vide = pd.DataFrame([{c: None for c in df.columns}] * 3)
    df = pd.concat([df, vide], ignore_index=True).sample(frac=1, random_state=mois)
    df.to_excel(OUT / f"ventes_2026_{mois:02d}.xlsx", index=False)

print(f"12 fichiers créés dans {OUT}")
