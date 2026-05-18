# 📊 EEX Prix — Collecte automatique des prix de clôture

> Collecte quotidienne des **settlement prices EEX** (Power FR Baseload, Peakload & Natural Gas PEG)
> via GitHub Actions → fichier Excel cumulatif versionné dans le repo.

---

## Table des matières

1. [Vue d'ensemble](#1-vue-densemble)
2. [Structure du repo](#2-structure-du-repo)
3. [Fonctionnement pas à pas](#3-fonctionnement-pas-à-pas)
4. [Contrats collectés](#4-contrats-collectés)
5. [API EEX — Référence technique](#5-api-eex--référence-technique)
6. [Fichier de sortie `data/eex_prix.xlsx`](#6-fichier-de-sortie-dataeex_prixxlsx)
7. [GitHub Actions — CI/CD](#7-github-actions--cicd)
8. [Installation locale](#8-installation-locale)
9. [Maintenance & dépannage](#9-maintenance--dépannage)
10. [Règles métier](#10-règles-métier)
11. [Historique des versions](#11-historique-des-versions)

---

## 1. Vue d'ensemble

### Objectif

Ce repo collecte automatiquement, **chaque matin lundi–vendredi**, les settlement prices publiés la veille à 21h00 CET par l'EEX pour :

| Commodité | Granularités collectées | Préfixe label |
|---|---|---|
| Power France Baseload | CAL 2027–2030, QTR 2027–2030, 18 mois glissants | `ELEC-` |
| Power France Peakload | CAL 2027–2030 | `PEAK-` |
| Natural Gas PEG | CAL 2027–2030, QTR 2027–2030, 18 mois glissants | `GAZ-` |

**~70 contrats collectés par jour**, ajoutés dans `data/eex_prix.xlsx` versionné dans le repo.

### Timing

> ⚠️ Les prix EEX sont publiés à **21h00 CET** (la veille).
> La collecte tournant à 1h00 UTC (hiver) / 0h00 UTC (été) récupère donc les settlements de **J-1**.

---

## 2. Structure du repo

```
├── .github/
│   └── workflows/
│       └── eex_daily.yml        # GitHub Actions — déclenchement & push auto
├── data/
│   └── eex_prix.xlsx            # Fichier Excel cumulatif (versionné)
├── scripts/
│   └── collect_eex.py           # Script principal de collecte
├── requirements.txt             # requests==2.32.3, openpyxl==3.1.5
└── .gitignore
```

---

## 3. Fonctionnement pas à pas

```
GitHub Actions (cron 0h00 ou 1h00 UTC, L–V)
        │
        ▼
1. Checkout du repo (fetch-depth: 0)
        │
        ▼
2. Setup Python 3.11 + pip install requirements.txt + python-dateutil
        │
        ▼
3. python scripts/collect_eex.py
   ├── Génère ~70 contrats dynamiques (CAL/QTR/MTH)
   ├── Charge data/eex_prix.xlsx (ou le crée si absent)
   ├── Vérifie anti-doublon (date du jour déjà présente ?)
   ├── Appel API EEX pour chaque contrat (GET, sleep 0.4s)
   │     ├── Prix reçu  → ligne "Settlement"
   │     ├── Prix null  → fallback dernier connu → ligne "Reporté J-1" (fond jaune)
   │     └── Prix null + pas d'historique → ligne ignorée
   └── Sauvegarde eex_prix.xlsx (mise en forme complète)
        │
        ▼
4. git add data/eex_prix.xlsx
   git commit -m "data: EEX prices YYYY-MM-DD"
   git push
```

---

## 4. Contrats collectés

### Power FR Baseload (`ELEC-`)

| Type | Contrats | ShortCode EEX |
|---|---|---|
| CAL | `ELEC-CAL-2027` à `ELEC-CAL-2030` | `F7BY` |
| QTR | `ELEC-Q1/Q2/Q3/Q4-2027` à `2030` (16 contrats) | `F7BQ` |
| MTH | `ELEC-Jan-YYYY` … 18 mois glissants depuis aujourd'hui | `F7BM` |

### Power FR Peakload (`PEAK-`)

| Type | Contrats | ShortCode EEX |
|---|---|---|
| CAL | `PEAK-CAL-2027` à `PEAK-CAL-2030` | `F7PY` |

### Natural Gas PEG (`GAZ-`)

| Type | Contrats | ShortCode EEX |
|---|---|---|
| CAL | `GAZ-CAL-2027` à `GAZ-CAL-2030` | `G5BY` |
| QTR | `GAZ-Q1/Q2/Q3/Q4-2027` à `2030` (16 contrats) | `G5BQ` |
| MTH | `GAZ-Jan-YYYY` … 18 mois glissants depuis aujourd'hui | `G5BM` |

> **Pourquoi les préfixes ?**
> Sans `ELEC-`/`GAZ-`, les labels `CAL-2027` ou `Q1-2027` seraient identiques entre électricité et gaz. Les formules RECHERCHEV/VLOOKUP prenaient la première occurrence (gaz ~25 €/MWh au lieu d'électricité ~55 €/MWh). Les préfixes éliminent définitivement cette ambiguïté.

---

## 5. API EEX — Référence technique

### Endpoint

```
GET https://api.eex-group.com/pub/market-data/price-ticker
```

### Paramètres

| Paramètre | Description | Valeurs utilisées |
|---|---|---|
| `shortCode` | Code produit EEX | `F7BY`, `F7BQ`, `F7BM`, `F7PY`, `G5BY`, `G5BQ`, `G5BM` |
| `area` | Zone géographique | `FR` (Power) · `PEG` (Gas) |
| `product` | Famille produit | `Base` · `Peak` · `Physical` |
| `commodity` | Matière première | `POWER` · `NATGAS` |
| `pricing` | Type de pricing | `F` (Futures) |
| `maturity` | Code de maturité | Format `YYYYMM` |

### Encodage des maturités

| Type | Format | Exemples |
|---|---|---|
| CAL (année) | `YYYY01` | `202701`, `202801` |
| Q1 | `YYYY01` | `202701` |
| Q2 | `YYYY04` | `202704` |
| Q3 | `YYYY07` | `202707` |
| Q4 | `YYYY10` | `202710` |
| Month | `YYYYMM` | `202601` = Jan 2026 |

### Extraction du prix dans la réponse JSON

```python
data    = r.json()
headers = data.get("header", [])    # ex. ["date", "settlPx", ...]
idx     = headers.index("settlPx")  # index de la colonne settlement
price   = data["data"][0][idx]      # première ligne = contrat demandé
```

### Rate limiting

Un `time.sleep(0.4)` est appliqué après chaque appel.
Pour ~70 contrats → durée totale d'exécution ≈ **30 secondes**.

---

## 6. Fichier de sortie `data/eex_prix.xlsx`

### Onglets

| Onglet | Commodités incluses |
|---|---|
| `Power FR` | Power FR Baseload + Power FR Peakload |
| `Natural Gas PEG` | Natural Gas PEG |

### Colonnes (identiques dans les deux onglets)

| # | Colonne | Type | Exemple |
|---|---|---|---|
| A | `Date` | Date `YYYY-MM-DD` | `2026-05-12` |
| B | `Contrat` | Texte | `ELEC-CAL-2027` |
| C | `Type` | Texte | `CAL`, `QTR`, `MTH` |
| D | `Prix (€/MWh)` | Nombre `#,##0.000` | `58.420` |
| E | `Source` | Texte | `Settlement` ou `Reporté J-1` |

### Mise en forme automatique (appliquée à chaque run)

| Élément | Style |
|---|---|
| En-tête ligne 1 | Fond bleu marine `#1F4E79`, texte blanc, gras, hauteur 22px |
| Lignes paires | Fond bleu très clair `#EBF3FB` |
| Lignes impaires | Fond blanc |
| **Ligne "Reporté J-1"** | **Fond jaune `#FFF2CC` sur toute la ligne** |
| Colonne Prix | Aligné à droite, format `#,##0.000` |
| Filtre automatique | Activé sur toutes les colonnes |
| Freeze | Ligne 1 gelée (scroll vertical) |

---

## 7. GitHub Actions — CI/CD

### Déclenchement (`.github/workflows/eex_daily.yml`)

```yaml
on:
  schedule:
    - cron: "0 1 * * 1-5"   # hiver : 1h00 UTC = 2h00 CET
    - cron: "0 0 * * 1-5"   # été   : 0h00 UTC = 2h00 CEST
  workflow_dispatch:          # déclenchement manuel possible
```

> Les deux crons sont nécessaires pour couvrir le changement d'heure hiver/été.
> GitHub Actions tourne en UTC — sans les deux règles, le run serait décalé d'une heure 6 mois sur 12.

### Permissions requises

```yaml
permissions:
  contents: write   # pour que le bot puisse git push
```

### Déclencher manuellement

`GitHub → Actions → EEX Daily Price Collection → Run workflow`

### Lire les logs d'un run

`GitHub → Actions → [run du jour] → collect → Run collector`

Sortie attendue :
```
=== Collecte EEX — 2026-05-12 ===
  ELEC-CAL-2027: 58.420 €/MWh
  ELEC-Q1-2027:  55.100 €/MWh
  ...
  GAZ-CAL-2027:  24.350 €/MWh  ⚠️ reporté
✓ Terminé — 68 lignes ajoutées (2 reportées, 0 ignorées)
```

---

## 8. Installation locale

### Prérequis

- Python 3.8+
- Git

### Setup

```bash
git clone <url-du-repo>
cd <repo>
pip install -r requirements.txt python-dateutil
```

### Lancer manuellement

```bash
python scripts/collect_eex.py
```

Le fichier `data/eex_prix.xlsx` est créé ou mis à jour localement.

### Vérification des dépendances

```bash
python -c "import requests, openpyxl; from dateutil.relativedelta import relativedelta; print('OK')"
```

---

## 9. Maintenance & dépannage

### Le workflow ne se déclenche plus

GitHub désactive automatiquement les crons sur les repos **inactifs depuis plus de 60 jours**.

➜ Aller dans `Actions → EEX Daily Price Collection → Enable workflow` ou pousser un commit.

### Erreur `403 Forbidden` sur l'API EEX

L'API EEX exige les headers `Origin` et `Referer`. Vérifier dans `collect_eex.py` :

```python
REQUEST_HEADERS = {
    "Origin":  "https://www.eex.com",
    "Referer": "https://www.eex.com/",
}
```

Si le problème persiste, l'EEX a peut-être renforcé sa protection — intercepter une requête réelle depuis Chrome pour copier les headers exacts.

### Erreur `429 Too Many Requests`

Augmenter le délai entre requêtes :

```python
time.sleep(0.4)  →  time.sleep(0.8)
```

### Un ShortCode ne retourne plus de données (API change)

1. Ouvrir `https://www.eex.com/en/market-data/market-data-hub` dans Chrome
2. `F12 → Network → filtrer "price-ticker"`
3. Naviguer sur le produit concerné → inspecter la requête XHR
4. Mettre à jour le `shortCode` dans `build_contracts()` dans `scripts/collect_eex.py`

### Ajouter un contrat (ex. CAL 2031)

Dans `scripts/collect_eex.py`, fonction `build_contracts()` :

```python
contracts.append({
    "label":     "ELEC-CAL-2031",
    "commodite": "Power FR Baseload",
    "type":      "CAL",
    "params": {
        "shortCode": "F7BY",
        "area":      "FR",
        "product":   "Base",
        "commodity": "POWER",
        "pricing":   "F",
        "maturity":  "203101",
    },
})
```

### Ajouter un nouvel onglet Excel

Dans `SHEET_MAP` en haut de `collect_eex.py` :

```python
SHEET_MAP = {
    "Power FR":          ["Power FR Baseload", "Power FR Peakload"],
    "Natural Gas PEG":   ["Natural Gas PEG"],
    "Mon Nouvel Onglet": ["Ma Nouvelle Commodite"],  # ← ajouter ici
}
```

### Réinitialiser le fichier Excel

```bash
git rm data/eex_prix.xlsx
git commit -m "reset: suppression eex_prix.xlsx pour recréation"
git push
```
Le script recrée automatiquement un fichier vierge au prochain run.

### Migration depuis l'ancien format mono-feuille

Le script détecte automatiquement les anciennes feuilles nommées `Prices` ou `Sheet` et migre les données vers les nouveaux onglets. Aucune action manuelle requise.

---

## 10. Règles métier

| Règle | Comportement |
|---|---|
| **Anti-doublon** | Si une ligne avec la date du jour existe déjà dans l'un des onglets → `sys.exit(0)`, rien n'est écrit |
| **Prix null → fallback J-1** | L'API retourne null → le dernier prix connu pour ce `Contrat` est repris, `Source = "Reporté J-1"`, fond jaune |
| **Prix null + aucun historique** | Ligne non écrite (ignorée silencieusement) |
| **Mois glissants** | Les 18 contrats MTH sont recalculés à chaque run depuis `date.today()` via `dateutil.relativedelta` |
| **Reformatage complet** | La mise en forme Excel (couleurs, bordures, filtres, largeurs) est réappliquée intégralement à chaque run |

---

## 11. Historique des versions

| Phase | Description |
|---|---|
| **Phase 1** | Collecte initiale Power FR Baseload (CAL, QTR, MTH) — découverte API par interception réseau Chrome |
| **Phase 2** | Ajout Natural Gas PEG (`G5BY`/`G5BQ`/`G5BM`) — mois glissants dynamiques |
| **Phase 3** | Ajout Power FR Peakload CAL 2027–2030 (shortCode `F7PY`) |
| **Phase 4** | Correction bug labels dupliqués → préfixes `ELEC-`/`PEAK-`/`GAZ-` |
| **Phase 5** | Migration CSV → XLSX natif — fallback J-1 automatique — coloration jaune |
| **Phase 6** | Fichier unique cumulatif + anti-doublon — passage à **GitHub Actions** |
| **Phase 7** | Deux onglets Excel (`Power FR` / `Natural Gas PEG`) — mise en forme unifiée |

---

*Maintenu par l'équipe Pricing — SAS Capitole Énergie*
