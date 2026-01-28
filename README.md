# Gantt Chart Generator -

## Description

Script Python d'automatisation complète pour générer des diagrammes de Gantt à partir des données d'inventaire des modèles murins génétiquement modifiés.

**Fonctionnalités:**
- ✅ Parsing automatique du fichier Excel (feuille "tableau complet")
- ✅ Création d'un Gantt chart **par product line**
- ✅ Identification des modèles par **internal code**
- ✅ Sous-titre affichant **target** + **dates HO/validation**
- ✅ Export en **slides HTML interactives** (13 slides)
- ✅ Export en **CSV de synthèse**
- ✅ Visualisation des timelines en SVG natif (pas de dépendance graphique)

---

## Installation

### Prérequis
- Python 3.8+
- pandas
- openpyxl

### Setup

```bash
# Installer les dépendances
pip install pandas openpyxl

# Vérifier l'installation
python gantt_generator.py --help
```

---

## Utilisation

### Mode basique
```bash
python gantt_generator.py models-list-20260105-CRUPPE.xlsx
```

Génère:
- `gantt_slides.html` - Slides HTML (13 pages)
- `gantt_models_export.csv` - Données d'export

### Mode personnalisé
```bash
python gantt_generator.py models-list-20260105-CRUPPE.xlsx \
    --html custom_gantt.html \
    --csv custom_export.csv
```

---

## Architecture des Gantt Charts

### Structure par Product Line

Un Gantt chart est généré **pour chaque product line** trouvée dans les données:

1. **models to target sensors triggering the inflammation** (9 modèles)
2. **models to target sensors triggering the inflammasome** (6 modèles)
3. **models to target inflammatory and signaling pathways** (8 modèles)
4. **models to target cytokines and chemokines** (18 modèles)
5. **Models to target release of mediators triggering inflammation** (23 modèles)
6. **Models targeting autoreactive B cells** (7 modèles)
7. **models targeting T cells** (39 modèles) ← Le plus grand
8. **Models to target myeloid cells** (4 modèles)
9. **Models to enhance predictibility of PK** (14 modèles)
10. **models of TAA** (5 modèles)
11. **A classer** (9 modèles)

**Total: 142 modèles valides**

### Format de chaque Gantt

**Axe vertical (Y):**
- Code interne (ex: ICP150)
- Target associée (ex: cGAS/STING)

**Axe horizontal (X):**
- Timeline de 2025-12-24 à 2027-08-31 (~600 jours)
- Grille de dates tous les 90 jours

**Barres Gantt:**
- Longueur = durée entre disponibilité HO et fin validation
- Couleur = différenciation visuelle par modèle
- Tooltip = détails complets (code, target, dates, durée)
- Texte = durée en jours (si place suffisante)

---

## Fichiers de sortie

### 1. `gantt_slides.html`

Presentation HTML interactive avec:
- **Slide 1**: Couverture (résumé global)
- **Slides 2-12**: Un Gantt chart par product line
- **Slide 13**: Résumé statistique en tableau

**Visualisation:**
- Ouvrir dans un navigateur web
- Imprimer en PDF (Ctrl+P)
- Zoom disponible (Ctrl++ / Ctrl+-)

### 2. `gantt_models_export.csv`

Données tabulaires pour analyse complémentaire:
- Product Line
- Internal Code
- Target
- Status
- Date Disponibilité HO
- Date Fin Validation
- Durée (jours)

**Utilisation:**
- Importer dans Excel pour analyses supplémentaires
- Créer des filtres et tris personnalisés
- Générer des statistiques

---

## Spécifications techniques

### Colonnes extraites
```
- product line          : Catégorie principale (axe de séparation)
- internal code         : Identifiant unique du modèle (axe Y)
- target                : Gène/protéine ciblés (sous-titre)
- status                : État du modèle (validation, catalog, etc.)
- date disponibilité HO : Date de départ de la barre (disponibilité HO)
- data de fin validation: Date de fin de la barre (validation complète)
- duration_days         : Durée calculée (end - start)
```

### Filtrage automatique
- Exclusion des modèles sans `internal code`
- Exclusion des modèles sans dates valides
- Nettoyage des valeurs manquantes
- Tri par date de disponibilité

### Dimensions SVG
- Largeur: 1300px (responsive)
- Hauteur: Adaptée au nombre de modèles (~35px par modèle)
- Marge gauche: 260px (pour les labels)
- Impression: Compatible avec A4 paysage

---

## Automatisation future

### Extensions possibles

1. **Mise à jour automatique**
   ```bash
   # Tâche cron (Linux/Mac)
   0 9 * * MON python /path/to/gantt_generator.py /path/to/data.xlsx
   ```

2. **Envoi par email**
   ```python
   # À ajouter dans gantt_generator.py
   import smtplib
   send_email('team@example.com', 'gantt_slides.html')
   ```

3. **Synchronisation cloud**
   ```bash
   # Après génération
   aws s3 cp gantt_slides.html s3://bucket/reports/
   ```

4. **Notifications Slack**
   ```python
   webhook_url = "https://hooks.slack.com/..."
   requests.post(webhook_url, json={"text": "✅ Gantt charts generés"})
   ```

---

## Dépannage

### Erreur: "impossible de lire 'tableau complet'"
- Vérifier le nom exact de la feuille Excel
- S'assurer que le fichier n'est pas ouvert dans Excel
- Utiliser la version récente du fichier

### Gantt vide (0 modèles)
- Vérifier la colonne "internal code" (ne doit pas être vide)
- Vérifier les colonnes de dates
- Exécuter en debug: `python -u gantt_generator.py file.xlsx`

### HTML ne s'affiche pas correctement
- Utiliser un navigateur moderne (Chrome, Firefox, Safari)
- Vérifier les caractères spéciaux (é, ô, etc.)
- Exporter en PDF: navigateur > Imprimer > Enregistrer en PDF

---

## Performance

| Métriques | Valeur |
|-----------|--------|
| Temps de traitement | ~2-3 secondes |
| Taille HTML générée | ~3-5 MB |
| Nombre de modèles traités | 142 |
| Nombre de product lines | 11 |
| Durée moyenne par modèle | 213 jours |

---

## Changelog

### Version 1.0 (2026-01-27)
- ✅ Initial release
- ✅ Support 11 product lines
- ✅ Export HTML + CSV
- ✅ Gantt charts SVG natif

---

## Contact & Support

**Développement:** CRUPPE - Biologie Moléculaire
**Email:** docteur@cruppe.fr
**Localisation:** Lyon, Rhône-Alpes

---

## Licence

Internal use - CRUPPE

