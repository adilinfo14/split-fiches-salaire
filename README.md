# Split fiches de salaire (PDF)

Outil Python pour découper un PDF contenant plusieurs fiches de salaire
en un fichier PDF par employé.

## Fonctionnalités
- 1 page = 1 fiche de salaire
- Nommage automatique : NOM_PRENOM_MM-YYYY.pdf
- Génération d’un fichier de log
- Gestion des erreurs page par page
- Compatible PyInstaller (.exe)

## Installation
```bash
pip install -r requirements.txt

## Utilisation
```bash
python src/split_fiches.py "Décomptes salaires 12.2025.pdf"

