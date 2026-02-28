# POC Dashboard Python

Ce projet est une preuve de concept d'une application **Flask** qui génère des rapports de rendez‑vous.

## Structure des données
Chaque enregistrement contient les colonnes suivantes :

- `date_rdv` : date du rendez-vous
- `canal` : téléphone ou web
- `profil` : type de client
- `motif` et `sous_motif` : raison du rendez-vous
- `date_creation` : date de création du rendez-vous (peut précéder la date du rdv)
- `id_conseiller` : identifiant du conseiller
- `bureau` : ville (un bureau par ville)
- `mois`, `annee_mois` : mois et année utilisés pour l'agrégat
- `nb_rdv` : nombre de rendez-vous agrégés

Le script `app.py` crée des données factices si aucun fichier n'existe.

## Indicateurs calculés
La fonction `process_data()` produit :

1. Volumes de rendez-vous (total, mensuel, hebdomadaire)
2. Délai de traitement moyen
3. Répartition journalière cumulée
4. Volume planifié sur une fenêtre [M-2 , M+1] relative au dernier mois
5. Vision des motifs et des sous-motifs
6. Classement des régions et évolution mois‑à‑mois
7. Tendance d'évolution mensuelle de la charge

Quelques graphiques sont générés dans `output/` et intégrés dans le PowerPoint.

## Prévisions (forecast)

Une nouvelle fonctionnalité de prévision hebdomadaire a été ajoutée pour estimer
les volumes de rendez-vous des 6 mois suivants en s'appuyant sur les 6 derniers
mois réels. Le dossier `forecast/` contient une petite bibliothèque Python :

``forecast/model.py``

contenant la logique d'entraînement (`ExponentialSmoothing` de `statsmodels`)
et de génération de graphique. Le module entraîne un modèle simple, le
sauvegarde en ``output/forecast_model.pkl`` et produit ``output/forecast.png``.

Lors de l'appel standard (`/integrate`), le module est exécuté automatiquement
et la prévision figure dans le PPTX final.

Le code de génération de jeu de données a également été augmenté (6000 lignes) dans
`app.py` pour fournir suffisamment d'historique.

L'indicateur "prévision du nombre de rendez-vous" est inclus dans les résultats
et le PPTX final.

N'oubliez pas d'installer la dépendance supplémentaire :
```powershell
pip install statsmodels
```
## Installation et lancement
```powershell
cd path\to\poc-dashboard-python
python -m venv venv  # optionnel
.\venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
python app.py
```

Visiter `http://localhost:5000` pour déclencher les calculs, puis télécharger le rapport.

## Génération d’email/template 
La route `/generate-email` crée désormais **un modèle Outlook (.oft)** si le module `pywin32`
est présent et qu’Outlook est installé sur la machine. Le template reprend le sujet, les
destinataires, le corps et attache le PPTX. En cas d’absence de `win32com`, on retombe sur
l’ancien comportement qui exporte un `.eml`.

Pour utiliser l’OFT il faut ajouter `pywin32` à `requirements.txt` et lancer `pip install` :

```powershell
pip install pywin32
```

> ✨ `matplotlib` est utilisé pour les graphiques; vous pouvez ajouter `seaborn` dans
> `requirements.txt` si vous souhaitez des visualisations plus sophistiquées.
