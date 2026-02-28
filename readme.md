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

> ✨ `matplotlib` est utilisé pour les graphiques; vous pouvez ajouter `seaborn` dans
> `requirements.txt` si vous souhaitez des visualisations plus sophistiquées.
