# Cockpit Trésorerie — CSE LIDL ENTZHEIM

Dashboard de pilotage trésorerie basé sur les exports Pennylane.

## Fonctionnalités

- KPIs temps réel (trésorerie, flux, budget vs réalisé)
- Graphiques interactifs (évolution trésorerie, répartition dépenses)
- Tableau de suivi avec indicateurs RAG (Rouge/Ambre/Vert)
- Alertes automatiques
- Upload du fichier Excel Pennylane pour mise à jour

## Mise à jour des données

1. Exporter le Plan de trésorerie depuis **Pennylane**
2. Mettre à jour l'onglet **RÉALISÉ** dans `data.xlsx`
3. Uploader `data.xlsx` dans le dashboard **OU** committer le fichier mis à jour sur GitHub

## Déploiement local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Technologies

- [Streamlit](https://streamlit.io) — interface web
- [Plotly](https://plotly.com) — graphiques interactifs
- [Pandas](https://pandas.pydata.org) — traitement des données
- Source : export Excel [Pennylane](https://www.pennylane.com)
