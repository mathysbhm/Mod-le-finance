# Analyse et projection financière d'une entreprise (2019–2027)

Ce projet Python automatise l'analyse financière d'une entreprise en utilisant des données historiques, calcule des ratios clés, génère des projections jusqu’en 2027, et visualise les résultats dans un graphique clair.

---

## Objectifs du projet

- Charger des données financières (chiffre d'affaires, résultat net, capitaux propres, dettes…)
- Calculer automatiquement les ratios suivants :
  - Marge nette
  - ROE (Return on Equity)
  - Gearing (endettement net)
- Estimer la croissance moyenne annuelle à partir de l’historique
- Générer des projections financières sur plusieurs années (2024–2027)
- Visualiser l’évolution des ratios réels vs projetés

---

## Technologies utilisées

- Python 3.11+
- pandas (nécessaire pour traitement d'un fichier excel)
- matplotlib (nécessaire pour visualisation graphique)
- openpyxl (nécessaire pour export Excel)

---

## Fichiers du projet

- `entreprise.xlsx` : données financières historiques de l’entreprise
- `Python calculateur excel.py` : script principal avec calculs, projections et visualisation
- `rapport financier complet.xlsx` : export final (réel + projections)
- `README.md` : description du projet

---

## Exécution du script

1. Cloner le repo 
2. Installer les dépendances 
3. Lancer le script 

Le script affichera les ratios financiers dans le terminal et générera un graphique.

##À venir

- Ajout d’un mode interactif via Streamlit
- Lecture automatique des données depuis Yahoo Finance (API)

##Auteur

Projet réalisé par Mathys Brahmia Ferrier dans le cadre d'une montée en compétence Python & Finance d'entreprise.