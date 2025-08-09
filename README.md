# Automatisation Reportings Entreprise VBA

Ce fichier Excel est un tableau de reporting des ventes mensuelles de produits alimentaires.  
Il sert à suivre les performances de vente pour un mois donné .

### Contenu du fichier

- **Colonnes principales :**  
  - Date  
  - Référence produit  
  - Poids  
  - Désignation  
  - Labels (Gluten Free, Bio)  
  - Marque  
  - TVA  
  - Prix de vente  
  - Quantités vendues sur le mois  
  - Chiffre d’affaires par produit (calculé automatiquement)  
  - Fournisseur  

- **Calculs intégrés :**  
  - Chiffre d’affaires par produit = Prix de vente × Quantités vendues  
  - Chiffre d’affaires total sur le mois  

---

## Objectifs du projet

- Automatiser le calcul du chiffre d’affaires par produit et total.  
- Mettre en forme automatiquement les résultats pour une meilleure lisibilité.  
- Normaliser et nettoyer les données (labels Gluten Free et Bio).  
- Générer un titre dynamique indiquant le mois du reporting.  
- Produire des statistiques détaillées des ventes (moyenne, min, max, etc.) sur une feuille dédiée.  
- Permettre un nettoyage complet du reporting pour repartir à zéro facilement.  

![excel](https://github.com/Torkiell-Angoria/VBA-Excel-Automatisation-Reportings-Entreprise/blob/main/img/excel.gif)

## Outils utilisés

- **Microsoft Excel** (tableur)  
- **Macros VBA** (Visual Basic for Applications) pour automatiser les calculs, la mise en forme et la génération de rapports.  

---

## Description des macros principales

### `Sub Mise_a_jour_reporting(SeuilColor As Integer)`

- Insère une colonne "Chiffre d’affaires".  
- Calcule le CA par produit (Prix de vente × Quantité).  
- Formate la colonne CA en monétaire.  
- Insère deux lignes en haut du tableau pour titre et total.  
- Calcule et affiche la somme totale du CA.  
- Appelle :  
  - `titre_reporting` (affiche le titre avec le mois)  
  - `mise_en_forme_CA` (colore le CA selon un seuil donné)  
  - `glutenfree` (normalise la colonne Gluten Free)  
- Remplace "Oui" par "Bio" dans la colonne BIO.  
- Crée une feuille "Calcul" avec statistiques sur les quantités vendues.  

### `Sub glutenfree()`

- Parcourt la colonne dédiée au label Gluten Free.  
- Remplace toutes les occurrences "Oui" par "Gluten Free".  

### `Sub mise_en_forme_CA(SeuilColor As Integer)`

- Parcourt la colonne Chiffre d’affaires.  
- Met en rouge gras si CA < seuil, sinon vert gras.  

### `Sub titre_reporting()`

- Récupère la date du jour.  
- Extrait le mois et le convertit en nom français.  
- Écrit un titre dynamique dans la cellule E1, ex : "Reporting du mois de Août".  

### `Sub Supprimer_Lignes_Colonnes()`

- Supprime les deux premières lignes de la feuille "Main".  
- Supprime la colonne "Chiffre d’affaires".  
- Supprime la feuille "Calcul" si elle existe.  
- Affiche des messages de confirmation.

![macro](https://github.com/Torkiell-Angoria/VBA-Excel-Automatisation-Reportings-Entreprise/blob/main/img/macro.gif)
