# Golda2Prima

Le script golda2prima.py est conçu pour traiter des fichiers Excel contenant des données d'articles, en appliquant plusieurs transformations et extractions de colonnes spécifiques. Voici les principales étapes et fonctionnalités du script :

## Fonctionnalités du Script :

### Configuration des Chemins :

- Le script définit le dossier source où se trouvent les fichiers Excel à traiter.
- Il vérifie l'existence du dossier source et du fichier de correspondance nécessaire pour les transformations.

### Chargement des Données :

- Le script charge un fichier Excel de correspondance qui contient des mappages entre les codes sous-famille GOLDA et PRIMA.

Pour chaque fichier Excel dans le dossier source, le script effectue les opérations suivantes :

- Filtre les lignes avec des valeurs manquantes ou vides dans la colonne "Prix_euro".
- Convertit les poids en kilogrammes si la colonne "Poids" est présente.
- Ajoute des colonnes "REF APPEL 1" et "REF APPEL 2" basées sur "Ref_fournisseur".
- Met à jour les codes sous-famille et famille NU en utilisant le fichier de correspondance.

### Création des Articles de Consigne :

- Le script identifie les articles avec un montant de consigne et crée des lignes supplémentaires pour ces articles de consigne.
- Les articles de consigne sont ajoutés au fichier de sortie avec des colonnes spécifiques mises à jour.

Sauvegarde des Résultats :

- Le script extrait les colonnes nécessaires et les renomme pour correspondre à un format spécifique.
- Les données transformées sont sauvegardées dans un nouveau fichier Excel avec un nom basé sur le nom original du fichier et la date actuelle.

##  Création de fichier lecture équipementier

Pour utiliser le script de création de fichier lecture équipementier il faut :

1. Télécharger les fichiers « bruts » sur le site Golda des fournisseurs nécessaires.
2. Déplacer les fichiers téléchargés dans un dossier appelé Golda (voir ci-dessous).
3. Installer 7-Zip si ce n’est pas déjà fait.
4. Extraire les fichiers (clic droit - 7-Zip -> Extraire vers *)
5. Lancer le script golda2prima.py.

###  Structure du dossier

golda2prima/  
├── golda/  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── fichier1_xxx/  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── fichier1_xxx/fichier1_xxx.xlsx  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── fichier2_xxx/  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── fichier2_xxx/fichier2_xxx.xlsx  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── fichier1_xxx.zip (fichier ignoré)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── fichier2_xxx.zip (fichier ignoré)  
├── golda2prima.py (fichier exécutable contenant le code)  
├── Tableau_Correspondance_Golda.xlsx (correspondance FAM/SFAM GOLDA/PRIMA)  
├── Tableau_Correspondance_Marque.xlsx (correspondance CODE MARQUE GOLDA/PRIMA)

### Utilisation

Double cliquez sur golda2prima.py.

Une fenêtre apparait, indiquant le fichier sur lequel le script est en train de travailler,

Une fois terminé la fenêtre se ferme et les fichiers générés sont disponibles dans chaque sous dossier (où se trouve les fichiers « source »)

Pour recommencer une procédure il est préférable de supprimer définitivement les dossiers et de recommencer la procédure d’extraction des fichiers.

#### Famille / Sous-famille GOLDA / PRIMA

Le script est hyper dépendant du fichier de correspondance entre les familles golda et prima.
Plus se fichier sera remplit, pertinent et complet, plus les articles seront référencés dans les bonnes catégories.

Ce fichier nécessite encore des améliorations.

#### Tarif ID Rechange

Les fichiers disponibles sur la plateforme ID Rechange ne sont pas complets et nécessite que l’on dispose des informations articles soit dans notre base de données, soit dans un fichier du golda.
Il faut ensuite combiner les barèmes des fichiers IDR grâce au fichier généré de l’export article LECTURE_EQUIP_PIECE sur la sélection nécessaire.

#### Export article – lecture de tarif équipementier pneu / pièces

Afin de faciliter la procédure d’intégration de BF ID Rechange notre ERP dispose de « pré configurations » d’export de fichier au format de lecture équipementier.