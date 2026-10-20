# Cahier des Charges Technico-Fonctionnel : Application Desktop de Sauvegarde Automatisée Backblaze B2

## 1. Contexte & Objectif

L'objectif est de développer une application desktop avec interface graphique (GUI) destinée à un profil non-technique (secrétariat). L'application doit permettre le téléversement massif et sans limite de taille (de 90 Go à plus de 1 To) de dossiers de médias vers un compte Backblaze B2, la mise à jour dynamique d'un fichier Excel source, et la création d'un fichier Excel de sortie structuré avec les liens publics générés.

## 2. Interface Utilisateur (GUI) & Ergonomie

L'interface doit être simple, claire et intuitive pour un utilisateur non-technique :

1. **Section Configuration / Identifiants Backblaze B2 :**
   * Champ `B2 Key ID`
   * Champ `B2 Application Key`
   * Champ `Bucket Name`
   * *Mémorisation automatique des identifiants.*

2. **Section Sélection des Fichiers et Répertoires :**
   * **Répertoire Parent :** Champ permettant de saisir ou de choisir le répertoire parent. Le dernier chemin sélectionné doit être conservé par défaut pour les sessions futures, mais reste modifiable à chaque nouvelle sauvegarde.
   * **Fichier Excel Source (`dossiers.xlsx`) :** Bouton de sélection du fichier source (pouvant varier à chaque sauvegarde).
   * **Fichier Excel de Sortie (`resultats_backblaze.xlsx`) :** Bouton de sélection de l'emplacement du fichier final.

3. **Section Contrôle & Progression :**
   * Boutons : **"Lancer la Sauvegarde"** / **"Pause"** / **"Reprendre"**.
   * **Gestion de l'interruption :** L'utilisateur peut interrompre le traitement à tout moment. Au relancement, l'application reprend automatiquement le travail en cours.
   * Barre de progression globale et journal d'exécution (*logs*) lisible affichant l'avancement en temps réel.

## 3. Workflow & Règles Métier

### A. Fichier Excel Source (`dossiers.xlsx`) & Reprise Automatique

* Contient la liste des noms de dossiers à traiter (Colonne `Dossier`).
* Le chemin complet de chaque dossier est calculé dynamiquement : `Chemin_Parent_Actuel + Nom_Dossier`.
* Contient une colonne `Statut` :
  * Lors de l'exécution, l'application **vérifie la colonne `Statut`** pour chaque ligne.
  * Si la cellule contient la valeur **`Terminé`**, le dossier est **ignoré / sauté**.
  * Seuls les dossiers qui ne portent pas la mention `Terminé` sont traités.
  * Dès qu'un dossier a fini d'être intégralement téléversé et consigné dans l'Excel de sortie, l'application écrit immédiatement **`Terminé`** dans sa ligne correspondante du fichier source.

### B. Téléversement vers Backblaze B2

1. Pour chaque dossier éligible :
   * Parcourir le répertoire principal et **tous ses sous-dossiers**.
   * Téléverser **TOUS** les fichiers compatibles sur Backblaze B2 en conservant l'arborescence :
     * **Images :** `.png`, `.jpg`, `.jpeg`, `.gif`, `.webp`
     * **Vidéos :** `.mp4`, `.avi`, `.mov`, `.mkv`, `.webm`
     * **Fichiers Texte / Documents :** `.txt`, `.pdf`, `.doc`, `.docx`, `.csv`, `.rtf`
   * L'application doit gérer les erreurs réseau et les micro-coupures sans planter (politique de *retry* automatique), indispensable pour des volumes dépassant 1 To.

### C. Gestion des Liens pour l'Excel de Sortie (`resultats_backblaze.xlsx`)

1. **Filtrage pour l'Excel de sortie :**
   * Seuls les fichiers situés **directement à la racine du dossier** (et non dans ses sous-dossiers) sont ajoutés au fichier Excel.
   * **Fichiers Texte / Documents :** Téléversés sur Backblaze B2, mais leurs liens **ne doivent pas** figurer dans le fichier Excel final.

2. **Colonnes de l'Excel de sortie :**
   * `Dossier` : Nom du dossier.
   * `Image_1` à `Image_10` : 10 colonnes. Un lien d'image par cellule (limité aux **10 premières images** de la racine).
   * `Video_1` à `Video_5` : 5 colonnes. Un lien de vidéo par cellule (limité aux **5 premières vidéos** de la racine).

## 4. Structure des Fichiers Excel

### Fichier Source (`dossiers.xlsx`)

| Dossier | Statut |
| :--- | :--- |
| `Mariage_Jean` | Terminé |
| `Terrain_Mendong` | *(vide ou en cours)* |
| `Vacances_2026` | *(vide)* |

### Fichier de Sortie (`resultats_backblaze.xlsx`)

| Dossier | Image_1 | Image_2 | ... | Image_10 | Video_1 | ... | Video_5 |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `Mariage_Jean` | `https://.../img1.jpg` | `https://.../img2.jpg` | ... | *(vide si < 10)* | `https://.../vid1.mp4` | ... | *(vide si < 5)* |

## 5. Spécifications Techniques

* **Langage & GUI :** Python 3.10+ avec `PyQt6` ou `CustomTkinter`.
* **Packaging :** Compilation sous forme d'application exécutable autonome (`.exe` pour Windows) via `PyInstaller`, permettant une utilisation par un profil non-technique sans dépendances Python.
* **Bibliothèques principales :** `b2sdk`, `pandas`, `openpyxl`.