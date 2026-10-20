# Application Desktop de Sauvegarde Automatisée Backblaze B2

Application graphique (GUI) destinée au téléversement massif (90 Go à > 1 To) de dossiers médias vers un bucket **Backblaze B2**, avec reprise automatique sur interruption et génération structurée des liens publics dans Excel.

---

## Fonctionnalités Clés

1. **Configuration & Mémorisation Automatique**
   - Saisie des identifiants `B2 Key ID`, `B2 Application Key` et `Bucket Name`.
   - Sauvegarde automatique dans `config.json` (et `%APPDATA%\BackblazeB2Uploader\config.json`).
   - Bouton **Tester la connexion** pour vérifier immédiatement l'accès au bucket B2.

2. **Gestion des Fichiers Excel & Reprise Automatique**
   - **Répertoire Parent** mémorisé d'une session à l'autre et modifiable à tout moment.
   - **Excel Source (`dossiers.xlsx`)** :
     - Lit la colonne `Dossier` et la colonne `Statut`.
     - Ignore automatiquement tous les dossiers dont le `Statut` est **`Terminé`**.
     - Met à jour immédiatement la ligne avec **`Terminé`** dès qu'un dossier est intégralement téléversé et consigné.
     - Bouton **Créer modèle** permettant de générer automatiquement `dossiers.xlsx` à partir des sous-dossiers présents dans le Répertoire Parent.
   - **Excel de Sortie (`resultats_backblaze.xlsx`)** :
     - Contient les 16 colonnes réglementaires : `Dossier`, `Image_1` à `Image_10`, `Video_1` à `Video_5`.
     - Consigne uniquement les fichiers situés **directement à la racine** de chaque dossier (jusqu'aux 10 premières images et 5 premières vidéos).
     - Les documents texte/PDF (`.txt`, `.pdf`, `.doc`, `.docx`, `.csv`, `.rtf`) et les fichiers des sous-dossiers sont bien téléversés sur Backblaze B2 en conservant leur arborescence, mais leurs liens ne sont pas ajoutés dans l'Excel de sortie.

3. **Résilience Réseau (> 1 To) & Contrôle Temps Réel**
   - Boutons **Lancer la Sauvegarde**, **Pause**, **Reprendre** et **Interrompre**.
   - Politique de *retry* automatique avec *backoff* exponentiel en cas de micro-coupure Internet.
   - Vérification intelligente des fichiers déjà présents sur B2 (même chemin et même taille) pour ne pas re-téléverser inutilement les fichiers d'un dossier interrompu en cours de route.

---

## Lancement de l'Application

### Option 1 : Double-cliquer sur `Lancer_Application.bat` ou l'exécutable `.exe`
- Double-cliquez sur `Lancer_Application.bat` ou sur `dist\BackblazeB2Uploader\BackblazeB2Uploader.exe`.

### Option 2 : En ligne de commande Python
```bash
pip install -r requirements.txt
python app.py
```

### Option 3 : Recompiler l'exécutable `.exe` (PyInstaller)
```bash
python build_exe.py
```
