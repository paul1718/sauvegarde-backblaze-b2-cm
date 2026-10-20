---
trigger: always_on
description: Synchronisation automatique obligatoire sur GitHub à chaque modification locale
---

# Règle Locale du Projet : Mise à jour systématique sur GitHub

Pour ce projet (`Sauvegarde Backblaze B2 CM` — `https://github.com/paul1718/sauvegarde-backblaze-b2-cm.git`) :

À chaque création, modification ou suppression de fichier dans le projet :
1. Toujours enregistrer et pousser immédiatement les changements sur GitHub :
   ```powershell
   git add .
   git commit -m "Description claire de la mise à jour"
   git push origin main
   ```
2. Ne jamais terminer une tâche de modification sans avoir vérifié que `git push origin main` a réussi et que le dépôt distant est parfaitement à jour avec le dépôt local.
