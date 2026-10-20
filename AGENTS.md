# Règles Locales du Projet — Sauvegarde Backblaze B2 CM

## Synchronisation Automatique et Systématique sur GitHub

À chaque mise à jour, création ou modification d'un fichier dans ce dépôt local (`d:\Automatisation\Sauvegarde Backblaze B2 CM`) :

1. **Toujours** effectuer immédiatement la mise à jour sur GitHub (`https://github.com/paul1718/sauvegarde-backblaze-b2-cm.git`) :
   ```powershell
   git add .
   git commit -m "Description de la mise à jour"
   git push origin main
   ```
2. Un hook Git local (`.git/hooks/post-commit`) est également actif pour pousser automatiquement chaque nouveau commit vers `origin main`.
