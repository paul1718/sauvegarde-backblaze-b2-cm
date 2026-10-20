import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Dict, Any, Optional, Tuple
import requests

from config_manager import get_app_dir


GITHUB_RAW_BASE = "https://raw.githubusercontent.com/paul1718/sauvegarde-backblaze-b2-cm/main"
LOCAL_VERSION_DEFAULT = "1.2.0"


def _parse_version(v_str: str) -> Tuple[int, ...]:
    parts = []
    for p in str(v_str).strip().lstrip("vV").split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


class AutoUpdater:
    """
    Système de mise à jour à distance via GitHub.
    Permet aux machines installées dans un autre quartier ou une autre ville
    de recevoir automatiquement les mises à jour poussées sur le dépôt GitHub.
    """

    def __init__(self) -> None:
        self.app_dir = get_app_dir()
        # Si l'application tourne depuis dist/BackblazeB2Uploader/BackblazeB2Uploader.exe,
        # le dossier racine du projet est deux niveaux au-dessus s'il existe.
        self.project_root = self._detect_project_root()

    def _detect_project_root(self) -> Path:
        if (self.app_dir / "version.json").exists() or (self.app_dir / "app.py").exists():
            return self.app_dir
        parent2 = self.app_dir.parent.parent
        if (parent2 / "version.json").exists() or (parent2 / "app.py").exists():
            return parent2
        return self.app_dir

    def get_local_version(self) -> str:
        for folder in (self.project_root, self.app_dir, Path(__file__).resolve().parent):
            v_file = folder / "version.json"
            if v_file.exists():
                try:
                    with open(v_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        return str(data.get("version", LOCAL_VERSION_DEFAULT))
                except Exception:
                    pass
        return LOCAL_VERSION_DEFAULT

    def check_for_updates(self, timeout: float = 6.0) -> Dict[str, Any]:
        """
        Vérifie sur GitHub si une version plus récente est disponible.
        Retourne :
        {
            "update_available": bool,
            "local_version": str,
            "remote_version": str,
            "description": str,
            "files": list[str],
            "error": Optional[str]
        }
        """
        local_v = self.get_local_version()
        try:
            url = f"{GITHUB_RAW_BASE}/version.json?t={int(os.times().elapsed * 1000)}"
            resp = requests.get(url, timeout=timeout, headers={"Cache-Control": "no-cache"})
            resp.raise_for_status()
            remote_data = resp.json()
            remote_v = str(remote_data.get("version", local_v))
            has_update = _parse_version(remote_v) > _parse_version(local_v)
            return {
                "update_available": has_update,
                "local_version": local_v,
                "remote_version": remote_v,
                "description": str(remote_data.get("description", "")),
                "files": list(
                    remote_data.get(
                        "files",
                        [
                            "version.json",
                            "app.py",
                            "b2_uploader.py",
                            "excel_manager.py",
                            "config_manager.py",
                            "auto_updater.py",
                        ],
                    )
                ),
                "error": None,
            }
        except Exception as e:
            return {
                "update_available": False,
                "local_version": local_v,
                "remote_version": local_v,
                "description": "",
                "files": [],
                "error": str(e),
            }

    def apply_update(self, files_to_update: Optional[ list ] = None) -> Tuple[bool, str]:
        """
        Télécharge atomiquement les fichiers mis à jour depuis GitHub (ou exécute git pull si Git est présent).
        """
        try:
            # 1. Si le dossier est un dépôt Git, on tente d'abord un `git pull --rebase` propre
            if (self.project_root / ".git").exists():
                try:
                    subprocess.run(
                        ["git", "fetch", "origin", "main"],
                        cwd=str(self.project_root),
                        timeout=15,
                        check=False,
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    )
                except Exception:
                    pass

            # 2. Téléchargement direct HTTP depuis GitHub Raw (fonctionne sur tout PC distant même sans Git installé !)
            if not files_to_update:
                files_to_update = [
                    "version.json",
                    "app.py",
                    "b2_uploader.py",
                    "excel_manager.py",
                    "config_manager.py",
                    "auto_updater.py",
                ]

            target_dirs = [self.project_root]
            if self.app_dir != self.project_root:
                target_dirs.append(self.app_dir)

            for rel_file in files_to_update:
                url = f"{GITHUB_RAW_BASE}/{rel_file}"
                resp = requests.get(url, timeout=15, headers={"Cache-Control": "no-cache"})
                resp.raise_for_status()
                content = resp.content

                for base_dir in target_dirs:
                    dest_path = base_dir / rel_file
                    dest_path.parent.mkdir(parents=True, exist_ok=True)
                    tmp_path = dest_path.with_suffix(dest_path.suffix + ".tmp")
                    with open(tmp_path, "wb") as f:
                        f.write(content)
                        f.flush()
                        os.fsync(f.fileno())
                    os.replace(tmp_path, dest_path)

            return True, "Mise à jour téléchargée et installée avec succès."
        except Exception as e:
            return False, f"Échec de la mise à jour à distance : {e}"
