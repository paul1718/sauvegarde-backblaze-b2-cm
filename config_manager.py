import json
import os
import sys
from pathlib import Path
from typing import Dict, Any


def get_app_dir() -> Path:
    """Retourne le dossier de l'application (compatible script Python et exécutable PyInstaller)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class ConfigManager:
    """Gère la persistance automatique des identifiants B2 et des chemins de fichiers."""

    DEFAULT_CONFIG: Dict[str, Any] = {
        "b2_key_id": "",
        "b2_app_key": "",
        "bucket_name": "",
        "parent_dir": "",
        "source_excel": "",
        "output_excel": "",
    }

    def __init__(self) -> None:
        self.app_dir = get_app_dir()
        self.local_config_path = self.app_dir / "config.json"

        appdata = os.environ.get("APPDATA")
        if appdata:
            self.appdata_dir = Path(appdata) / "BackblazeB2Uploader"
        else:
            self.appdata_dir = Path.home() / ".backblaze_b2_uploader"
        self.appdata_config_path = self.appdata_dir / "config.json"

    def _get_read_path(self) -> Path:
        if self.local_config_path.exists():
            return self.local_config_path
        if self.appdata_config_path.exists():
            return self.appdata_config_path
        return self.local_config_path

    def load(self) -> Dict[str, Any]:
        """Charge la configuration sauvegardée ou retourne les valeurs par défaut."""
        config = dict(self.DEFAULT_CONFIG)
        read_path = self._get_read_path()
        if read_path.exists():
            try:
                with open(read_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    for key in self.DEFAULT_CONFIG:
                        if key in loaded and loaded[key] is not None:
                            config[key] = str(loaded[key])
            except Exception:
                pass

        # Valeurs par défaut intelligentes pour les fichiers Excel s'ils existent dans le dossier courant
        if not config["source_excel"]:
            default_src = self.app_dir / "dossiers.xlsx"
            if default_src.exists():
                config["source_excel"] = str(default_src)

        if not config["output_excel"]:
            default_out = self.app_dir / "resultats_backblaze.xlsx"
            config["output_excel"] = str(default_out)

        return config

    def save(self, data: Dict[str, Any]) -> None:
        """Sauvegarde la configuration sur le disque (local + AppData en secours)."""
        clean_data = {}
        for key, default_val in self.DEFAULT_CONFIG.items():
            val = data.get(key, default_val)
            clean_data[key] = str(val).strip() if val is not None else ""

        # Sauvegarde locale à côté de l'exécutable
        try:
            with open(self.local_config_path, "w", encoding="utf-8") as f:
                json.dump(clean_data, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

        # Sauvegarde dans AppData pour garantir la persistance même si le dossier de l'exe est restreint
        try:
            self.appdata_dir.mkdir(parents=True, exist_ok=True)
            with open(self.appdata_config_path, "w", encoding="utf-8") as f:
                json.dump(clean_data, f, indent=4, ensure_ascii=False)
        except Exception:
            pass
