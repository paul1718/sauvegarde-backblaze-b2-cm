import json
import os
import shutil
import sys
from pathlib import Path
from typing import Dict, Any, Optional


def get_app_dir() -> Path:
    """Retourne le dossier de l'application (compatible script Python et exécutable PyInstaller)."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        # Si exécuté depuis dist/BackblazeB2Uploader/BackblazeB2Uploader.exe, remonter à la racine du projet si présente
        if exe_dir.name == "BackblazeB2Uploader" and exe_dir.parent.name == "dist":
            proj_root = exe_dir.parent.parent
            if (proj_root / "app.py").exists() or (proj_root / "config.json").exists():
                return proj_root
        return exe_dir
    return Path(__file__).resolve().parent


def atomic_write_json(path: Path, data: Dict[str, Any]) -> None:
    """
    Écrit un fichier JSON de manière 100% atomique (anti-corruption en cas de coupure de courant).
    1. Sauvegarde l'ancien fichier valide en `.bak`
    2. Écrit dans `.tmp` + `flush()` + `os.fsync()`
    3. Remplace atomiquement via `os.replace()`
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    bak_path = path.with_suffix(path.suffix + ".bak")

    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())

    if path.exists():
        try:
            shutil.copy2(path, bak_path)
        except Exception:
            pass

    os.replace(tmp_path, path)


def safe_read_json(path: Path) -> Optional[Dict[str, Any]]:
    """
    Lit un fichier JSON avec bascule automatique sur la sauvegarde `.bak`
    si le fichier principal a subi un dommage matériel.
    """
    bak_path = path.with_suffix(path.suffix + ".bak")
    for candidate in (path, bak_path):
        if candidate.exists():
            try:
                with open(candidate, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    if candidate == bak_path:
                        try:
                            shutil.copy2(bak_path, path)
                        except Exception:
                            pass
                    return loaded
            except Exception:
                continue
    return None


class ConfigManager:
    """Gère la persistance atomique des identifiants B2, des chemins et de l'état de reprise après coupure."""

    DEFAULT_CONFIG: Dict[str, Any] = {
        "b2_key_id": "",
        "b2_app_key": "",
        "bucket_name": "",
        "parent_dir": "",
        "source_excel": "",
        "output_excel": "",
        "auto_resume_after_power_loss": True,
    }

    def __init__(self) -> None:
        self.app_dir = get_app_dir()
        self.local_config_path = self.app_dir / "config.json"
        self.local_state_path = self.app_dir / "state_journal.json"

        appdata = os.environ.get("APPDATA")
        if appdata:
            self.appdata_dir = Path(appdata) / "BackblazeB2Uploader"
        else:
            self.appdata_dir = Path.home() / ".backblaze_b2_uploader"
        self.appdata_config_path = self.appdata_dir / "config.json"
        self.appdata_state_path = self.appdata_dir / "state_journal.json"

    def load(self) -> Dict[str, Any]:
        """Charge la configuration sauvegardée ou retourne les valeurs par défaut."""
        config = dict(self.DEFAULT_CONFIG)
        loaded = safe_read_json(self.local_config_path) or safe_read_json(self.appdata_config_path)
        if loaded:
            for key in self.DEFAULT_CONFIG:
                if key in loaded and loaded[key] is not None:
                    if isinstance(self.DEFAULT_CONFIG[key], bool):
                        config[key] = bool(loaded[key])
                    else:
                        config[key] = str(loaded[key])

        if not config["source_excel"]:
            default_src = self.app_dir / "dossiers.xlsx"
            if default_src.exists():
                config["source_excel"] = str(default_src)

        if not config["output_excel"]:
            default_out = self.app_dir / "resultats_backblaze.xlsx"
            config["output_excel"] = str(default_out)

        return config

    def save(self, data: Dict[str, Any]) -> None:
        """Sauvegarde atomique de la configuration sur le disque (local + AppData)."""
        clean_data: Dict[str, Any] = {}
        for key, default_val in self.DEFAULT_CONFIG.items():
            val = data.get(key, default_val)
            if isinstance(default_val, bool):
                clean_data[key] = bool(val)
            else:
                clean_data[key] = str(val).strip() if val is not None else ""

        for target in (self.local_config_path, self.appdata_config_path):
            try:
                atomic_write_json(target, clean_data)
            except Exception:
                pass

    # =========================================================================
    # Journal d'état anti-coupure de courant (state_journal.json)
    # =========================================================================
    def load_state_journal(self) -> Dict[str, Any]:
        """
        Charge le journal d'état permettant de reprendre exactement au fichier près
        après une coupure brusque d'électricité.
        """
        default_state: Dict[str, Any] = {
            "active_session": False,
            "current_folder": "",
            "uploaded_files": {},  # b2_key -> {"size": int, "url": str}
        }
        loaded = safe_read_json(self.local_state_path) or safe_read_json(self.appdata_state_path)
        if loaded:
            default_state["active_session"] = bool(loaded.get("active_session", False))
            default_state["current_folder"] = str(loaded.get("current_folder", ""))
            if isinstance(loaded.get("uploaded_files"), dict):
                default_state["uploaded_files"] = loaded["uploaded_files"]
        return default_state

    def save_state_journal(self, state: Dict[str, Any]) -> None:
        """Enregistre atomiquement le journal d'état sur le disque."""
        for target in (self.local_state_path, self.appdata_state_path):
            try:
                atomic_write_json(target, state)
            except Exception:
                pass

    def mark_session_active(self, active: bool, current_folder: str = "") -> None:
        state = self.load_state_journal()
        state["active_session"] = active
        if current_folder or not active:
            state["current_folder"] = current_folder
        self.save_state_journal(state)

    def record_uploaded_file(self, b2_key: str, file_size: int, url: str) -> None:
        state = self.load_state_journal()
        uploaded = state.get("uploaded_files", {})
        uploaded[b2_key] = {"size": int(file_size), "url": str(url)}
        state["uploaded_files"] = uploaded
        self.save_state_journal(state)
