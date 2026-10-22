import os
import re
import socket
import time
import threading
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, Set, Any
from urllib.parse import quote, urlsplit, urlunsplit

from b2sdk.v2 import InMemoryAccountInfo, B2Api
from b2sdk.v2.exception import B2Error


IMAGE_EXTENSIONS: Set[str] = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
VIDEO_EXTENSIONS: Set[str] = {".mp4", ".avi", ".mov", ".mkv", ".webm"}
DOC_EXTENSIONS: Set[str] = {".txt", ".pdf", ".doc", ".docx", ".csv", ".rtf"}
ALL_ALLOWED_EXTENSIONS: Set[str] = IMAGE_EXTENSIONS | VIDEO_EXTENSIONS | DOC_EXTENSIONS


def natural_sort_key(path: Path) -> List[object]:
    """Clé de tri naturel (ex: img2.jpg avant img10.jpg)."""
    return [
        int(text) if text.isdigit() else text.lower()
        for text in re.split(r"(\d+)", path.name)
    ]


def format_size(num_bytes: float) -> str:
    """Formate une taille en octets de manière lisible (Ko, Mo, Go, To)."""
    for unit in ["o", "Ko", "Mo", "Go", "To"]:
        if abs(num_bytes) < 1024.0:
            if unit == "o":
                return f"{int(num_bytes)} {unit}"
            return f"{num_bytes:.2f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.2f} Po"


def normalize_b2_url(raw_url: str) -> str:
    """S'assure que l'URL publique Backblaze B2 est proprement encodée (espaces, accents)."""
    parts = urlsplit(raw_url)
    if "%" not in parts.path:
        encoded_path = "/".join(quote(seg, safe="") for seg in parts.path.split("/"))
    else:
        encoded_path = parts.path
    return urlunsplit((parts.scheme, parts.netloc, encoded_path, parts.query, parts.fragment))


def is_tav_folder(excel_dossier_name: str, folder_path: Optional[Path] = None) -> bool:
    """
    Vérifie si le nom du dossier commence par 'TAV' (insensible à la casse).
    """
    if excel_dossier_name.strip().upper().startswith("TAV"):
        return True
    if folder_path is not None and folder_path.name.strip().upper().startswith("TAV"):
        return True
    return False


def resolve_matching_folder(parent_dir: Path, excel_dossier_name: str) -> Optional[Path]:
    """
    Recherche dans `parent_dir` le dossier correspondant à `excel_dossier_name` :
    1. Correspondance exacte directe (`parent_dir / excel_dossier_name`)
    2. Correspondance exacte insensible à la casse (ex: 'tv326' == 'TV326')
    3. Correspondance par préfixe : le nom du dossier sur le disque commence par le nom
       du dossier Excel suivi OBLIGATOIREMENT d'un espace (ex: 'TV326 - ...' ou 'TV326 ...').
       Il y a toujours un espace après le préfixe, ce qui évite formellement d'associer
       'TV326' avec 'TV3260' ou 'TV3261'.
    """
    clean_name = excel_dossier_name.strip()
    if not clean_name or not parent_dir.exists() or not parent_dir.is_dir():
        return None

    # 1. Correspondance directe exacte
    exact_path = parent_dir / clean_name
    if exact_path.exists() and exact_path.is_dir():
        return exact_path

    # 2. Liste triée des sous-dossiers
    prefix_lower = clean_name.lower()
    try:
        subdirs = [d for d in parent_dir.iterdir() if d.is_dir() and not d.name.startswith(".")]
    except Exception:
        return None

    subdirs.sort(key=natural_sort_key)

    # 3. Recherche de correspondance exacte insensible à la casse
    for d in subdirs:
        if d.name.strip().lower() == prefix_lower:
            return d

    # 4. Recherche par préfixe : il y a TOUJOURS un espace après le préfixe
    # (ex: 'tv326 ' permet d'associer 'TV326 - ...', mais JAMAIS 'TV3260')
    prefix_with_space = prefix_lower + " "
    for d in subdirs:
        if d.name.strip().lower().startswith(prefix_with_space):
            return d

    return None


def is_internet_connected(timeout: float = 3.0) -> bool:
    """
    Vérifie si la connexion Internet vers Backblaze B2 est active.
    """
    targets = [
        ("api.backblazeb2.com", 443),
        ("1.1.1.1", 443),
    ]
    for host, port in targets:
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except OSError:
            continue
    return False


class UploadCancelledError(Exception):
    """Exception interne levée lorsque l'utilisateur demande l'arrêt."""
    pass


class B2UploaderService:
    """
    Service gérant la connexion à Backblaze B2, la résolution des dossiers par préfixe,
    la gestion de l'image standard pour les dossiers 'TAV' vides, la comparaison/mise à jour
    intelligente des dossiers déjà sauvegardés, et la pause/reprise automatique sur coupure Internet.
    """

    def __init__(
        self,
        key_id: str,
        app_key: str,
        bucket_name: str,
        log_callback: Optional[Callable[[str, str], None]] = None,
        status_callback: Optional[Callable[[str], None]] = None,
        pause_event: Optional[threading.Event] = None,
        stop_event: Optional[threading.Event] = None,
    ) -> None:
        self.key_id = key_id.strip()
        self.app_key = app_key.strip()
        self.bucket_name = bucket_name.strip()
        self.log_callback = log_callback
        self.status_callback = status_callback
        self.pause_event = pause_event or threading.Event()
        if pause_event is None:
            self.pause_event.set()
        self.stop_event = stop_event or threading.Event()

        self.info = InMemoryAccountInfo()
        self.b2_api = B2Api(self.info)
        self.bucket = None

    def _log(self, message: str, level: str = "INFO") -> None:
        if self.log_callback:
            self.log_callback(message, level)

    def _status(self, text: str) -> None:
        if self.status_callback:
            self.status_callback(text)

    def _check_pause_and_stop(self) -> None:
        if self.stop_event.is_set():
            raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")
        if not self.pause_event.is_set():
            self._log("Sauvegarde en pause manuelle. En attente de reprise...", "WARN")
            while not self.pause_event.is_set():
                if self.stop_event.is_set():
                    raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")
                time.sleep(0.25)
            self._log("Reprise de la sauvegarde...", "INFO")

    def _wait_for_internet_if_down(self, context_msg: str = "") -> bool:
        """
        Si la connexion Internet est coupée, met automatiquement la sauvegarde en pause
        et attend que la connexion revienne.
        """
        self._check_pause_and_stop()
        if is_internet_connected():
            return False

        prefix = f"({context_msg}) " if context_msg else ""
        self._log(
            f"📡 {prefix}Connexion Internet coupée. Mise en pause automatique en attente du retour d'Internet...",
            "WARN",
        )
        self._status("📡 Connexion Internet coupée — En pause automatique (reprise dès le retour du réseau)...")

        waited_seconds = 0
        while not is_internet_connected():
            self._check_pause_and_stop()
            self._sleep_interruptible(5.0)
            waited_seconds += 5
            if waited_seconds % 60 == 0:
                mins = waited_seconds // 60
                self._status(
                    f"📡 En attente de la connexion Internet ({mins} min écoulées)... Reprise automatique dès le retour."
                )

        self._log("🌐 Connexion Internet rétablie ! Reconnexion à Backblaze B2 et reprise automatique...", "SUCCESS")
        self._status("🌐 Connexion Internet rétablie ! Reprise du téléversement...")
        return True

    def connect(self, max_auth_retries: int = 5) -> None:
        last_error: Optional[Exception] = None
        attempt = 0
        while attempt < max_auth_retries:
            self._check_pause_and_stop()
            self._wait_for_internet_if_down("Connexion B2")

            attempt += 1
            try:
                self.info = InMemoryAccountInfo()
                self.b2_api = B2Api(self.info)
                self.b2_api.authorize_account("production", self.key_id, self.app_key)
                self.bucket = self.b2_api.get_bucket_by_name(self.bucket_name)
                return
            except UploadCancelledError:
                raise
            except Exception as e:
                last_error = e
                err_str = str(e).lower()
                is_auth_rejection = any(
                    tok in err_str
                    for tok in ("unauthorized", "bad_auth_token", "invalid_account_id", "unsupported")
                )
                if not is_internet_connected() or (max_auth_retries > 1 and not is_auth_rejection):
                    if not is_internet_connected():
                        attempt = max(0, attempt - 1)
                        continue
                if attempt < max_auth_retries:
                    wait_sec = min(2 ** min(attempt, 4), 20)
                    self._log(
                        f"Tentative de connexion B2 ({attempt}/{max_auth_retries}) échouée : {e}. Nouvel essai dans {wait_sec}s...",
                        "WARN",
                    )
                    self._sleep_interruptible(wait_sec)
        raise ConnectionError(f"Impossible de se connecter à Backblaze B2 : {last_error}")

    def _sleep_interruptible(self, seconds: float) -> None:
        end_time = time.time() + seconds
        while time.time() < end_time:
            if self.stop_event.is_set():
                raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")
            time.sleep(0.2)

    @staticmethod
    def scan_folder_files(folder_path: Path) -> List[Tuple[Path, str, bool]]:
        """
        Parcourt récursivement `folder_path` et tous ses sous-dossiers.
        Retourne une liste triée de tuples :
        (chemin_local_complet, chemin_relatif_b2, est_a_la_racine)
        Uniquement pour les fichiers ayant une extension autorisée (Images, Vidéos, Documents).
        """
        eligible_files: List[Tuple[Path, str, bool]] = []
        folder_name = folder_path.name

        for root, dirs, files in os.walk(folder_path):
            dirs.sort(key=lambda d: d.lower())
            root_path = Path(root)
            is_root = root_path.resolve() == folder_path.resolve()

            file_paths = [root_path / f for f in files]
            file_paths.sort(key=natural_sort_key)

            for file_path in file_paths:
                ext = file_path.suffix.lower()
                if ext in ALL_ALLOWED_EXTENSIONS:
                    rel_inside_folder = file_path.relative_to(folder_path).as_posix()
                    b2_key = f"{folder_name}/{rel_inside_folder}"
                    eligible_files.append((file_path, b2_key, is_root))

        return eligible_files

    def _fetch_existing_b2_files_detailed(self, folder_name: str) -> Dict[str, Dict[str, Any]]:
        """
        Récupère la liste détaillée des fichiers déjà présents sur B2 dans `folder_name/` :
        { b2_key: {"size": int, "file_id": str} }
        """
        existing: Dict[str, Dict[str, Any]] = {}
        if not self.bucket:
            return existing
        try:
            prefix = f"{folder_name}/"
            for file_version, _ in self.bucket.ls(folder_to_list=prefix, recursive=True):
                existing[file_version.file_name] = {
                    "size": int(file_version.size),
                    "file_id": str(file_version.id_),
                }
        except Exception:
            pass
        return existing

    def ensure_tav_standard_image_url(self, image_path_or_url: str) -> str:
        """
        Garantit l'obtention d'un lien public unique et permanent pour l'image standard
        destinée aux dossiers vides commençant par 'TAV' :
        - Si `image_path_or_url` est déjà une URL HTTP(S), la retourne directement.
        - Si c'est un fichier image local, vérifie s'il est déjà dans `_Image_Standard_TAV/` sur B2
          (même taille) ; sinon le téléverse une seule fois pour toutes et retourne son URL publique.
        """
        val = image_path_or_url.strip()
        if not val:
            raise ValueError(
                "Aucune image standard TAV n'est configurée. Veuillez sélectionner l'image standard pour les dossiers TAV vides."
            )

        if val.lower().startswith(("http://", "https://")):
            return val

        local_img = Path(val)
        if not local_img.exists() or not local_img.is_file():
            raise FileNotFoundError(f"Le fichier d'image standard TAV est introuvable : {val}")

        b2_key = f"_Image_Standard_TAV/{local_img.name}"
        existing_std = self._fetch_existing_b2_files_detailed("_Image_Standard_TAV")
        local_size = local_img.stat().st_size

        if b2_key in existing_std and existing_std[b2_key]["size"] == local_size:
            raw_url = self.bucket.get_download_url(b2_key)
            return normalize_b2_url(raw_url)

        self._log(f"Téléversement initial de l'image standard TAV ('{local_img.name}') sur Backblaze B2...", "INFO")
        url, _ = self._upload_single_file_resilient(
            local_path=local_img,
            b2_key=b2_key,
            existing_b2_files={},
            local_journal_files={},
        )
        self._log(f"✔ Image standard TAV enregistrée sur B2 : {url}", "SUCCESS")
        return url

    def _upload_single_file_resilient(
        self,
        local_path: Path,
        b2_key: str,
        existing_b2_files: Dict[str, Dict[str, Any]],
        local_journal_files: Dict[str, Any],
        max_consecutive_server_errors: int = 12,
    ) -> Tuple[str, bool]:
        """
        Téléverse un fichier sur Backblaze B2 avec résilience totale aux coupures Internet.
        Retourne (url_publique, etait_deja_present_et_identique).
        """
        self._check_pause_and_stop()
        local_size = local_path.stat().st_size

        # Vérification directe sur le bucket Backblaze B2 (taille identique)
        if b2_key in existing_b2_files and existing_b2_files[b2_key]["size"] == local_size:
            raw_url = self.bucket.get_download_url(b2_key)
            return normalize_b2_url(raw_url), True

        consecutive_errors = 0
        last_err: Optional[Exception] = None

        while consecutive_errors < max_consecutive_server_errors:
            self._check_pause_and_stop()
            if self._wait_for_internet_if_down(local_path.name):
                try:
                    self.connect(max_auth_retries=5)
                except UploadCancelledError:
                    raise
                except Exception:
                    pass

            try:
                if self.bucket is None:
                    self.connect(max_auth_retries=5)

                self.bucket.upload_local_file(
                    local_file=str(local_path),
                    file_name=b2_key,
                )
                raw_url = self.bucket.get_download_url(b2_key)
                return normalize_b2_url(raw_url), False

            except UploadCancelledError:
                raise
            except (B2Error, OSError, Exception) as e:
                last_err = e
                if not is_internet_connected():
                    self._wait_for_internet_if_down(local_path.name)
                    try:
                        self.connect(max_auth_retries=5)
                    except UploadCancelledError:
                        raise
                    except Exception:
                        pass
                    continue

                consecutive_errors += 1
                wait_sec = min(3 * (2 ** min(consecutive_errors - 1, 4)), 45)
                self._log(
                    f"Micro-coupure ou erreur temporaire sur '{local_path.name}' "
                    f"(essai {consecutive_errors}/{max_consecutive_server_errors}) : {e}. "
                    f"Nouvelle tentative dans {wait_sec}s...",
                    "WARN",
                )
                self._sleep_interruptible(wait_sec)
                try:
                    self.connect(max_auth_retries=3)
                except UploadCancelledError:
                    raise
                except Exception:
                    pass

        raise RuntimeError(
            f"Échec du téléversement de '{local_path.name}' après {max_consecutive_server_errors} essais : {last_err}"
        )

    def upload_folder(
        self,
        folder_path: Path,
        local_journal_files: Optional[Dict[str, Any]] = None,
        on_file_uploaded: Optional[Callable[[int, Path, str, str], None]] = None,
        on_root_links_updated: Optional[Callable[[List[str], List[str], int, int], None]] = None,
    ) -> Tuple[List[str], List[str], int, int, str]:
        """
        Compare le dossier source local `folder_path` avec le dossier sauvegardé sur B2 :
        - Si le dossier est déjà sauvegardé et 100% identique sur B2 : ne fait aucun téléversement,
          mais génère et retourne les liens des images et vidéos racine pour l'Excel final (`sync_state = "identical"`).
        - Si le dossier a changé (nouveaux fichiers, fichiers modifiés ou fichiers supprimés localement) :
          met à jour le dossier sauvegardé sur B2 et retourne les liens pour l'Excel final (`sync_state = "updated"`).
        - Si c'est un nouveau dossier sur B2 : téléverse tout et retourne les liens (`sync_state = "created"`).
        """
        self._check_pause_and_stop()
        if not folder_path.exists() or not folder_path.is_dir():
            raise FileNotFoundError(f"Le dossier est introuvable : {folder_path}")

        eligible_files = self.scan_folder_files(folder_path)
        existing_b2 = self._fetch_existing_b2_files_detailed(folder_path.name)
        journal_map = local_journal_files if isinstance(local_journal_files, dict) else {}

        was_already_on_b2 = len(existing_b2) > 0
        local_keys_sizes: Dict[str, int] = {
            b2_key: local_path.stat().st_size for local_path, b2_key, _ in eligible_files
        }

        # Détection des différences entre le dossier source local et le dossier sauvegardé sur B2
        changed_or_new_keys = [
            k for k, sz in local_keys_sizes.items()
            if k not in existing_b2 or existing_b2[k]["size"] != sz
        ]
        obsolete_b2_keys = [
            k for k in existing_b2.keys()
            if k not in local_keys_sizes
        ]

        is_identical_to_saved = was_already_on_b2 and len(changed_or_new_keys) == 0 and len(obsolete_b2_keys) == 0

        if is_identical_to_saved:
            self._log(
                f"  ✓ Dossier '{folder_path.name}' déjà sauvegardé et 100% identique sur B2 : aucun téléversement nécessaire.",
                "INFO",
            )
            sync_state = "identical"
        elif was_already_on_b2:
            self._log(
                f"  🔄 Dossier '{folder_path.name}' modifié depuis la dernière sauvegarde "
                f"({len(changed_or_new_keys)} fichier(s) à ajouter/mettre à jour, {len(obsolete_b2_keys)} fichier(s) obsolète(s)) : "
                "mise à jour du dossier sur B2...",
                "INFO",
            )
            sync_state = "updated"
            # Suppression des fichiers obsolètes sur B2 pour que le dossier sauvegardé reflète exactement la source
            for obs_key in obsolete_b2_keys:
                self._check_pause_and_stop()
                try:
                    file_id = existing_b2[obs_key]["file_id"]
                    self.bucket.delete_file_version(file_id, obs_key)
                    self._log(f"  🗑 Fichier supprimé du dossier B2 (absent de la source) : {obs_key}", "INFO")
                except Exception:
                    pass
        else:
            sync_state = "created"

        root_image_links: List[str] = []
        root_video_links: List[str] = []
        total_files = len(eligible_files)
        total_bytes = 0

        for idx, (local_path, b2_key, is_root) in enumerate(eligible_files, start=1):
            self._check_pause_and_stop()
            file_size = local_path.stat().st_size
            ext = local_path.suffix.lower()

            url, already_present = self._upload_single_file_resilient(
                local_path=local_path,
                b2_key=b2_key,
                existing_b2_files=existing_b2,
                local_journal_files=journal_map,
            )
            total_bytes += file_size

            if not is_identical_to_saved:
                if already_present:
                    self._log(
                        f"  ✓ [{idx}/{total_files}] Déjà à jour sur B2 : {b2_key} ({format_size(file_size)})",
                        "INFO",
                    )
                else:
                    self._log(
                        f"  ↑ [{idx}/{total_files}] Téléversé / Mis à jour : {b2_key} ({format_size(file_size)})",
                        "SUCCESS",
                    )

            root_updated = False
            if is_root:
                if ext in IMAGE_EXTENSIONS and len(root_image_links) < 10:
                    root_image_links.append(url)
                    root_updated = True
                elif ext in VIDEO_EXTENSIONS and len(root_video_links) < 5:
                    root_video_links.append(url)
                    root_updated = True

            if on_file_uploaded:
                on_file_uploaded(file_size, local_path, b2_key, url)

            if (root_updated or idx == total_files) and on_root_links_updated:
                on_root_links_updated(root_image_links, root_video_links, idx, total_files)

        return root_image_links, root_video_links, total_files, total_bytes, sync_state
