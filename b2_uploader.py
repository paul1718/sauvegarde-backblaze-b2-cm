import os
import re
import time
import threading
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, Set
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
    # Si le chemin n'est pas déjà encodé avec %, on encode proprement chaque segment
    if "%" not in parts.path:
        encoded_path = "/".join(quote(seg, safe="") for seg in parts.path.split("/"))
    else:
        encoded_path = parts.path
    return urlunsplit((parts.scheme, parts.netloc, encoded_path, parts.query, parts.fragment))


class UploadCancelledError(Exception):
    """Exception interne levée lorsque l'utilisateur demande l'arrêt."""
    pass


class B2UploaderService:
    """
    Service gérant la connexion à Backblaze B2, le parcours récursif des dossiers,
    la reprise sur micro-coupures réseau (retry exponentiel) et la génération des liens publics.
    """

    def __init__(
        self,
        key_id: str,
        app_key: str,
        bucket_name: str,
        log_callback: Optional[Callable[[str, str], None]] = None,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        pause_event: Optional[threading.Event] = None,
        stop_event: Optional[threading.Event] = None,
    ) -> None:
        self.key_id = key_id.strip()
        self.app_key = app_key.strip()
        self.bucket_name = bucket_name.strip()
        self.log_callback = log_callback
        self.progress_callback = progress_callback
        # pause_event: set() = en cours d'exécution, clear() = en pause
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

    def _check_pause_and_stop(self) -> None:
        """Vérifie si l'utilisateur a mis en pause ou arrêté le processus."""
        if self.stop_event.is_set():
            raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")
        if not self.pause_event.is_set():
            self._log("Sauvegarde en pause. En attente de reprise...", "WARN")
            while not self.pause_event.is_set():
                if self.stop_event.is_set():
                    raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")
                time.sleep(0.25)
            self._log("Reprise de la sauvegarde...", "INFO")

    def connect(self, max_retries: int = 5) -> None:
        """Authentifie le compte Backblaze B2 et récupère le Bucket avec gestion de retry."""
        last_error: Optional[Exception] = None
        for attempt in range(1, max_retries + 1):
            self._check_pause_and_stop()
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
                if attempt < max_retries:
                    wait_sec = min(2 ** attempt, 30)
                    self._log(
                        f"Tentative de connexion B2 ({attempt}/{max_retries}) échouée : {e}. Nouvel essai dans {wait_sec}s...",
                        "WARN",
                    )
                    self._sleep_interruptible(wait_sec)
        raise ConnectionError(f"Impossible de se connecter à Backblaze B2 : {last_error}")

    def _sleep_interruptible(self, seconds: float) -> None:
        """Attend `seconds` secondes tout en restant réactif à l'arrêt immédiat."""
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
            # Tri des sous-dossiers pour un parcours déterministe
            dirs.sort(key=lambda d: d.lower())
            root_path = Path(root)
            is_root = root_path.resolve() == folder_path.resolve()

            # Tri naturel des fichiers du répertoire courant
            file_paths = [root_path / f for f in files]
            file_paths.sort(key=natural_sort_key)

            for file_path in file_paths:
                ext = file_path.suffix.lower()
                if ext in ALL_ALLOWED_EXTENSIONS:
                    rel_inside_folder = file_path.relative_to(folder_path).as_posix()
                    b2_key = f"{folder_name}/{rel_inside_folder}"
                    eligible_files.append((file_path, b2_key, is_root))

        return eligible_files

    def _fetch_existing_b2_files(self, folder_name: str) -> Dict[str, int]:
        """
        Récupère la liste des fichiers déjà présents sur B2 dans `folder_name/`
        avec leur taille en octets pour accélérer les reprises après interruption.
        """
        existing: Dict[str, int] = {}
        if not self.bucket:
            return existing
        try:
            prefix = f"{folder_name}/"
            for file_version, _ in self.bucket.ls(folder_to_list=prefix, recursive=True):
                existing[file_version.file_name] = int(file_version.size)
        except Exception:
            # Si le listage échoue, on téléversera normalement sans bloquer
            pass
        return existing

    def _upload_single_file_with_retry(
        self,
        local_path: Path,
        b2_key: str,
        existing_files: Dict[str, int],
        max_retries: int = 8,
    ) -> str:
        """
        Téléverse un fichier sur Backblaze B2 avec politique de retry robuste
        et retourne son URL publique.
        """
        self._check_pause_and_stop()
        local_size = local_path.stat().st_size

        # Si le fichier existe déjà sur B2 avec exactement la même taille (reprise après coupure)
        if b2_key in existing_files and existing_files[b2_key] == local_size:
            raw_url = self.bucket.get_download_url(b2_key)
            return normalize_b2_url(raw_url)

        last_err: Optional[Exception] = None
        for attempt in range(1, max_retries + 1):
            self._check_pause_and_stop()
            try:
                if self.bucket is None:
                    self.connect(max_retries=3)

                self.bucket.upload_local_file(
                    local_file=str(local_path),
                    file_name=b2_key,
                )
                raw_url = self.bucket.get_download_url(b2_key)
                return normalize_b2_url(raw_url)

            except UploadCancelledError:
                raise
            except (B2Error, OSError, Exception) as e:
                last_err = e
                if attempt < max_retries:
                    wait_sec = min(3 * (2 ** (attempt - 1)), 60)
                    self._log(
                        f"Micro-coupure ou erreur réseau sur '{local_path.name}' "
                        f"(tentative {attempt}/{max_retries}) : {e}. "
                        f"Nouvel essai dans {wait_sec}s...",
                        "WARN",
                    )
                    self._sleep_interruptible(wait_sec)
                    # Ré-authentification préventive en cas d'expiration de jeton B2 sur de très longs transferts
                    try:
                        self.connect(max_retries=2)
                    except UploadCancelledError:
                        raise
                    except Exception:
                        pass

        raise RuntimeError(
            f"Échec du téléversement de '{local_path.name}' après {max_retries} tentatives : {last_err}"
        )

    def upload_folder(
        self,
        folder_path: Path,
        on_file_uploaded: Optional[Callable[[int, Path], None]] = None,
    ) -> Tuple[List[str], List[str], int, int]:
        """
        Téléverse l'intégralité d'un dossier (et tous ses sous-dossiers) vers Backblaze B2.
        Retourne :
        - root_image_links : jusqu'aux 10 premiers liens d'images situées à la racine
        - root_video_links : jusqu'aux 5 premiers liens de vidéos situées à la racine
        - total_files_uploaded : nombre total de fichiers traités dans ce dossier
        - total_bytes_uploaded : volume total d'octets traités dans ce dossier
        """
        self._check_pause_and_stop()
        if not folder_path.exists() or not folder_path.is_dir():
            raise FileNotFoundError(f"Le dossier est introuvable : {folder_path}")

        eligible_files = self.scan_folder_files(folder_path)
        existing_b2 = self._fetch_existing_b2_files(folder_path.name)

        root_image_links: List[str] = []
        root_video_links: List[str] = []
        total_files = len(eligible_files)
        total_bytes = 0

        for idx, (local_path, b2_key, is_root) in enumerate(eligible_files, start=1):
            self._check_pause_and_stop()
            file_size = local_path.stat().st_size
            ext = local_path.suffix.lower()

            already_on_b2 = b2_key in existing_b2 and existing_b2[b2_key] == file_size
            url = self._upload_single_file_with_retry(local_path, b2_key, existing_b2)
            total_bytes += file_size

            if already_on_b2:
                self._log(
                    f"  ✓ [{idx}/{total_files}] Déjà sur B2 : {b2_key} ({format_size(file_size)})",
                    "INFO",
                )
            else:
                self._log(
                    f"  ↑ [{idx}/{total_files}] Téléversé : {b2_key} ({format_size(file_size)})",
                    "SUCCESS",
                )

            # Règle métier C : Seuls les fichiers à la racine du dossier (hors documents)
            # sont ajoutés à l'Excel de sortie (max 10 images, max 5 vidéos)
            if is_root:
                if ext in IMAGE_EXTENSIONS and len(root_image_links) < 10:
                    root_image_links.append(url)
                elif ext in VIDEO_EXTENSIONS and len(root_video_links) < 5:
                    root_video_links.append(url)
                # Les documents (.txt, .pdf, .doc, .docx, .csv, .rtf) sont bien téléversés
                # mais leurs liens ne doivent pas figurer dans l'Excel de sortie.

            if on_file_uploaded:
                on_file_uploaded(file_size, local_path)

        return root_image_links, root_video_links, total_files, total_bytes
