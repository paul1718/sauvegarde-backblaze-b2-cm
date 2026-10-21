import io
import json
import os
import shutil
import time
import zipfile
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


MEDIA_COLUMNS: List[str] = (
    [f"Image_{i}" for i in range(1, 11)]
    + [f"Video_{i}" for i in range(1, 6)]
)

ALL_STANDARD_COLUMNS: List[str] = ["Dossier", "Statut"] + MEDIA_COLUMNS


class ExcelPermissionError(Exception):
    """Erreur levée lorsqu'un fichier Excel est ouvert dans une autre application."""
    pass


def _fsync_file_Write(path: Path, data_bytes: bytes) -> None:
    """Écrit des octets sur le disque et force le vidage matériel (fsync) sur le disque dur/SSD."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as f:
        f.write(data_bytes)
        f.flush()
        os.fsync(f.fileno())


def _is_valid_xlsx_file(path: Path) -> bool:
    """Vérifie qu'un fichier .xlsx existe, est une archive ZIP complète non tronquée et s'ouvre dans openpyxl."""
    if not path.exists():
        return False
    try:
        if path.stat().st_size < 100:
            return False
        with zipfile.ZipFile(path, "r") as zf:
            if zf.testzip() is not None:
                return False
        wb = openpyxl.load_workbook(path)
        wb.close()
        return True
    except Exception:
        return False


class ExcelManager:
    """
    Gère la lecture et la mise à jour du fichier Excel unique avec protection matérielle
    5 niveaux contre les coupures brusques d'électricité ("coupure de lumière") :
    1. Génération et validation ZIP en mémoire RAM (`io.BytesIO`) avant tout accès disque.
    2. Écriture dans un fichier temporaire `.~tmp` avec `flush()` + `os.fsync()` (forçage physique SSD/HDD).
    3. Double copie de secours synchronisée (`_backup.xlsx` local + copie miroir dans `%APPDATA%`).
    4. Journal de transaction (`excel_wal.json` avec `fsync`) mémorisant la dernière ligne écrite.
    5. Auto-réparation transparente à l'ouverture si une coupure de courant est survenue.
    """

    @staticmethod
    def _is_valid_xlsx_file(path: Path) -> bool:
        return _is_valid_xlsx_file(path)

    @staticmethod
    def _normalize_header(val: Any) -> str:
        if val is None:
            return ""
        return str(val).strip().lower()

    @staticmethod
    def is_completed_status(val: Any) -> bool:
        if val is None:
            return False
        normalized = str(val).strip().lower()
        return normalized in ("terminé", "termine", "terminée", "terminee")

    @staticmethod
    def is_empty_folder_status(val: Any) -> bool:
        if val is None:
            return False
        normalized = str(val).strip().lower()
        return normalized == "dossier vide"

    @staticmethod
    def _get_backup_path(path: Path) -> Path:
        return path.with_name(f"{path.stem}_backup{path.suffix}")

    @staticmethod
    def _get_appdata_backup_path(path: Path) -> Path:
        appdata = os.environ.get("APPDATA")
        base = Path(appdata) / "BackblazeB2Uploader" / "excel_backups" if appdata else Path.home() / ".backblaze_b2_uploader" / "excel_backups"
        base.mkdir(parents=True, exist_ok=True)
        return base / f"{path.stem}_safe_mirror{path.suffix}"

    @staticmethod
    def _get_wal_path(path: Path) -> Path:
        return path.with_name(f".~wal_{path.stem}.json")

    @staticmethod
    def _get_temp_path(path: Path) -> Path:
        return path.with_name(f".~tmp_{path.stem}{path.suffix}")

    @classmethod
    def _record_wal_entry(
        cls,
        path: Path,
        row_index: int,
        dossier_name: str,
        image_links: Optional[List[str]],
        video_links: Optional[List[str]],
        status: Optional[str],
    ) -> None:
        """Enregistre la transaction en cours dans un journal WAL atomique avec fsync avant de toucher au fichier Excel."""
        wal_path = cls._get_wal_path(path)
        wal_tmp = wal_path.with_suffix(".json.tmp")
        payload = {
            "row_index": row_index,
            "dossier_name": dossier_name,
            "image_links": image_links,
            "video_links": video_links,
            "status": status,
            "timestamp": time.time(),
        }
        try:
            raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            _fsync_file_Write(wal_tmp, raw)
            os.replace(wal_tmp, wal_path)
        except Exception:
            pass

    @classmethod
    def _clear_wal(cls, path: Path) -> None:
        wal_path = cls._get_wal_path(path)
        if wal_path.exists():
            try:
                wal_path.unlink()
            except Exception:
                pass

    @classmethod
    def _load_workbook_safe(cls, path: Path) -> openpyxl.Workbook:
        """
        Charge le classeur Excel en vérifiant son intégrité ZIP/XML complète.
        Si une coupure de courant a abîmé le fichier principal, restaure automatiquement
        la copie de secours intacte (`_backup.xlsx` ou miroir `%APPDATA%`) et rejoue le journal WAL.
        """
        bak_path = cls._get_backup_path(path)
        appdata_bak = cls._get_appdata_backup_path(path)
        tmp_path = cls._get_temp_path(path)

        # Si le fichier principal est valide, on l'ouvre directement (et on rejoue un éventuel WAL d'une coupure précédente)
        if _is_valid_xlsx_file(path):
            wb = openpyxl.load_workbook(path)
            if cls._get_wal_path(path).exists():
                cls._replay_wal_if_present(wb, path)
            return wb

        # Sinon : le fichier principal a été endommagé par une coupure brusque de courant !
        # On cherche la copie intacte la plus récente parmi : .~tmp (si complet), _backup.xlsx, ou miroir AppData
        for candidate in (tmp_path, bak_path, appdata_bak):
            if _is_valid_xlsx_file(candidate):
                data = candidate.read_bytes()
                _fsync_file_Write(path, data)
                wb = openpyxl.load_workbook(path)
                cls._replay_wal_if_present(wb, path)
                return wb

        # Si aucune copie de secours n'existe encore, tente l'ouverture standard pour lever l'erreur appropriée
        return openpyxl.load_workbook(path)

    @classmethod
    def _replay_wal_if_present(cls, wb: openpyxl.Workbook, path: Path) -> None:
        """Rejoue la dernière modification enregistrée dans le journal WAL si une coupure a eu lieu pendant la sauvegarde."""
        wal_path = cls._get_wal_path(path)
        if not wal_path.exists():
            return
        try:
            wal_data = json.loads(wal_path.read_text(encoding="utf-8"))
            ws = wb.active
            col_map, _ = cls._ensure_columns(ws)
            row_idx = int(wal_data.get("row_index", 2))
            dossier_name = str(wal_data.get("dossier_name", ""))
            target_row = cls._find_target_row(ws, col_map["Dossier"], row_idx, dossier_name)

            img_links = wal_data.get("image_links")
            if isinstance(img_links, list):
                for i in range(1, 11):
                    val = img_links[i - 1] if (i - 1) < len(img_links) else ""
                    ws.cell(row=target_row, column=col_map[f"Image_{i}"], value=val if val else None)

            vid_links = wal_data.get("video_links")
            if isinstance(vid_links, list):
                for i in range(1, 6):
                    val = vid_links[i - 1] if (i - 1) < len(vid_links) else ""
                    ws.cell(row=target_row, column=col_map[f"Video_{i}"], value=val if val else None)

            status = wal_data.get("status")
            if status:
                sc = ws.cell(row=target_row, column=col_map["Statut"], value=str(status))
                cls._apply_status_style(sc, str(status))

            cls._save_workbook_atomic(wb, path)
            cls._clear_wal(path)
        except Exception:
            pass

    @classmethod
    def _ensure_columns(cls, ws: openpyxl.worksheet.worksheet.Worksheet) -> Tuple[Dict[str, int], bool]:
        """
        Repère les colonnes existantes ('Dossier', 'Statut', 'Image_1'..'Image_10', 'Video_1'..'Video_5')
        sans jamais supprimer ni écraser les autres colonnes du fichier Excel.
        Crée dans les colonnes suivantes celles qui manquent encore.
        """
        col_map: Dict[str, int] = {}
        max_col = max(ws.max_column or 1, 1)

        for col_idx in range(1, max_col + 1):
            raw_val = ws.cell(row=1, column=col_idx).value
            norm = cls._normalize_header(raw_val)
            if not norm:
                continue
            if norm in ("dossier", "dossiers", "nom_dossier", "nom du dossier") and "Dossier" not in col_map:
                col_map["Dossier"] = col_idx
            elif norm in ("statut", "status", "état", "etat") and "Statut" not in col_map:
                col_map["Statut"] = col_idx
            else:
                for media_col in MEDIA_COLUMNS:
                    if norm == media_col.lower() and media_col not in col_map:
                        col_map[media_col] = col_idx

        if "Dossier" not in col_map:
            raise ValueError(
                "La colonne 'Dossier' est introuvable dans la première ligne du fichier Excel."
            )

        last_used_col = 0
        for col_idx in range(1, max_col + 1):
            if ws.cell(row=1, column=col_idx).value is not None:
                last_used_col = col_idx

        modified = False
        header_fill = PatternFill(start_color="102A5E", end_color="102A5E", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)

        for col_name in ["Statut"] + MEDIA_COLUMNS:
            if col_name not in col_map:
                last_used_col += 1
                col_map[col_name] = last_used_col
                c = ws.cell(row=1, column=last_used_col, value=col_name)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                col_letter = get_column_letter(last_used_col)
                ws.column_dimensions[col_letter].width = 20 if col_name == "Statut" else 36
                modified = True

        return col_map, modified

    @classmethod
    def _apply_status_style(cls, cell: openpyxl.cell.cell.Cell, status: str) -> None:
        """
        Applique le style visuel au statut :
        - 'Dossier vide' : colorié en ROUGE (texte rouge vif #E51B23 gras + fond rouge clair #FADBD8)
        - 'Terminé' : texte vert gras
        - 'En cours' : texte bleu gras
        """
        if cls.is_empty_folder_status(status):
            cell.font = Font(color="E51B23", bold=True)
            cell.fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
        elif cls.is_completed_status(status):
            cell.font = Font(color="16A34A", bold=True)
            cell.fill = PatternFill(fill_type=None)
        elif status.lower().startswith("en cours"):
            cell.font = Font(color="0878F9", bold=True)
            cell.fill = PatternFill(fill_type=None)
        elif status.lower().startswith("erreur"):
            cell.font = Font(color="E51B23", bold=True)
            cell.fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
        else:
            cell.font = Font(bold=False)
            cell.fill = PatternFill(fill_type=None)

    @classmethod
    def read_source_folders(cls, source_path: str) -> List[Dict[str, Any]]:
        path = Path(source_path)
        bak_path = cls._get_backup_path(path)
        appdata_bak = cls._get_appdata_backup_path(path)
        if not path.exists() and not bak_path.exists() and not appdata_bak.exists():
            raise FileNotFoundError(f"Le fichier Excel est introuvable : {source_path}")

        wb = cls._load_workbook_safe(path)
        ws = wb.active

        col_map, need_save = cls._ensure_columns(ws)
        dossier_col = col_map["Dossier"]
        statut_col = col_map["Statut"]

        folders: List[Dict[str, Any]] = []
        for row_idx in range(2, (ws.max_row or 1) + 1):
            raw_dossier = ws.cell(row=row_idx, column=dossier_col).value
            if raw_dossier is None or str(raw_dossier).strip() == "":
                continue
            dossier_name = str(raw_dossier).strip()
            raw_statut = ws.cell(row=row_idx, column=statut_col).value
            statut_str = "" if raw_statut is None else str(raw_statut).strip()
            completed = cls.is_completed_status(statut_str)
            empty_status = cls.is_empty_folder_status(statut_str)

            existing_images: List[str] = []
            for i in range(1, 11):
                v = ws.cell(row=row_idx, column=col_map[f"Image_{i}"]).value
                existing_images.append("" if v is None else str(v).strip())

            existing_videos: List[str] = []
            for i in range(1, 6):
                v = ws.cell(row=row_idx, column=col_map[f"Video_{i}"]).value
                existing_videos.append("" if v is None else str(v).strip())

            folders.append(
                {
                    "row_index": row_idx,
                    "dossier": dossier_name,
                    "statut": statut_str,
                    "completed": completed,
                    "empty_status": empty_status,
                    "existing_images": existing_images,
                    "existing_videos": existing_videos,
                }
            )

        # Dès la première lecture, on crée/met à jour les copies de secours matérielles
        cls._save_workbook_atomic(wb, path)
        wb.close()
        return folders

    @classmethod
    def _find_target_row(
        cls,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        dossier_col: int,
        row_index: int,
        dossier_name: str,
    ) -> int:
        current_val = ws.cell(row=row_index, column=dossier_col).value
        if current_val is not None and str(current_val).strip() == dossier_name:
            return row_index
        for r in range(2, (ws.max_row or 1) + 1):
            v = ws.cell(row=r, column=dossier_col).value
            if v is not None and str(v).strip() == dossier_name:
                return r
        return row_index

    @classmethod
    def update_source_status(
        cls,
        source_path: str,
        row_index: int,
        dossier_name: str,
        status: str = "Terminé",
    ) -> None:
        path = Path(source_path)
        wb = cls._load_workbook_safe(path)
        cls._record_wal_entry(path, row_index, dossier_name, None, None, status)
        ws = wb.active

        col_map, _ = cls._ensure_columns(ws)
        target_row = cls._find_target_row(ws, col_map["Dossier"], row_index, dossier_name)

        status_cell = ws.cell(row=target_row, column=col_map["Statut"], value=status)
        cls._apply_status_style(status_cell, status)

        cls._save_workbook_atomic(wb, path)
        wb.close()
        cls._clear_wal(path)

    @classmethod
    def update_row_links(
        cls,
        excel_path: str,
        row_index: int,
        dossier_name: str,
        image_links: List[str],
        video_links: List[str],
        status: Optional[str] = None,
    ) -> None:
        path = Path(excel_path)
        wb = cls._load_workbook_safe(path)
        cls._record_wal_entry(path, row_index, dossier_name, image_links, video_links, status)
        ws = wb.active

        col_map, _ = cls._ensure_columns(ws)
        target_row = cls._find_target_row(ws, col_map["Dossier"], row_index, dossier_name)

        for i in range(1, 11):
            col_idx = col_map[f"Image_{i}"]
            new_val = image_links[i - 1] if (i - 1) < len(image_links) else ""
            cell = ws.cell(row=target_row, column=col_idx)
            if (cell.value or "") != new_val:
                cell.value = new_val if new_val != "" else None

        for i in range(1, 6):
            col_idx = col_map[f"Video_{i}"]
            new_val = video_links[i - 1] if (i - 1) < len(video_links) else ""
            cell = ws.cell(row=target_row, column=col_idx)
            if (cell.value or "") != new_val:
                cell.value = new_val if new_val != "" else None

        if status is not None:
            status_cell = ws.cell(row=target_row, column=col_map["Statut"], value=status)
            cls._apply_status_style(status_cell, status)

        cls._save_workbook_atomic(wb, path)
        wb.close()
        cls._clear_wal(path)

    @classmethod
    def create_source_from_directory(cls, excel_path: str, parent_dir: Optional[str] = None) -> int:
        path = Path(excel_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        subfolders: List[str] = []
        if parent_dir and os.path.isdir(parent_dir):
            try:
                subfolders = sorted(
                    [
                        d.name
                        for d in Path(parent_dir).iterdir()
                        if d.is_dir() and not d.name.startswith(".")
                    ],
                    key=lambda s: s.lower(),
                )
            except Exception:
                subfolders = []

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dossiers"

        header_fill = PatternFill(start_color="102A5E", end_color="102A5E", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)

        for col_idx, h in enumerate(ALL_STANDARD_COLUMNS, start=1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center")

        for row_idx, folder_name in enumerate(subfolders, start=2):
            ws.cell(row=row_idx, column=1, value=folder_name)
            ws.cell(row=row_idx, column=2, value="")

        ws.column_dimensions["A"].width = 34
        ws.column_dimensions["B"].width = 20
        for col_idx in range(3, len(ALL_STANDARD_COLUMNS) + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = 36
        ws.row_dimensions[1].height = 24
        ws.freeze_panes = "C2"

        cls._save_workbook_atomic(wb, path)
        wb.close()
        return len(subfolders)

    @classmethod
    def _save_workbook_atomic(
        cls,
        wb: openpyxl.Workbook,
        path: Path,
        retries: int = 6,
        delay: float = 0.8,
    ) -> None:
        """
        Sauvegarde matérielle anti-coupure d'électricité ("coupure de lumière") :
        1. Sérialise d'abord le classeur en mémoire RAM (`io.BytesIO`) et vérifie que l'archive ZIP est 100% valide.
        2. Si l'ancien fichier sur disque est valide, met à jour les copies de secours (`_backup.xlsx` et miroir `%APPDATA%`) avec `fsync`.
        3. Écrit les octets validés dans `.~tmp` avec `flush()` + `os.fsync()` (forçage physique sur le disque).
        4. Remplace atomiquement le fichier cible via `os.replace(tmp_path, path)`.
        5. Met à jour la copie `_backup.xlsx` avec `fsync`.
        """
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = cls._get_temp_path(path)
        bak_path = cls._get_backup_path(path)
        appdata_bak = cls._get_appdata_backup_path(path)

        # 1. Génération et contrôle d'intégrité ZIP en mémoire vive
        mem_buffer = io.BytesIO()
        wb.save(mem_buffer)
        xlsx_bytes = mem_buffer.getvalue()

        with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zf:
            if zf.testzip() is not None:
                raise IOError("Le classeur Excel généré en mémoire est invalide.")

        last_err: Optional[Exception] = None
        for attempt in range(1, retries + 1):
            try:
                # 2. Écriture physique forcée (fsync) dans le fichier temporaire
                _fsync_file_Write(tmp_path, xlsx_bytes)

                # 3. Sauvegarde préalable de l'ancien fichier valide (si existant)
                if _is_valid_xlsx_file(path):
                    try:
                        old_bytes = path.read_bytes()
                        _fsync_file_Write(bak_path, old_bytes)
                    except Exception:
                        pass

                # 4. Remplacement atomique NTFS
                os.replace(tmp_path, path)

                # 5. Mise à jour immédiate des 2 copies de secours (locale + AppData) avec fsync
                try:
                    _fsync_file_Write(bak_path, xlsx_bytes)
                except Exception:
                    pass
                try:
                    _fsync_file_Write(appdata_bak, xlsx_bytes)
                except Exception:
                    pass
                return

            except PermissionError as e:
                last_err = e
                time.sleep(delay * attempt)
            except Exception as e:
                last_err = e
                time.sleep(delay)
            finally:
                if tmp_path.exists():
                    try:
                        tmp_path.unlink()
                    except Exception:
                        pass

        if isinstance(last_err, PermissionError):
            raise ExcelPermissionError(
                f"Impossible d'enregistrer '{path.name}'. Veuillez fermer ce fichier dans Microsoft Excel puis réessayer."
            ) from last_err
        if last_err:
            raise last_err
