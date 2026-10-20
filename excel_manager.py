import os
import shutil
import time
from pathlib import Path
from typing import List, Dict, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


OUTPUT_COLUMNS: List[str] = (
    ["Dossier"]
    + [f"Image_{i}" for i in range(1, 11)]
    + [f"Video_{i}" for i in range(1, 6)]
)


class ExcelPermissionError(Exception):
    """Erreur levée lorsqu'un fichier Excel est ouvert dans une autre application."""
    pass


class ExcelManager:
    """
    Gère la lecture et la mise à jour des fichiers Excel Source et Sortie
    avec protection anti-corruption en cas de coupure brusque d'électricité
    (écriture atomique via fichier temporaire + copie de secours .bak + auto-réparation).
    """

    @staticmethod
    def _normalize_header(val: Any) -> str:
        if val is None:
            return ""
        return str(val).strip().lower()

    @staticmethod
    def is_completed_status(val: Any) -> bool:
        """Vérifie si la valeur de la cellule Statut correspond à 'Terminé'."""
        if val is None or pd.isna(val):
            return False
        normalized = str(val).strip().lower()
        return normalized in ("terminé", "termine", "terminée", "terminee")

    @staticmethod
    def _get_backup_path(path: Path) -> Path:
        return path.with_name(f"{path.stem}_backup{path.suffix}")

    @staticmethod
    def _get_temp_path(path: Path) -> Path:
        return path.with_name(f".~tmp_{path.stem}{path.suffix}")

    @classmethod
    def _load_workbook_safe(cls, path: Path) -> openpyxl.Workbook:
        """
        Charge un classeur Excel en toute sécurité.
        Si le fichier principal a été endommagé par une coupure électrique brutale,
        restaure automatiquement la copie de secours `_backup.xlsx`.
        """
        bak_path = cls._get_backup_path(path)
        try:
            return openpyxl.load_workbook(path)
        except Exception as primary_err:
            if bak_path.exists():
                try:
                    wb = openpyxl.load_workbook(bak_path)
                    shutil.copy2(bak_path, path)
                    return wb
                except Exception:
                    pass
            raise primary_err

    @classmethod
    def _read_excel_dataframe_safe(cls, path: Path) -> pd.DataFrame:
        """
        Lit un fichier Excel en DataFrame avec restauration automatique depuis `_backup.xlsx`
        si le fichier principal est corrompu.
        """
        bak_path = cls._get_backup_path(path)
        if not path.exists() and not bak_path.exists():
            return pd.DataFrame(columns=OUTPUT_COLUMNS)

        for candidate in (path, bak_path):
            if candidate.exists():
                try:
                    df = pd.read_excel(candidate, dtype=str).fillna("")
                    if candidate == bak_path:
                        try:
                            shutil.copy2(bak_path, path)
                        except Exception:
                            pass
                    return df
                except Exception:
                    continue
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    @classmethod
    def read_source_folders(cls, source_path: str) -> List[Dict[str, Any]]:
        """
        Lit le fichier Excel source (dossiers.xlsx).
        Assure la présence des colonnes 'Dossier' et 'Statut'.
        Retourne une liste de dictionnaires :
        [{'row_index': int, 'dossier': str, 'statut': str, 'completed': bool}, ...]
        """
        path = Path(source_path)
        bak_path = cls._get_backup_path(path)
        if not path.exists() and not bak_path.exists():
            raise FileNotFoundError(f"Le fichier Excel source est introuvable : {source_path}")

        wb = cls._load_workbook_safe(path)
        ws = wb.active

        dossier_col: Optional[int] = None
        statut_col: Optional[int] = None

        max_col = max(ws.max_column or 1, 1)
        for col_idx in range(1, max_col + 1):
            cell_val = cls._normalize_header(ws.cell(row=1, column=col_idx).value)
            if cell_val in ("dossier", "dossiers", "nom_dossier", "nom du dossier"):
                dossier_col = col_idx
            elif cell_val in ("statut", "status", "état", "etat"):
                statut_col = col_idx

        if dossier_col is None:
            wb.close()
            raise ValueError(
                "La colonne 'Dossier' est introuvable dans la première ligne du fichier Excel source."
            )

        need_save = False
        if statut_col is None:
            statut_col = max_col + 1
            header_cell = ws.cell(row=1, column=statut_col, value="Statut")
            header_cell.font = Font(bold=True)
            need_save = True

        folders: List[Dict[str, Any]] = []
        for row_idx in range(2, (ws.max_row or 1) + 1):
            raw_dossier = ws.cell(row=row_idx, column=dossier_col).value
            if raw_dossier is None or str(raw_dossier).strip() == "":
                continue
            dossier_name = str(raw_dossier).strip()
            raw_statut = ws.cell(row=row_idx, column=statut_col).value
            statut_str = "" if raw_statut is None else str(raw_statut).strip()
            completed = cls.is_completed_status(statut_str)

            folders.append(
                {
                    "row_index": row_idx,
                    "dossier": dossier_name,
                    "statut": statut_str,
                    "completed": completed,
                }
            )

        if need_save:
            cls._save_workbook_atomic(wb, path)
        wb.close()
        return folders

    @classmethod
    def update_source_status(
        cls,
        source_path: str,
        row_index: int,
        dossier_name: str,
        status: str = "Terminé",
    ) -> None:
        """
        Met à jour immédiatement et atomiquement la colonne 'Statut'
        pour la ligne correspondante dans le fichier source.
        """
        path = Path(source_path)
        wb = cls._load_workbook_safe(path)
        ws = wb.active

        dossier_col: Optional[int] = None
        statut_col: Optional[int] = None
        max_col = max(ws.max_column or 1, 1)

        for col_idx in range(1, max_col + 1):
            cell_val = cls._normalize_header(ws.cell(row=1, column=col_idx).value)
            if cell_val in ("dossier", "dossiers", "nom_dossier", "nom du dossier"):
                dossier_col = col_idx
            elif cell_val in ("statut", "status", "état", "etat"):
                statut_col = col_idx

        if statut_col is None:
            statut_col = max_col + 1
            ws.cell(row=1, column=statut_col, value="Statut")

        target_row = row_index
        if dossier_col is not None:
            current_val = ws.cell(row=target_row, column=dossier_col).value
            if current_val is None or str(current_val).strip() != dossier_name:
                for r in range(2, (ws.max_row or 1) + 1):
                    v = ws.cell(row=r, column=dossier_col).value
                    if v is not None and str(v).strip() == dossier_name:
                        target_row = r
                        break

        status_cell = ws.cell(row=target_row, column=statut_col, value=status)
        if cls.is_completed_status(status):
            status_cell.font = Font(color="16A34A", bold=True)
        elif status.lower().startswith("en cours"):
            status_cell.font = Font(color="0878F9", bold=True)
        elif status.lower().startswith("erreur"):
            status_cell.font = Font(color="E51B23", bold=True)

        cls._save_workbook_atomic(wb, path)
        wb.close()

    @classmethod
    def upsert_output_result(
        cls,
        output_path: str,
        dossier_name: str,
        image_links: List[str],
        video_links: List[str],
    ) -> None:
        """
        Ajoute ou met à jour en temps réel la ligne du dossier dans le fichier Excel de sortie (resultats_backblaze.xlsx).
        Utilise une écriture atomique pour garantir que le fichier n'est jamais corrompu en cas de coupure de courant.
        """
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        row_data: Dict[str, str] = {"Dossier": dossier_name}
        for i in range(1, 11):
            row_data[f"Image_{i}"] = image_links[i - 1] if (i - 1) < len(image_links) else ""
        for i in range(1, 6):
            row_data[f"Video_{i}"] = video_links[i - 1] if (i - 1) < len(video_links) else ""

        df = cls._read_excel_dataframe_safe(path)

        for col in OUTPUT_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[OUTPUT_COLUMNS]

        mask = df["Dossier"].astype(str).str.strip() == dossier_name
        if mask.any():
            for col, val in row_data.items():
                df.loc[mask, col] = val
        else:
            new_row_df = pd.DataFrame([row_data], columns=OUTPUT_COLUMNS)
            df = pd.concat([df, new_row_df], ignore_index=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Résultats Backblaze"

        header_fill = PatternFill(start_color="102A5E", end_color="102A5E", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
            for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
                val = str(row[col_name]) if pd.notna(row[col_name]) else ""
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                if col_idx == 1:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(vertical="center")
                else:
                    cell.alignment = Alignment(vertical="center")

        ws.column_dimensions["A"].width = 28
        for col_idx in range(2, len(OUTPUT_COLUMNS) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 38
        ws.row_dimensions[1].height = 24
        ws.freeze_panes = "B2"

        cls._save_workbook_atomic(wb, path)
        wb.close()

    @classmethod
    def create_source_from_directory(cls, excel_path: str, parent_dir: Optional[str] = None) -> int:
        """
        Crée ou initialise un fichier Excel source `dossiers.xlsx`.
        Si `parent_dir` est fourni et valide, liste automatiquement les sous-dossiers directs.
        """
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

        headers = ["Dossier", "Statut"]
        for col_idx, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center")

        for row_idx, folder_name in enumerate(subfolders, start=2):
            ws.cell(row=row_idx, column=1, value=folder_name)
            ws.cell(row=row_idx, column=2, value="")

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 20
        ws.row_dimensions[1].height = 24
        ws.freeze_panes = "A2"

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
        Sauvegarde un classeur Excel de manière 100% atomique (anti-coupure de courant) :
        1. Écrit dans un fichier temporaire `.~tmp_<nom>.xlsx`
        2. Vérifie que le fichier temporaire est bien lisible (ZIP intègre)
        3. Copie l'ancien fichier vers `<nom>_backup.xlsx`
        4. Remplace atomiquement le fichier cible avec `os.replace(tmp_path, path)`
        """
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = cls._get_temp_path(path)
        bak_path = cls._get_backup_path(path)

        last_err: Optional[Exception] = None
        for attempt in range(1, retries + 1):
            try:
                # 1. Écriture complète dans le fichier temporaire
                wb.save(tmp_path)

                # 2. Copie de sécurité de l'ancien fichier valide
                if path.exists() and path.stat().st_size > 0:
                    try:
                        shutil.copy2(path, bak_path)
                    except Exception:
                        pass

                # 3. Remplacement atomique NTFS
                os.replace(tmp_path, path)

                # 4. Mise à jour de la copie de secours avec la dernière version validée
                try:
                    shutil.copy2(path, bak_path)
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
