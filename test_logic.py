import os
import tempfile
from pathlib import Path
from typing import Dict, Any, List, Tuple
import pandas as pd

from config_manager import ConfigManager
from excel_manager import ExcelManager, OUTPUT_COLUMNS
from b2_uploader import (
    B2UploaderService,
    is_internet_connected,
    resolve_matching_folder,
    is_tav_folder,
)
from auto_updater import AutoUpdater
from generate_icon import create_app_icon


class FakeB2FileVersion:
    def __init__(self, file_name: str, size: int, file_id: str) -> None:
        self.file_name = file_name
        self.size = size
        self.id_ = file_id


class FakeB2Bucket:
    def __init__(self) -> None:
        self.files: Dict[str, FakeB2FileVersion] = {}
        self.upload_calls: List[str] = []
        self.delete_calls: List[str] = []

    def ls(self, folder_to_list: str = "", recursive: bool = True) -> List[Tuple[FakeB2FileVersion, str]]:
        return [
            (fv, "")
            for name, fv in list(self.files.items())
            if name.startswith(folder_to_list)
        ]

    def upload_local_file(self, local_file: str, file_name: str) -> FakeB2FileVersion:
        sz = Path(local_file).stat().st_size
        fv = FakeB2FileVersion(file_name, sz, f"id_{len(self.files) + 1}")
        self.files[file_name] = fv
        self.upload_calls.append(file_name)
        return fv

    def delete_file_version(self, file_id: str, file_name: str) -> None:
        self.files.pop(file_name, None)
        self.delete_calls.append(file_name)

    def get_download_url(self, file_name: str) -> str:
        return f"https://f000.backblazeb2.com/file/mon-bucket/{file_name}"


def run_tests() -> None:
    create_app_icon()

    print("1. Test de la résolution de dossier par préfixe...")
    with tempfile.TemporaryDirectory() as tmpdir:
        parent = Path(tmpdir)
        folder_exact = parent / "Mariage_Jean"
        folder_prefix = parent / "TAV_204 - Terrain Bastos (Lot 3)"
        folder_empty_tav = parent / "TAV_999_Parcelle_Vide"
        folder_empty_other = parent / "Anniversaire_Vide_2026"

        folder_exact.mkdir()
        folder_prefix.mkdir()
        folder_empty_tav.mkdir()
        folder_empty_other.mkdir()

        # Vérification que "TAV_204" dans Excel trouve bien "TAV_204 - Terrain Bastos (Lot 3)"
        matched = resolve_matching_folder(parent, "TAV_204")
        assert matched is not None and matched.name == "TAV_204 - Terrain Bastos (Lot 3)", f"Obtenu : {matched}"
        assert is_tav_folder("TAV_999", folder_empty_tav) is True
        assert is_tav_folder("Anniversaire", folder_empty_other) is False
        print("   -> OK : Correspondance exacte et par préfixe validées.")

        print("2. Test des dossiers vides ('TAV' avec image standard unique vs hors 'TAV' sans rien)...")
        source_xlsx = parent / "dossiers.xlsx"
        output_xlsx = parent / "resultats_backblaze.xlsx"
        df_src = pd.DataFrame(
            [
                {"Dossier": "TAV_999", "Statut": ""},
                {"Dossier": "Anniversaire", "Statut": ""},
            ]
        )
        df_src.to_excel(source_xlsx, index=False)

        uploader = B2UploaderService("k", "a", "b")
        uploader.bucket = FakeB2Bucket()

        std_img_file = parent / "image_standard_unique.jpg"
        std_img_file.write_bytes(b"standard_image_bytes_12345")

        # Premier appel : téléverse l'image standard une fois
        std_url_1 = uploader.ensure_tav_standard_image_url(str(std_img_file))
        # Second appel : réutilise exactement le même lien sans re-téléverser
        std_url_2 = uploader.ensure_tav_standard_image_url(str(std_img_file))
        assert std_url_1 == std_url_2
        assert len(uploader.bucket.upload_calls) == 1, "L'image standard TAV ne doit être téléversée qu'une seule fois !"

        # Application de la règle pour TAV_999 (vide) -> Image standard dans Excel final + "Dossier vide" dans Excel source
        ExcelManager.upsert_output_result(str(output_xlsx), "TAV_999", [std_url_1], [])
        ExcelManager.update_source_status(str(source_xlsx), 2, "TAV_999", "Dossier vide")

        # Application de la règle pour Anniversaire (vide, non TAV) -> Rien dans Excel final + "Dossier vide" dans Excel source
        ExcelManager.update_source_status(str(source_xlsx), 3, "Anniversaire", "Dossier vide")

        folders_after = ExcelManager.read_source_folders(str(source_xlsx))
        assert folders_after[0]["statut"] == "Dossier vide" and folders_after[0]["empty_status"] is True
        assert folders_after[1]["statut"] == "Dossier vide" and folders_after[1]["empty_status"] is True

        df_out = ExcelManager._read_excel_dataframe_safe(output_xlsx)
        assert len(df_out) == 1
        assert df_out.loc[0, "Dossier"] == "TAV_999"
        assert df_out.loc[0, "Image_1"] == std_url_1
        assert df_out.loc[0, "Image_2"] == ""
        print("   -> OK : Gestion des dossiers vides ('Dossier vide' + image standard unique pour 'TAV') validée.")

        print("3. Test de la synchronisation intelligente d'un dossier déjà sauvegardé (identique vs modifié)...")
        (folder_prefix / "photo1.jpg").write_bytes(b"img_data_1")
        (folder_prefix / "video1.mp4").write_bytes(b"vid_data_1")

        # 1ère sauvegarde : création sur B2
        uploader.bucket.upload_calls.clear()
        imgs1, vids1, _, _, state1 = uploader.upload_folder(folder_prefix)
        assert state1 == "created"
        assert len(uploader.bucket.upload_calls) == 2
        assert len(imgs1) == 1 and len(vids1) == 1

        # 2ème sauvegarde sans modification : le dossier source est identique au dossier sauvegardé
        # -> ne rien faire sur B2, mais retourner quand même les liens pour l'Excel final
        uploader.bucket.upload_calls.clear()
        imgs2, vids2, _, _, state2 = uploader.upload_folder(folder_prefix)
        assert state2 == "identical", f"Attendu 'identical', obtenu '{state2}'"
        assert len(uploader.bucket.upload_calls) == 0, "Aucun fichier ne doit être re-téléversé si le dossier est identique !"
        assert imgs2 == imgs1 and vids2 == vids1

        # 3ème sauvegarde après ajout d'une nouvelle photo2.jpg dans le dossier source
        # -> mise à jour uniquement du fichier nouveau dans le dossier sauvegardé + liens complets pour l'Excel final
        (folder_prefix / "photo2.jpg").write_bytes(b"img_data_2")
        uploader.bucket.upload_calls.clear()
        imgs3, vids3, _, _, state3 = uploader.upload_folder(folder_prefix)
        assert state3 == "updated", f"Attendu 'updated', obtenu '{state3}'"
        assert len(uploader.bucket.upload_calls) == 1, "Seul le nouveau fichier photo2.jpg doit être téléversé !"
        assert len(imgs3) == 2 and len(vids3) == 1
        print("   -> OK : Comparaison et mise à jour différentielle des dossiers déjà sauvegardés validées !")

    print("\nTOUS LES TESTS SONT PASSÉS AVEC SUCCÈS !")


if __name__ == "__main__":
    run_tests()
