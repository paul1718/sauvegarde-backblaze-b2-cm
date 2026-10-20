import os
import tempfile
from pathlib import Path
import pandas as pd

from config_manager import ConfigManager
from excel_manager import ExcelManager, OUTPUT_COLUMNS
from b2_uploader import B2UploaderService


def run_tests() -> None:
    print("1. Test du parcours récursif et du filtrage des extensions...")
    with tempfile.TemporaryDirectory() as tmpdir:
        parent = Path(tmpdir)
        folder = parent / "Mariage_Jean"
        subfolder = folder / "Sous_Dossier_Coulisses"
        subfolder.mkdir(parents=True)

        # Création de 12 images à la racine, 6 vidéos à la racine, 2 documents à la racine, 1 fichier non autorisé (.exe)
        for i in range(1, 13):
            (folder / f"img_{i}.jpg").write_bytes(b"fake_image_data")
        for i in range(1, 7):
            (folder / f"vid_{i}.mp4").write_bytes(b"fake_video_data")
        (folder / "contrat.pdf").write_bytes(b"fake_pdf")
        (folder / "notes.txt").write_bytes(b"fake_txt")
        (folder / "ignore_me.exe").write_bytes(b"fake_exe")

        # Création de fichiers dans le sous-dossier (doivent être téléversés sur B2 mais PAS dans l'Excel de sortie)
        (subfolder / "sub_img.png").write_bytes(b"sub_img")
        (subfolder / "sub_vid.mov").write_bytes(b"sub_vid")
        (subfolder / "sub_doc.docx").write_bytes(b"sub_doc")

        scanned = B2UploaderService.scan_folder_files(folder)
        # 12 images racine + 6 vidéos racine + 2 docs racine + 3 fichiers sous-dossier = 23 fichiers éligibles
        assert len(scanned) == 23, f"Attendu 23 fichiers éligibles, obtenu {len(scanned)}"

        # Vérification du tri naturel : img_2.jpg doit être avant img_10.jpg
        root_images = [p.name for p, _, is_root in scanned if is_root and p.suffix == ".jpg"]
        assert root_images[:3] == ["img_1.jpg", "img_2.jpg", "img_3.jpg"], f"Tri naturel incorrect : {root_images}"
        assert root_images[9] == "img_10.jpg"

        # Vérification des chemins relatifs B2 (arborescence conservée avec /)
        b2_keys = [key for _, key, _ in scanned]
        assert "Mariage_Jean/img_1.jpg" in b2_keys
        assert "Mariage_Jean/contrat.pdf" in b2_keys
        assert "Mariage_Jean/Sous_Dossier_Coulisses/sub_img.png" in b2_keys
        print("   -> OK : Parcours récursif, tri naturel et filtrage des extensions validés.")

        print("2. Test de la gestion du fichier Excel Source (dossiers.xlsx) et reprise automatique...")
        source_xlsx = parent / "dossiers.xlsx"
        df_src = pd.DataFrame(
            [
                {"Dossier": "Mariage_Jean", "Statut": "Terminé"},
                {"Dossier": "Terrain_Mendong", "Statut": ""},
                {"Dossier": "Vacances_2026", "Statut": "En cours"},
            ]
        )
        df_src.to_excel(source_xlsx, index=False)

        folders = ExcelManager.read_source_folders(str(source_xlsx))
        assert len(folders) == 3
        assert folders[0]["dossier"] == "Mariage_Jean" and folders[0]["completed"] is True
        assert folders[1]["dossier"] == "Terrain_Mendong" and folders[1]["completed"] is False
        assert folders[2]["dossier"] == "Vacances_2026" and folders[2]["completed"] is False

        # Mise à jour du statut de Terrain_Mendong -> Terminé
        ExcelManager.update_source_status(str(source_xlsx), folders[1]["row_index"], "Terrain_Mendong", "Terminé")
        folders_after = ExcelManager.read_source_folders(str(source_xlsx))
        assert folders_after[1]["completed"] is True
        print("   -> OK : Lecture et mise à jour de 'Statut = Terminé' validées.")

        print("3. Test de la génération du fichier Excel de Sortie (resultats_backblaze.xlsx)...")
        output_xlsx = parent / "resultats_backblaze.xlsx"
        fake_img_links = [f"https://f000.backblazeb2.com/file/bucket/Terrain_Mendong/img{i}.jpg" for i in range(1, 11)]
        fake_vid_links = [f"https://f000.backblazeb2.com/file/bucket/Terrain_Mendong/vid{i}.mp4" for i in range(1, 4)]

        ExcelManager.upsert_output_result(
            output_path=str(output_xlsx),
            dossier_name="Terrain_Mendong",
            image_links=fake_img_links,
            video_links=fake_vid_links,
        )

        df_out = pd.read_excel(output_xlsx, dtype=str).fillna("")
        assert list(df_out.columns) == OUTPUT_COLUMNS, f"Colonnes incorrectes : {list(df_out.columns)}"
        assert len(df_out) == 1
        assert df_out.loc[0, "Dossier"] == "Terrain_Mendong"
        assert df_out.loc[0, "Image_1"] == fake_img_links[0]
        assert df_out.loc[0, "Image_10"] == fake_img_links[9]
        assert df_out.loc[0, "Video_1"] == fake_vid_links[0]
        assert df_out.loc[0, "Video_3"] == fake_vid_links[2]
        assert df_out.loc[0, "Video_4"] == ""
        assert df_out.loc[0, "Video_5"] == ""
        print("   -> OK : Structure des 16 colonnes (Dossier + Image_1..10 + Video_1..5) validée.")

    # Création d'un fichier d'exemple dossiers.xlsx dans le répertoire du projet s'il n'existe pas
    project_dossiers = Path(__file__).resolve().parent / "dossiers.xlsx"
    if not project_dossiers.exists():
        ExcelManager.create_source_from_directory(str(project_dossiers))
        print(f"4. Modèle initial créé : {project_dossiers}")

    print("\nTOUS LES TESTS SONT PASSÉS AVEC SUCCÈS !")


if __name__ == "__main__":
    run_tests()
