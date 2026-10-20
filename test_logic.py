import os
import tempfile
from pathlib import Path
import pandas as pd

from config_manager import ConfigManager
from excel_manager import ExcelManager, OUTPUT_COLUMNS
from b2_uploader import B2UploaderService, is_internet_connected
from auto_updater import AutoUpdater


def run_tests() -> None:
    print("1. Test du parcours récursif et du filtrage des extensions...")
    with tempfile.TemporaryDirectory() as tmpdir:
        parent = Path(tmpdir)
        folder = parent / "Mariage_Jean"
        subfolder = folder / "Sous_Dossier_Coulisses"
        subfolder.mkdir(parents=True)

        for i in range(1, 13):
            (folder / f"img_{i}.jpg").write_bytes(b"fake_image_data")
        for i in range(1, 7):
            (folder / f"vid_{i}.mp4").write_bytes(b"fake_video_data")
        (folder / "contrat.pdf").write_bytes(b"fake_pdf")
        (folder / "notes.txt").write_bytes(b"fake_txt")
        (folder / "ignore_me.exe").write_bytes(b"fake_exe")

        (subfolder / "sub_img.png").write_bytes(b"sub_img")
        (subfolder / "sub_vid.mov").write_bytes(b"sub_vid")
        (subfolder / "sub_doc.docx").write_bytes(b"sub_doc")

        scanned = B2UploaderService.scan_folder_files(folder)
        assert len(scanned) == 23, f"Attendu 23 fichiers éligibles, obtenu {len(scanned)}"

        root_images = [p.name for p, _, is_root in scanned if is_root and p.suffix == ".jpg"]
        assert root_images[:3] == ["img_1.jpg", "img_2.jpg", "img_3.jpg"]
        assert root_images[9] == "img_10.jpg"
        print("   -> OK : Parcours récursif, tri naturel et filtrage des extensions validés.")

        print("2. Test de la protection anti-coupure de courant sur les fichiers Excel (écriture atomique + restauration .bak)...")
        source_xlsx = parent / "dossiers.xlsx"
        df_src = pd.DataFrame(
            [
                {"Dossier": "Mariage_Jean", "Statut": "Terminé"},
                {"Dossier": "Terrain_Mendong", "Statut": ""},
                {"Dossier": "Vacances_2026", "Statut": "En cours"},
            ]
        )
        df_src.to_excel(source_xlsx, index=False)

        # Mise à jour atomique qui génère aussi dossiers_backup.xlsx
        folders = ExcelManager.read_source_folders(str(source_xlsx))
        ExcelManager.update_source_status(str(source_xlsx), folders[1]["row_index"], "Terrain_Mendong", "Terminé")

        backup_source = ExcelManager._get_backup_path(source_xlsx)
        assert backup_source.exists(), "Le fichier de secours dossiers_backup.xlsx doit exister"

        # Simulation d'une coupure brusque de courant qui aurait tronqué/corrompu dossiers.xlsx
        source_xlsx.write_bytes(b"FICHIER_CORROMPU_PAR_COUPURE_COURANT")
        recovered_folders = ExcelManager.read_source_folders(str(source_xlsx))
        assert len(recovered_folders) == 3
        assert recovered_folders[1]["completed"] is True
        print("   -> OK : Auto-réparation automatique de dossiers.xlsx après corruption par coupure de courant validée !")

        print("3. Test de l'enregistrement au fil de l'eau et de l'auto-réparation de resultats_backblaze.xlsx...")
        output_xlsx = parent / "resultats_backblaze.xlsx"
        fake_img_links = [f"https://f000.backblazeb2.com/file/bucket/Terrain_Mendong/img{i}.jpg" for i in range(1, 11)]
        fake_vid_links = [f"https://f000.backblazeb2.com/file/bucket/Terrain_Mendong/vid{i}.mp4" for i in range(1, 4)]

        # Simulation d'enregistrement progressif image par image
        ExcelManager.upsert_output_result(str(output_xlsx), "Terrain_Mendong", fake_img_links[:2], [])
        ExcelManager.upsert_output_result(str(output_xlsx), "Terrain_Mendong", fake_img_links, fake_vid_links)

        # Simulation de corruption sur resultats_backblaze.xlsx suivie d'une nouvelle écriture
        output_xlsx.write_bytes(b"CORROMPU")
        df_out = ExcelManager._read_excel_dataframe_safe(output_xlsx)
        assert list(df_out.columns) == OUTPUT_COLUMNS
        assert len(df_out) == 1
        assert df_out.loc[0, "Image_10"] == fake_img_links[9]
        assert df_out.loc[0, "Video_3"] == fake_vid_links[2]
        print("   -> OK : Sauvegarde progressive et auto-réparation de resultats_backblaze.xlsx validées !")

        print("4. Test du détecteur de connexion Internet et de l'AutoUpdater...")
        net_ok = is_internet_connected()
        print(f"   -> Connectivité Internet détectée : {net_ok}")
        updater = AutoUpdater()
        print(f"   -> Version locale de l'application : {updater.get_local_version()}")

    print("\nTOUS LES TESTS SONT PASSÉS AVEC SUCCÈS !")


if __name__ == "__main__":
    run_tests()
