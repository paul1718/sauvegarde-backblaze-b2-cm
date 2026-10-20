import os
import sys
import time
import queue
import subprocess
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional, List, Dict, Any
from PIL import Image

# Permet à l'exécutable PyInstaller (.exe) sur un PC distant (sans Python installé)
# de charger en priorité les fichiers .py mis à jour à distance depuis GitHub :
if getattr(sys, "frozen", False):
    _exe_dir = Path(sys.executable).resolve().parent
    if str(_exe_dir) not in sys.path:
        sys.path.insert(0, str(_exe_dir))
    _proj_root = _exe_dir.parent.parent
    if (_proj_root / "version.json").exists() and str(_proj_root) not in sys.path:
        sys.path.insert(0, str(_proj_root))

import customtkinter as ctk

from config_manager import ConfigManager, get_app_dir
from excel_manager import ExcelManager, ExcelPermissionError
from b2_uploader import (
    B2UploaderService,
    UploadCancelledError,
    format_size,
)
from auto_updater import AutoUpdater


# =============================================================================
# CHARTE GRAPHIQUE OFFICIELLE — SAUVEGARDE AUTOMATISÉE BACKBLAZE B2 (v1.0)
# Variables de référence (:root CSS du guide d'identité visuelle)
# =============================================================================
COLOR_NAVY = "#102A5E"             # Bleu profond : Titres, texte de marque, navigation latérale
COLOR_NAVY_LIGHT = "#193B7A"       # Survol dans la barre latérale
COLOR_PRIMARY = "#0878F9"          # Bleu principal : Boutons principaux, liens, actions
COLOR_PRIMARY_HOVER = "#0662CE"    # Survol bouton primaire
COLOR_CYAN = "#12B7F5"             # Cyan : Dégradés, accents, états actifs
COLOR_ACCENT_RED = "#E51B23"       # Rouge accent : Accent Backblaze B2, bouton d'alerte/danger
COLOR_ACCENT_RED_HOVER = "#C4141B" # Survol bouton danger
COLOR_BACKGROUND = "#FFFFFF"       # Blanc : Fond principal des cartes et espaces respirants
COLOR_SURFACE = "#F2F6FC"          # Fond clair : Arrière-plans de panneaux et zone centrale
COLOR_TEXT_SECONDARY = "#526078"   # Texte secondaire : Descriptions, métadonnées
COLOR_CARD_BORDER = "#DCE6F5"      # Bordure légère des cartes
COLOR_INPUT_BORDER = "#CBD5E1"     # Bordure des champs de formulaire
COLOR_SUCCESS = "#16A34A"          # Succès
COLOR_SUCCESS_BG = "#DCFCE7"
COLOR_WARNING = "#D97706"          # En cours / Avertissement
COLOR_WARNING_BG = "#FEF3C7"
COLOR_DANGER_BG = "#FEE2E2"

RADIUS_CARD = 10                   # --radius-card: 10px
FONT_FAMILY = "Segoe UI"           # Sans-serif nette compatible Windows (équivalent Inter / Poppins)

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


def find_asset_path(filename: str) -> Optional[Path]:
    """Recherche une ressource graphique dans `assets/` ou `pack_identite_sauvegarde_backblaze_b2/`."""
    base_dirs = [
        get_app_dir() / "assets",
        get_app_dir() / "pack_identite_sauvegarde_backblaze_b2",
        Path(__file__).resolve().parent / "assets",
        Path(__file__).resolve().parent / "pack_identite_sauvegarde_backblaze_b2",
    ]
    if getattr(sys, "frozen", False):
        exe_parent = Path(sys.executable).resolve().parent
        base_dirs.append(exe_parent / "assets")
        base_dirs.append(exe_parent / "_internal" / "assets")

    for d in base_dirs:
        candidate = d / filename
        if candidate.exists():
            return candidate
    return None


class BackblazeUploaderApp(ctk.CTk):
    """
    Application Desktop de Sauvegarde Automatisée vers Backblaze B2.
    Implémente intégralement la Charte Graphique et le Pack d'Identité Visuelle v1.0.
    """

    def __init__(self) -> None:
        super().__init__()

        self.updater = AutoUpdater()
        self.current_version = self.updater.get_local_version()

        self.title(f"Sauvegarde Automatisée Backblaze B2 — v{self.current_version}")
        self.geometry("1220x840")
        self.minsize(1060, 740)
        self.configure(fg_color=COLOR_SURFACE)

        # Chargement de l'icône officielle de la fenêtre
        for ico_name in ("icone_application_b2.ico", "app_icon.ico"):
            ico_path = find_asset_path(ico_name)
            if ico_path:
                try:
                    self.iconbitmap(str(ico_path))
                    break
                except Exception:
                    pass

        self.config_manager = ConfigManager()
        self.saved_config = self.config_manager.load()

        # File d'attente thread-safe pour les mises à jour de l'interface
        self.ui_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()

        # Événements de contrôle du thread de téléversement
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.stop_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self.is_running = False

        # Fichier log disque
        self.log_file_path = get_app_dir() / "sauvegarde_b2.log"

        # Images de marque (conservées en attributs pour éviter le garbage collection)
        self.sidebar_icon_img: Optional[ctk.CTkImage] = None
        self.header_logo_img: Optional[ctk.CTkImage] = None
        self._load_brand_images()

        # Construction de l'interface graphique selon la maquette Dashboard
        self._build_ui()
        self._load_config_into_ui()

        # Boucle de rafraîchissement de la file UI
        self.after(100, self._process_ui_queue)

        # Vérification silencieuse des mises à jour à distance au démarrage
        self.after(800, lambda: self._check_remote_updates(silent=True))

        # Vérification de reprise automatique après coupure brusque de courant
        self.after(1500, self._check_power_loss_auto_resume)

        # Fermeture propre de l'application
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_brand_images(self) -> None:
        """Charge le logo principal et l'icône en respectant strictement leurs proportions."""
        icon_file = find_asset_path("icone_128.png") or find_asset_path("icone_256.png")
        if icon_file:
            try:
                pil_icon = Image.open(icon_file).convert("RGBA")
                self.sidebar_icon_img = ctk.CTkImage(
                    light_image=pil_icon, dark_image=pil_icon, size=(44, 44)
                )
            except Exception:
                pass

        logo_file = find_asset_path("logo_principal.png")
        if logo_file:
            try:
                pil_logo = Image.open(logo_file).convert("RGBA")
                w, h = pil_logo.size
                target_h = 56
                target_w = int(w * (target_h / max(h, 1)))
                self.header_logo_img = ctk.CTkImage(
                    light_image=pil_logo, dark_image=pil_logo, size=(target_w, target_h)
                )
            except Exception:
                pass

    def _build_ui(self) -> None:
        """Construit la structure : Barre latérale Navy (#102A5E) + Zone Centrale Dashboard (#F2F6FC)."""
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # =====================================================================
        # BARRE LATÉRALE GAUCHE (Navigation & Identité — #102A5E)
        # =====================================================================
        self.sidebar = ctk.CTkFrame(
            self,
            width=250,
            corner_radius=0,
            fg_color=COLOR_NAVY,
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)
        self.sidebar.grid_rowconfigure(6, weight=1)

        # En-tête de marque dans la barre latérale
        brand_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand_frame.grid(row=0, column=0, padx=16, pady=(20, 16), sticky="ew")

        if self.sidebar_icon_img:
            icon_lbl = ctk.CTkLabel(brand_frame, text="", image=self.sidebar_icon_img)
            icon_lbl.grid(row=0, column=0, rowspan=2, padx=(0, 10), sticky="w")

        ctk.CTkLabel(
            brand_frame,
            text="Sauvegarde Automatisée",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLOR_BACKGROUND,
            anchor="w",
        ).grid(row=0, column=1, sticky="w")

        ctk.CTkLabel(
            brand_frame,
            text="Backblaze B2",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLOR_CYAN,
            anchor="w",
        ).grid(row=1, column=1, sticky="w")

        # Séparateur subtil
        sep = ctk.CTkFrame(self.sidebar, height=1, fg_color="#23427D")
        sep.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="ew")

        # Boutons de navigation latérale (conformes à la maquette Dashboard)
        self.nav_buttons: Dict[str, ctk.CTkButton] = {}
        nav_items = [
            ("dashboard", "🏠   Tableau de bord"),
            ("sauvegardes", "☁️   Sauvegardes & Fichiers"),
            ("parametres", "⚙️   Paramètres Backblaze B2"),
            ("historique", "📜   Historique & Journaux"),
        ]
        for idx, (nav_key, nav_label) in enumerate(nav_items, start=2):
            btn = ctk.CTkButton(
                self.sidebar,
                text=nav_label,
                anchor="w",
                height=40,
                corner_radius=8,
                font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold" if nav_key == "dashboard" else "normal"),
                fg_color=COLOR_PRIMARY if nav_key == "dashboard" else "transparent",
                text_color=COLOR_BACKGROUND,
                hover_color=COLOR_PRIMARY_HOVER if nav_key == "dashboard" else COLOR_NAVY_LIGHT,
                command=lambda k=nav_key: self._switch_view(k),
            )
            btn.grid(row=idx, column=0, padx=14, pady=4, sticky="ew")
            self.nav_buttons[nav_key] = btn

        # Encart d'information "Protection Active" en bas de la barre latérale
        info_card = ctk.CTkFrame(self.sidebar, fg_color="#163570", corner_radius=10)
        info_card.grid(row=7, column=0, padx=14, pady=8, sticky="ew")

        ctk.CTkLabel(
            info_card,
            text="🛡️ Résilience Active",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_CYAN,
        ).pack(anchor="w", padx=12, pady=(10, 2))

        ctk.CTkLabel(
            info_card,
            text="• Anti-corruption Excel (.bak)\n• Pause auto si coupure Internet\n• Reprise auto après coupure",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color="#D0E1FD",
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 10))

        # Pied de barre latérale : Mise à jour à distance & Version
        footer_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        footer_frame.grid(row=8, column=0, padx=14, pady=(4, 16), sticky="ew")
        footer_frame.grid_columnconfigure(0, weight=1)

        self.update_btn = ctk.CTkButton(
            footer_frame,
            text="🔄 Vérifier mise à jour",
            height=34,
            corner_radius=8,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color="#1D4184",
            text_color=COLOR_BACKGROUND,
            hover_color=COLOR_PRIMARY,
            command=lambda: self._check_remote_updates(silent=False),
        )
        self.update_btn.grid(row=0, column=0, pady=(0, 6), sticky="ew")

        ctk.CTkLabel(
            footer_frame,
            text=f"Version {self.current_version} • Charte v1.0",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color="#8CA5D6",
        ).grid(row=1, column=0, sticky="w")

        # =====================================================================
        # ZONE PRINCIPALE DROITE (Fond clair #F2F6FC)
        # =====================================================================
        self.main_area = ctk.CTkFrame(self, fg_color=COLOR_SURFACE, corner_radius=0)
        self.main_area.grid(row=0, column=1, sticky="nsew", padx=20, pady=16)
        self.main_area.grid_columnconfigure(0, weight=1)
        self.main_area.grid_rowconfigure(4, weight=1)

        # ---------------------------------------------------------------------
        # 0. Barre supérieure avec Logo officiel, message d'accueil et état global
        # ---------------------------------------------------------------------
        top_bar = ctk.CTkFrame(
            self.main_area,
            fg_color=COLOR_BACKGROUND,
            corner_radius=RADIUS_CARD,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
        )
        top_bar.grid(row=0, column=0, pady=(0, 10), sticky="ew")
        top_bar.grid_columnconfigure(1, weight=1)

        if self.header_logo_img:
            logo_lbl = ctk.CTkLabel(top_bar, text="", image=self.header_logo_img)
            logo_lbl.grid(row=0, column=0, rowspan=2, padx=(16, 12), pady=10, sticky="w")

        welcome_title = ctk.CTkLabel(
            top_bar,
            text="Bonjour ! 👋  Vos fichiers, sauvegardés automatiquement.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=18, weight="bold"),
            text_color=COLOR_NAVY,
        )
        welcome_title.grid(row=0, column=1, padx=4, pady=(10, 0), sticky="w")

        welcome_sub = ctk.CTkLabel(
            top_bar,
            text="Vos données sont en sécurité avec Backblaze B2 — Suivi clair et reprise automatique des opérations.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=COLOR_TEXT_SECONDARY,
        )
        welcome_sub.grid(row=1, column=1, padx=4, pady=(0, 10), sticky="w")

        # Badge d'état global dans la barre supérieure
        self.global_badge = ctk.CTkLabel(
            top_bar,
            text="  ● Système prêt  ",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=COLOR_SUCCESS_BG,
            text_color=COLOR_SUCCESS,
            corner_radius=14,
            height=30,
        )
        self.global_badge.grid(row=0, column=2, rowspan=2, padx=16, pady=10, sticky="e")

        # ---------------------------------------------------------------------
        # 1. Rangée des 4 Cartes KPI (comme sur la Maquette Dashboard)
        # ---------------------------------------------------------------------
        self.kpi_row = ctk.CTkFrame(self.main_area, fg_color="transparent")
        self.kpi_row.grid(row=1, column=0, pady=(0, 10), sticky="ew")
        self.kpi_row.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.kpi_state_val = self._create_kpi_card(
            self.kpi_row, 0, "✔", "#DCFCE7", COLOR_SUCCESS, "Sauvegarde automatique", "● Prêt à démarrer"
        )
        self.kpi_volume_val = self._create_kpi_card(
            self.kpi_row, 1, "☁️", "#E0F2FE", COLOR_PRIMARY, "Volume transféré", "0 o / 0 o"
        )
        self.kpi_folders_val = self._create_kpi_card(
            self.kpi_row, 2, "📁", "#EDE9FE", "#6D28D9", "Dossiers traités", "0 / 0 dossiers"
        )
        self.kpi_files_val = self._create_kpi_card(
            self.kpi_row, 3, "🛡️", "#E0F7FA", COLOR_CYAN, "Fichiers sauvegardés", "0 / 0 fichiers"
        )

        # ---------------------------------------------------------------------
        # 2. Grille centrale : Carte Configuration B2 + Carte Sélection des Fichiers
        # ---------------------------------------------------------------------
        self.forms_row = ctk.CTkFrame(self.main_area, fg_color="transparent")
        self.forms_row.grid(row=2, column=0, pady=(0, 10), sticky="ew")
        self.forms_row.grid_columnconfigure(0, weight=5)
        self.forms_row.grid_columnconfigure(1, weight=7)

        # --- CARTE A : Configuration & Identifiants Backblaze B2 ---
        self.card_b2 = ctk.CTkFrame(
            self.forms_row,
            fg_color=COLOR_BACKGROUND,
            corner_radius=RADIUS_CARD,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
        )
        self.card_b2.grid(row=0, column=0, padx=(0, 6), sticky="nsew")
        self.card_b2.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self.card_b2,
            text="🔑  Identifiants Backblaze B2",
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=0, column=0, columnspan=3, padx=16, pady=(12, 8), sticky="w")

        # B2 Key ID
        ctk.CTkLabel(
            self.card_b2,
            text="B2 Key ID",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=1, column=0, padx=(16, 8), pady=5, sticky="w")
        self.key_id_var = ctk.StringVar()
        self.key_id_entry = ctk.CTkEntry(
            self.card_b2,
            textvariable=self.key_id_var,
            placeholder_text="Ex: 005a1b2c3d4e5f...",
            fg_color=COLOR_SURFACE,
            border_color=COLOR_INPUT_BORDER,
            text_color=COLOR_NAVY,
            height=32,
        )
        self.key_id_entry.grid(row=1, column=1, columnspan=2, padx=(0, 16), pady=5, sticky="ew")

        # B2 Application Key
        ctk.CTkLabel(
            self.card_b2,
            text="Application Key",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=2, column=0, padx=(16, 8), pady=5, sticky="w")
        self.app_key_var = ctk.StringVar()
        self.app_key_entry = ctk.CTkEntry(
            self.card_b2,
            textvariable=self.app_key_var,
            show="•",
            placeholder_text="Clé secrète d'application B2",
            fg_color=COLOR_SURFACE,
            border_color=COLOR_INPUT_BORDER,
            text_color=COLOR_NAVY,
            height=32,
        )
        self.app_key_entry.grid(row=2, column=1, padx=(0, 6), pady=5, sticky="ew")

        self.show_key_btn = ctk.CTkButton(
            self.card_b2,
            text="Afficher",
            width=74,
            height=32,
            corner_radius=8,
            fg_color=COLOR_BACKGROUND,
            border_width=1,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=self._toggle_app_key_visibility,
        )
        self.show_key_btn.grid(row=2, column=2, padx=(0, 16), pady=5, sticky="e")

        # Bucket Name + Bouton Tester
        ctk.CTkLabel(
            self.card_b2,
            text="Bucket Name",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=3, column=0, padx=(16, 8), pady=(5, 14), sticky="w")
        self.bucket_var = ctk.StringVar()
        self.bucket_entry = ctk.CTkEntry(
            self.card_b2,
            textvariable=self.bucket_var,
            placeholder_text="Nom du Bucket B2",
            fg_color=COLOR_SURFACE,
            border_color=COLOR_INPUT_BORDER,
            text_color=COLOR_NAVY,
            height=32,
        )
        self.bucket_entry.grid(row=3, column=1, padx=(0, 6), pady=(5, 14), sticky="ew")

        self.test_conn_btn = ctk.CTkButton(
            self.card_b2,
            text="Tester B2",
            width=96,
            height=32,
            corner_radius=8,
            fg_color=COLOR_BACKGROUND,
            border_width=1,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=self._test_b2_connection,
        )
        self.test_conn_btn.grid(row=3, column=2, padx=(0, 16), pady=(5, 14), sticky="e")

        # --- CARTE B : Sélection des Fichiers et Répertoires ---
        self.card_files = ctk.CTkFrame(
            self.forms_row,
            fg_color=COLOR_BACKGROUND,
            corner_radius=RADIUS_CARD,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
        )
        self.card_files.grid(row=0, column=1, padx=(6, 0), sticky="nsew")
        self.card_files.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self.card_files,
            text="📂  Sources & Fichiers Excel",
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=0, column=0, columnspan=4, padx=16, pady=(12, 8), sticky="w")

        # Répertoire Parent
        ctk.CTkLabel(
            self.card_files,
            text="Répertoire Parent",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=1, column=0, padx=(16, 8), pady=5, sticky="w")
        self.parent_dir_var = ctk.StringVar()
        self.parent_dir_entry = ctk.CTkEntry(
            self.card_files,
            textvariable=self.parent_dir_var,
            placeholder_text="Dossier contenant les sous-dossiers à sauvegarder...",
            fg_color=COLOR_SURFACE,
            border_color=COLOR_INPUT_BORDER,
            text_color=COLOR_NAVY,
            height=32,
        )
        self.parent_dir_entry.grid(row=1, column=1, columnspan=2, padx=(0, 6), pady=5, sticky="ew")
        self.browse_parent_btn = ctk.CTkButton(
            self.card_files,
            text="Parcourir...",
            width=100,
            height=32,
            corner_radius=8,
            fg_color=COLOR_BACKGROUND,
            border_width=1,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=self._browse_parent_dir,
        )
        self.browse_parent_btn.grid(row=1, column=3, padx=(0, 16), pady=5, sticky="e")

        # Excel Source (dossiers.xlsx)
        ctk.CTkLabel(
            self.card_files,
            text="Excel Source",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=2, column=0, padx=(16, 8), pady=5, sticky="w")
        self.source_excel_var = ctk.StringVar()
        self.source_excel_entry = ctk.CTkEntry(
            self.card_files,
            textvariable=self.source_excel_var,
            placeholder_text="Chemin de dossiers.xlsx...",
            fg_color=COLOR_SURFACE,
            border_color=COLOR_INPUT_BORDER,
            text_color=COLOR_NAVY,
            height=32,
        )
        self.source_excel_entry.grid(row=2, column=1, padx=(0, 6), pady=5, sticky="ew")

        self.gen_source_btn = ctk.CTkButton(
            self.card_files,
            text="Créer modèle",
            width=98,
            height=32,
            corner_radius=8,
            fg_color=COLOR_SURFACE,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_SECONDARY,
            hover_color="#E2E8F0",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            command=self._generate_template_excel,
        )
        self.gen_source_btn.grid(row=2, column=2, padx=(0, 6), pady=5, sticky="e")

        self.browse_source_btn = ctk.CTkButton(
            self.card_files,
            text="Parcourir...",
            width=100,
            height=32,
            corner_radius=8,
            fg_color=COLOR_BACKGROUND,
            border_width=1,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=self._browse_source_excel,
        )
        self.browse_source_btn.grid(row=2, column=3, padx=(0, 16), pady=5, sticky="e")

        # Excel Sortie (resultats_backblaze.xlsx)
        ctk.CTkLabel(
            self.card_files,
            text="Excel Sortie",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=3, column=0, padx=(16, 8), pady=(5, 14), sticky="w")
        self.output_excel_var = ctk.StringVar()
        self.output_excel_entry = ctk.CTkEntry(
            self.card_files,
            textvariable=self.output_excel_var,
            placeholder_text="Emplacement de resultats_backblaze.xlsx...",
            fg_color=COLOR_SURFACE,
            border_color=COLOR_INPUT_BORDER,
            text_color=COLOR_NAVY,
            height=32,
        )
        self.output_excel_entry.grid(row=3, column=1, columnspan=2, padx=(0, 6), pady=(5, 14), sticky="ew")
        self.browse_output_btn = ctk.CTkButton(
            self.card_files,
            text="Parcourir...",
            width=100,
            height=32,
            corner_radius=8,
            fg_color=COLOR_BACKGROUND,
            border_width=1,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=self._browse_output_excel,
        )
        self.browse_output_btn.grid(row=3, column=3, padx=(0, 16), pady=(5, 14), sticky="e")

        # ---------------------------------------------------------------------
        # 3. Carte de Contrôle, Barre de Progression & Actions (Section 6 Charte)
        # ---------------------------------------------------------------------
        self.card_control = ctk.CTkFrame(
            self.main_area,
            fg_color=COLOR_BACKGROUND,
            corner_radius=RADIUS_CARD,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
        )
        self.card_control.grid(row=3, column=0, pady=(0, 10), sticky="ew")
        self.card_control.grid_columnconfigure(0, weight=1)

        btn_bar = ctk.CTkFrame(self.card_control, fg_color="transparent")
        btn_bar.grid(row=0, column=0, padx=16, pady=(12, 8), sticky="ew")

        # Bouton Primaire : Bleu principal (#0878F9), texte blanc, libellé explicite
        self.start_btn = ctk.CTkButton(
            btn_bar,
            text="☁️  Lancer une sauvegarde",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            height=40,
            width=225,
            corner_radius=20,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_PRIMARY_HOVER,
            text_color=COLOR_BACKGROUND,
            command=self._start_backup,
        )
        self.start_btn.pack(side="left", padx=(0, 10))

        # Boutons Secondaires : Fond blanc (#FFFFFF), bordure bleu pâle (#0878F9), texte bleu
        self.pause_btn = ctk.CTkButton(
            btn_bar,
            text="⏸  Pause",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            height=40,
            width=125,
            corner_radius=20,
            fg_color=COLOR_BACKGROUND,
            border_width=1.5,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            state="disabled",
            command=self._pause_backup,
        )
        self.pause_btn.pack(side="left", padx=5)

        self.resume_btn = ctk.CTkButton(
            btn_bar,
            text="⏯  Reprendre",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            height=40,
            width=135,
            corner_radius=20,
            fg_color=COLOR_BACKGROUND,
            border_width=1.5,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            state="disabled",
            command=self._resume_backup,
        )
        self.resume_btn.pack(side="left", padx=5)

        # Bouton d'Alerte / Danger : Rouge accent (#E51B23), texte blanc, avec confirmation
        self.stop_btn = ctk.CTkButton(
            btn_bar,
            text="⚠️  Interrompre",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            height=40,
            width=145,
            corner_radius=20,
            fg_color=COLOR_ACCENT_RED,
            hover_color=COLOR_ACCENT_RED_HOVER,
            text_color=COLOR_BACKGROUND,
            state="disabled",
            command=self._stop_backup,
        )
        self.stop_btn.pack(side="left", padx=5)

        # Bouton secondaire à droite : Ouvrir Résultats Excel
        self.open_output_btn = ctk.CTkButton(
            btn_bar,
            text="📊  Ouvrir Résultats Excel",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            height=40,
            width=195,
            corner_radius=20,
            fg_color=COLOR_BACKGROUND,
            border_width=1.5,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            command=self._open_output_excel,
        )
        self.open_output_btn.pack(side="right", padx=(6, 0))

        # Bandeau de progression dans la carte de contrôle
        prog_box = ctk.CTkFrame(
            self.card_control,
            fg_color=COLOR_SURFACE,
            corner_radius=8,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
        )
        prog_box.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="ew")
        prog_box.grid_columnconfigure(0, weight=1)

        status_row = ctk.CTkFrame(prog_box, fg_color="transparent")
        status_row.grid(row=0, column=0, padx=12, pady=(8, 2), sticky="ew")
        status_row.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            status_row,
            text="Prêt à démarrer une sauvegarde vers le stockage cloud Backblaze B2.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=COLOR_NAVY,
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.percent_label = ctk.CTkLabel(
            status_row,
            text="0.0 %",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=COLOR_PRIMARY,
        )
        self.percent_label.grid(row=0, column=1, sticky="e")

        self.progress_bar = ctk.CTkProgressBar(
            prog_box,
            height=12,
            corner_radius=6,
            fg_color="#DCE6F5",
            progress_color=COLOR_PRIMARY,
        )
        self.progress_bar.grid(row=1, column=0, padx=12, pady=6, sticky="ew")
        self.progress_bar.set(0.0)

        self.Counters_label = ctk.CTkLabel(
            prog_box,
            text="Dossiers : 0 / 0   •   Fichiers : 0 / 0   •   Volume transféré : 0 o / 0 o",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=COLOR_TEXT_SECONDARY,
        )
        self.Counters_label.grid(row=2, column=0, padx=12, pady=(0, 8), sticky="w")

        # ---------------------------------------------------------------------
        # 4. Carte Journal d'Exécution & Historique des Opérations
        # ---------------------------------------------------------------------
        self.card_logs = ctk.CTkFrame(
            self.main_area,
            fg_color=COLOR_BACKGROUND,
            corner_radius=RADIUS_CARD,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
        )
        self.card_logs.grid(row=4, column=0, sticky="nsew")
        self.card_logs.grid_columnconfigure(0, weight=1)
        self.card_logs.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(self.card_logs, fg_color="transparent")
        log_header.grid(row=0, column=0, padx=16, pady=(10, 4), sticky="ew")

        ctk.CTkLabel(
            log_header,
            text="📜  Historique des opérations & Journal d'exécution en temps réel",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLOR_NAVY,
        ).pack(side="left")

        ctk.CTkButton(
            log_header,
            text="Effacer le journal",
            width=125,
            height=26,
            corner_radius=6,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            fg_color=COLOR_SURFACE,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_SECONDARY,
            hover_color="#E2E8F0",
            command=self._clear_logs,
        ).pack(side="right")

        self.log_textbox = ctk.CTkTextbox(
            self.card_logs,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=COLOR_SURFACE,
            text_color=COLOR_NAVY,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=8,
            wrap="word",
            state="disabled",
        )
        self.log_textbox.grid(row=1, column=0, padx=16, pady=(2, 14), sticky="nsew")

        # Sauvegarde automatique dès qu'un champ change
        for var in (
            self.key_id_var,
            self.app_key_var,
            self.bucket_var,
            self.parent_dir_var,
            self.source_excel_var,
            self.output_excel_var,
        ):
            var.trace_add("write", lambda *_: self._save_current_config())

    def _create_kpi_card(
        self,
        parent: ctk.CTkFrame,
        col: int,
        icon_char: str,
        badge_bg: str,
        badge_fg: str,
        title: str,
        initial_val: str,
    ) -> ctk.CTkLabel:
        """Crée une carte d'indicateur clé (KPI Card) fidèle à la Maquette Dashboard."""
        padx = (0, 6) if col == 0 else ((6, 0) if col == 3 else (6, 6))
        card = ctk.CTkFrame(
            parent,
            fg_color=COLOR_BACKGROUND,
            corner_radius=RADIUS_CARD,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
        )
        card.grid(row=0, column=col, padx=padx, sticky="ew")
        card.grid_columnconfigure(1, weight=1)

        icon_badge = ctk.CTkLabel(
            card,
            text=icon_char,
            width=40,
            height=40,
            corner_radius=20,
            fg_color=badge_bg,
            text_color=badge_fg,
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
        )
        icon_badge.grid(row=0, column=0, rowspan=2, padx=(12, 10), pady=12, sticky="w")

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=COLOR_TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=1, padx=(0, 12), pady=(10, 0), sticky="ew")

        val_lbl = ctk.CTkLabel(
            card,
            text=initial_val,
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLOR_NAVY,
            anchor="w",
        )
        val_lbl.grid(row=1, column=1, padx=(0, 12), pady=(0, 10), sticky="ew")
        return val_lbl

    def _switch_view(self, view_key: str) -> None:
        """Gère la navigation latérale de la maquette tout en gardant le travail actif."""
        for k, btn in self.nav_buttons.items():
            is_active = k == view_key
            btn.configure(
                fg_color=COLOR_PRIMARY if is_active else "transparent",
                hover_color=COLOR_PRIMARY_HOVER if is_active else COLOR_NAVY_LIGHT,
                font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold" if is_active else "normal"),
            )

        if view_key == "dashboard":
            self.kpi_row.grid()
            self.forms_row.grid()
            self.card_b2.grid()
            self.card_files.grid()
            self.card_control.grid()
            self.card_logs.grid()
        elif view_key == "sauvegardes":
            self.kpi_row.grid()
            self.forms_row.grid()
            self.card_b2.grid_remove()
            self.card_files.grid()
            self.card_control.grid()
            self.card_logs.grid()
        elif view_key == "parametres":
            self.kpi_row.grid()
            self.forms_row.grid()
            self.card_b2.grid()
            self.card_files.grid()
            self.card_control.grid()
            self.card_logs.grid()
        elif view_key == "historique":
            self.kpi_row.grid()
            self.forms_row.grid_remove()
            self.card_control.grid()
            self.card_logs.grid()

    def _update_global_badge(self, text: str, mode: str = "ready") -> None:
        """Met à jour le badge d'état en haut à droite avec texte + couleur (Règle 6 Charte)."""
        if mode == "running":
            self.global_badge.configure(
                text=f"  ⟳ {text}  ", fg_color="#E0F2FE", text_color=COLOR_PRIMARY
            )
            self.kpi_state_val.configure(text="● Active (En cours)", text_color=COLOR_PRIMARY)
        elif mode == "success":
            self.global_badge.configure(
                text=f"  ✔ {text}  ", fg_color=COLOR_SUCCESS_BG, text_color=COLOR_SUCCESS
            )
            self.kpi_state_val.configure(text="✔ Sauvegarde réussie", text_color=COLOR_SUCCESS)
        elif mode == "warning":
            self.global_badge.configure(
                text=f"  ⚠ {text}  ", fg_color=COLOR_WARNING_BG, text_color=COLOR_WARNING
            )
            self.kpi_state_val.configure(text=f"⚠ {text}", text_color=COLOR_WARNING)
        elif mode == "error":
            self.global_badge.configure(
                text=f"  ✖ {text}  ", fg_color=COLOR_DANGER_BG, text_color=COLOR_ACCENT_RED
            )
            self.kpi_state_val.configure(text="✖ Action requise", text_color=COLOR_ACCENT_RED)
        else:
            self.global_badge.configure(
                text=f"  ● {text}  ", fg_color=COLOR_SUCCESS_BG, text_color=COLOR_SUCCESS
            )
            self.kpi_state_val.configure(text="● Prêt à démarrer", text_color=COLOR_NAVY)

    def _load_config_into_ui(self) -> None:
        """Pré-remplit les champs avec les dernières valeurs mémorisées."""
        self.key_id_var.set(self.saved_config.get("b2_key_id", ""))
        self.app_key_var.set(self.saved_config.get("b2_app_key", ""))
        self.bucket_var.set(self.saved_config.get("bucket_name", ""))
        self.parent_dir_var.set(self.saved_config.get("parent_dir", ""))
        self.source_excel_var.set(self.saved_config.get("source_excel", ""))
        self.output_excel_var.set(self.saved_config.get("output_excel", ""))
        self._append_log(
            f"Application v{self.current_version} prête. Configuration simple et reprise après échec activées.",
            "INFO",
        )

    def _save_current_config(self) -> None:
        """Sauvegarde atomique des valeurs actuelles dans config.json."""
        data = {
            "b2_key_id": self.key_id_var.get().strip(),
            "b2_app_key": self.app_key_var.get().strip(),
            "bucket_name": self.bucket_var.get().strip(),
            "parent_dir": self.parent_dir_var.get().strip(),
            "source_excel": self.source_excel_var.get().strip(),
            "output_excel": self.output_excel_var.get().strip(),
            "auto_resume_after_power_loss": True,
        }
        self.config_manager.save(data)

    def _toggle_app_key_visibility(self) -> None:
        if self.app_key_entry.cget("show") == "•":
            self.app_key_entry.configure(show="")
            self.show_key_btn.configure(text="Masquer")
        else:
            self.app_key_entry.configure(show="•")
            self.show_key_btn.configure(text="Afficher")

    # =========================================================================
    # Reprise automatique après coupure brusque d'électricité
    # =========================================================================
    def _check_power_loss_auto_resume(self) -> None:
        state = self.config_manager.load_state_journal()
        if not state.get("active_session", False):
            return

        key_id = self.key_id_var.get().strip()
        app_key = self.app_key_var.get().strip()
        bucket = self.bucket_var.get().strip()
        parent_dir = self.parent_dir_var.get().strip()
        source_excel = self.source_excel_var.get().strip()

        if key_id and app_key and bucket and os.path.isdir(parent_dir) and os.path.isfile(source_excel):
            last_folder = state.get("current_folder", "")
            self._append_log(
                f"⚡ Reprise après échec électrique détectée (dernier dossier : '{last_folder}'). "
                "Reprise automatique de la sauvegarde...",
                "WARN",
            )
            self._update_global_badge("Reprise auto", "warning")
            self.status_label.configure(text="⚡ Reprise automatique suite à une coupure de courant...")
            self.after(1200, self._start_backup)

    # =========================================================================
    # Mise à jour à distance via GitHub (pour postes distants)
    # =========================================================================
    def _check_remote_updates(self, silent: bool = False) -> None:
        if not silent:
            self.update_btn.configure(state="disabled", text="Vérification...")

        def _worker() -> None:
            res = self.updater.check_for_updates()
            self.ui_queue.put({"type": "update_check_result", "result": res, "silent": silent})

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_remote_update_now(self, files: List[str], remote_version: str) -> None:
        self.update_btn.configure(state="disabled", text="Mise à jour...")
        self._append_log(f"🔄 Téléchargement de la mise à jour distante v{remote_version} depuis GitHub...", "INFO")

        def _worker() -> None:
            ok, msg = self.updater.apply_update(files)
            self.ui_queue.put(
                {
                    "type": "update_applied",
                    "success": ok,
                    "message": msg,
                    "remote_version": remote_version,
                }
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _restart_application(self) -> None:
        self._save_current_config()
        try:
            app_py = get_app_dir() / "app.py"
            if app_py.exists():
                pythonw = Path(sys.executable).with_name("pythonw.exe")
                runner = str(pythonw) if pythonw.exists() else sys.executable
                subprocess.Popen([runner, str(app_py)], cwd=str(get_app_dir()))
            else:
                subprocess.Popen([sys.executable] + sys.argv, cwd=str(get_app_dir()))
        except Exception:
            pass
        self.destroy()

    # =========================================================================
    # Sélecteurs de fichiers / répertoires
    # =========================================================================
    def _browse_parent_dir(self) -> None:
        initial = self.parent_dir_var.get().strip()
        if not initial or not os.path.isdir(initial):
            initial = str(get_app_dir())
        chosen = filedialog.askdirectory(
            title="Sélectionner le Répertoire Parent contenant les dossiers à sauvegarder",
            initialdir=initial,
        )
        if chosen:
            self.parent_dir_var.set(os.path.normpath(chosen))

    def _browse_source_excel(self) -> None:
        initial_file = self.source_excel_var.get().strip()
        initial_dir = str(Path(initial_file).parent) if initial_file else str(get_app_dir())
        chosen = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel Source (dossiers.xlsx)",
            initialdir=initial_dir,
            filetypes=[("Fichiers Excel", "*.xlsx")],
        )
        if chosen:
            self.source_excel_var.set(os.path.normpath(chosen))

    def _browse_output_excel(self) -> None:
        initial_file = self.output_excel_var.get().strip()
        initial_dir = str(Path(initial_file).parent) if initial_file else str(get_app_dir())
        chosen = filedialog.asksaveasfilename(
            title="Sélectionner l'emplacement du fichier Excel de Sortie (resultats_backblaze.xlsx)",
            initialdir=initial_dir,
            initialfile="resultats_backblaze.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
        )
        if chosen:
            self.output_excel_var.set(os.path.normpath(chosen))

    def _generate_template_excel(self) -> None:
        parent_dir = self.parent_dir_var.get().strip()
        target_excel = self.source_excel_var.get().strip()
        if not target_excel:
            target_excel = str(get_app_dir() / "dossiers.xlsx")

        if os.path.exists(target_excel):
            resp = messagebox.askyesno(
                "Confirmation",
                f"Le fichier '{Path(target_excel).name}' existe déjà.\n\n"
                "Voulez-vous le remplacer par une nouvelle liste de dossiers issue du Répertoire Parent ?",
            )
            if not resp:
                return

        try:
            count = ExcelManager.create_source_from_directory(target_excel, parent_dir if parent_dir else None)
            self.source_excel_var.set(os.path.normpath(target_excel))
            if count > 0:
                self._append_log(
                    f"Fichier source '{target_excel}' créé avec {count} sous-dossier(s) détecté(s).",
                    "SUCCESS",
                )
                messagebox.showinfo(
                    "Modèle Excel créé",
                    f"Le fichier '{Path(target_excel).name}' a été créé avec succès ({count} dossiers ajoutés).",
                )
            else:
                self._append_log(f"Modèle Excel vide créé : '{target_excel}'.", "INFO")
                messagebox.showinfo(
                    "Modèle Excel créé",
                    f"Le modèle '{Path(target_excel).name}' a été créé.\n"
                    "Astuce : si vous sélectionnez d'abord un Répertoire Parent, ses sous-dossiers y seront listés automatiquement.",
                )
        except Exception as e:
            self._append_log(f"Erreur lors de la création du modèle Excel : {e}", "ERROR")
            messagebox.showerror("Erreur", f"Impossible de créer le fichier Excel :\n{e}")

    def _open_output_excel(self) -> None:
        output_path = self.output_excel_var.get().strip()
        if not output_path or not os.path.exists(output_path):
            messagebox.showinfo(
                "Fichier non disponible",
                "Le fichier Excel de sortie n'existe pas encore. Il sera généré dès le premier fichier média sauvegardé.",
            )
            return
        try:
            os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier :\n{e}")

    # =========================================================================
    # Test de connexion Backblaze B2
    # =========================================================================
    def _test_b2_connection(self) -> None:
        key_id = self.key_id_var.get().strip()
        app_key = self.app_key_var.get().strip()
        bucket = self.bucket_var.get().strip()

        if not key_id or not app_key or not bucket:
            messagebox.showwarning(
                "Paramètres manquants",
                "Veuillez renseigner le 'B2 Key ID', la 'B2 Application Key' et le 'Bucket Name'.",
            )
            return

        self._save_current_config()
        self.test_conn_btn.configure(state="disabled", text="Test...")
        self._append_log(f"Vérification de la connexion au stockage cloud Backblaze B2 ('{bucket}')...", "INFO")

        def _worker() -> None:
            try:
                service = B2UploaderService(key_id, app_key, bucket)
                service.connect(max_auth_retries=1)
                self.ui_queue.put(
                    {
                        "type": "conn_test_result",
                        "success": True,
                        "msg": f"Connexion au bucket '{bucket}' vérifiée avec succès.",
                    }
                )
            except Exception as e:
                self.ui_queue.put(
                    {
                        "type": "conn_test_result",
                        "success": False,
                        "msg": f"Connexion au stockage indisponible. Vérifiez les paramètres puis réessayez.\nDétail : {e}",
                    }
                )

        threading.Thread(target=_worker, daemon=True).start()

    # =========================================================================
    # Contrôles : Lancer / Pause / Reprendre / Interrompre
    # =========================================================================
    def _start_backup(self) -> None:
        if self.is_running:
            return

        key_id = self.key_id_var.get().strip()
        app_key = self.app_key_var.get().strip()
        bucket_name = self.bucket_var.get().strip()
        parent_dir = self.parent_dir_var.get().strip()
        source_excel = self.source_excel_var.get().strip()
        output_excel = self.output_excel_var.get().strip()

        if not key_id or not app_key or not bucket_name:
            messagebox.showwarning(
                "Identifiants Backblaze B2 requis",
                "Veuillez renseigner vos identifiants Backblaze B2 (Key ID, Application Key et Bucket Name).",
            )
            return

        if not parent_dir or not os.path.isdir(parent_dir):
            messagebox.showwarning(
                "Répertoire Parent invalide",
                "Veuillez sélectionner un Répertoire Parent existant sur votre ordinateur.",
            )
            return

        if not source_excel or not os.path.isfile(source_excel):
            messagebox.showwarning(
                "Fichier Excel Source introuvable",
                "Veuillez sélectionner un fichier Excel source (dossiers.xlsx) existant.",
            )
            return

        if not output_excel:
            output_excel = str(Path(source_excel).parent / "resultats_backblaze.xlsx")
            self.output_excel_var.set(output_excel)

        self._save_current_config()
        self.config_manager.mark_session_active(True)

        self.is_running = True
        self.pause_event.set()
        self.stop_event.clear()

        self._set_controls_state(running=True, paused=False)
        self._update_global_badge("Sauvegarde en cours", "running")
        self.status_label.configure(text="Initialisation du transfert vers le stockage cloud...")

        self.worker_thread = threading.Thread(
            target=self._backup_worker,
            args=(key_id, app_key, bucket_name, parent_dir, source_excel, output_excel),
            daemon=True,
        )
        self.worker_thread.start()

    def _pause_backup(self) -> None:
        if not self.is_running or not self.pause_event.is_set():
            return
        self.pause_event.clear()
        self._set_controls_state(running=True, paused=True)
        self._update_global_badge("En pause", "warning")
        self.status_label.configure(text="⏸ Sauvegarde en pause — Cliquez sur « Reprendre » pour poursuivre.")
        self._append_log("Pause demandée par l'utilisateur. Le transfert se met en pause...", "WARN")

    def _resume_backup(self) -> None:
        if not self.is_running or self.pause_event.is_set():
            return
        self.pause_event.set()
        self._set_controls_state(running=True, paused=False)
        self._update_global_badge("Sauvegarde en cours", "running")
        self.status_label.configure(text="▶ Transfert vers le stockage cloud en cours...")
        self._append_log("Reprise de la sauvegarde.", "INFO")

    def _stop_backup(self) -> None:
        if not self.is_running:
            return
        resp = messagebox.askyesno(
            "Interrompre la sauvegarde",
            "Voulez-vous interrompre la sauvegarde en cours ?\n\n"
            "Tous les fichiers déjà envoyés sont enregistrés et la sauvegarde reprendra là où elle s'est arrêtée au prochain lancement.",
        )
        if not resp:
            return
        self.config_manager.mark_session_active(False)
        self.stop_event.set()
        self.pause_event.set()
        self.stop_btn.configure(state="disabled")
        self._update_global_badge("Interruption...", "warning")
        self.status_label.configure(text="Arrêt propre en cours...")
        self._append_log("Interruption demandée par l'utilisateur. Arrêt propre en cours...", "WARN")

    def _set_controls_state(self, running: bool, paused: bool) -> None:
        if not running:
            self.start_btn.configure(state="normal")
            self.pause_btn.configure(state="disabled")
            self.resume_btn.configure(state="disabled")
            self.stop_btn.configure(state="disabled")
            self.test_conn_btn.configure(state="normal")
            self.browse_parent_btn.configure(state="normal")
            self.browse_source_btn.configure(state="normal")
            self.browse_output_btn.configure(state="normal")
            self.gen_source_btn.configure(state="normal")
        else:
            self.start_btn.configure(state="disabled")
            self.test_conn_btn.configure(state="disabled")
            self.browse_parent_btn.configure(state="disabled")
            self.browse_source_btn.configure(state="disabled")
            self.browse_output_btn.configure(state="disabled")
            self.gen_source_btn.configure(state="disabled")
            self.stop_btn.configure(state="normal")
            if paused:
                self.pause_btn.configure(state="disabled")
                self.resume_btn.configure(state="normal")
            else:
                self.pause_btn.configure(state="normal")
                self.resume_btn.configure(state="disabled")

    # =========================================================================
    # Thread principal de traitement et téléversement
    # =========================================================================
    def _backup_worker(
        self,
        key_id: str,
        app_key: str,
        bucket_name: str,
        parent_dir: str,
        source_excel: str,
        output_excel: str,
    ) -> None:
        try:
            self._queue_log("=" * 72, "INFO")
            self._queue_log("DÉMARRAGE DE LA SAUVEGARDE VERS LE STOCKAGE CLOUD BACKBLAZE B2", "INFO")
            self._queue_log(f"Répertoire Parent : {parent_dir}", "INFO")
            self._queue_log(f"Fichier Source    : {source_excel}", "INFO")
            self._queue_log(f"Fichier Sortie    : {output_excel}", "INFO")

            folders_info = ExcelManager.read_source_folders(source_excel)
            total_listed = len(folders_info)
            already_completed = [f for f in folders_info if f["completed"]]
            pending_folders = [f for f in folders_info if not f["completed"]]

            self._queue_log(
                f"Analyse de '{Path(source_excel).name}' : {total_listed} dossier(s) au total | "
                f"{len(already_completed)} déjà 'Terminé' (ignorés) | {len(pending_folders)} à traiter.",
                "INFO",
            )

            for skipped in already_completed:
                self._queue_log(f"⏭ Dossier ignoré (Statut = Terminé) : {skipped['dossier']}", "INFO")

            if not pending_folders:
                self.config_manager.mark_session_active(False)
                self._queue_progress(
                    folders_done=total_listed,
                    folders_total=total_listed,
                    files_done=0,
                    files_total=0,
                    bytes_done=0,
                    bytes_total=0,
                    status_text="✔ Sauvegarde terminée — Tous les dossiers sont à jour.",
                    fraction=1.0,
                )
                self._queue_log("Sauvegarde terminée : aucun dossier restant à traiter.", "SUCCESS")
                self.ui_queue.put({"type": "finished", "cancelled": False, "error": None})
                return

            self._queue_status("Analyse de l'arborescence et calcul du volume à sauvegarder...")
            total_files_all = 0
            total_bytes_all = 0
            valid_pending: List[Dict[str, Any]] = []

            for item in pending_folders:
                if self.stop_event.is_set():
                    raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")
                folder_path = Path(parent_dir) / item["dossier"]
                if not folder_path.exists() or not folder_path.is_dir():
                    self._queue_log(
                        f"⚠ Dossier '{item['dossier']}' introuvable dans '{parent_dir}' — ignoré.",
                        "WARN",
                    )
                    continue

                files_in_folder = B2UploaderService.scan_folder_files(folder_path)
                folder_bytes = sum(fp.stat().st_size for fp, _, _ in files_in_folder)
                item["folder_path"] = folder_path
                item["files_count"] = len(files_in_folder)
                item["folder_bytes"] = folder_bytes
                total_files_all += len(files_in_folder)
                total_bytes_all += folder_bytes
                valid_pending.append(item)

            if not valid_pending:
                self.config_manager.mark_session_active(False)
                self._queue_log(
                    "Aucun des dossiers à traiter n'a été trouvé dans le Répertoire Parent sélectionné.",
                    "WARN",
                )
                self.ui_queue.put(
                    {
                        "type": "finished",
                        "cancelled": False,
                        "error": "Aucun dossier valide trouvé dans le Répertoire Parent.",
                    }
                )
                return

            self._queue_log(
                f"Volume global à traiter : {len(valid_pending)} dossier(s), "
                f"{total_files_all} fichier(s) éligible(s), taille totale : {format_size(total_bytes_all)}.",
                "INFO",
            )

            self._queue_status(f"Connexion au bucket Backblaze B2 '{bucket_name}'...")
            uploader = B2UploaderService(
                key_id=key_id,
                app_key=app_key,
                bucket_name=bucket_name,
                log_callback=self._queue_log,
                status_callback=self._queue_status,
                pause_event=self.pause_event,
                stop_event=self.stop_event,
            )
            uploader.connect()
            self._queue_log(f"Connecté avec succès au bucket Backblaze B2 '{bucket_name}'.", "SUCCESS")

            folders_done_count = len(already_completed)
            files_done_count = 0
            bytes_done_count = 0

            state_journal = self.config_manager.load_state_journal()
            local_journal_files: Dict[str, Any] = state_journal.get("uploaded_files", {})

            self._queue_progress(
                folders_done=folders_done_count,
                folders_total=total_listed,
                files_done=files_done_count,
                files_total=total_files_all,
                bytes_done=bytes_done_count,
                bytes_total=total_bytes_all,
                status_text="Transfert vers le stockage cloud en cours...",
                fraction=0.0,
            )

            for idx, item in enumerate(valid_pending, start=1):
                dossier_name = item["dossier"]
                row_index = item["row_index"]
                folder_path: Path = item["folder_path"]

                self.config_manager.mark_session_active(True, current_folder=dossier_name)

                self._queue_log("-" * 64, "INFO")
                self._queue_log(
                    f"📁 [{idx}/{len(valid_pending)}] Traitement du dossier : '{dossier_name}' "
                    f"({item['files_count']} fichier(s), {format_size(item['folder_bytes'])})",
                    "INFO",
                )
                self._queue_status(f"En cours : {dossier_name} ({idx}/{len(valid_pending)})")

                ExcelManager.update_source_status(
                    source_path=source_excel,
                    row_index=row_index,
                    dossier_name=dossier_name,
                    status="En cours",
                )

                def _on_file_done(file_size: int, file_path: Path, b2_key: str, url: str) -> None:
                    nonlocal files_done_count, bytes_done_count
                    files_done_count += 1
                    bytes_done_count += file_size

                    local_journal_files[b2_key] = {"size": int(file_size), "url": str(url)}
                    self.config_manager.record_uploaded_file(b2_key, file_size, url)

                    if total_bytes_all > 0:
                        frac = min(bytes_done_count / total_bytes_all, 1.0)
                    elif total_files_all > 0:
                        frac = min(files_done_count / total_files_all, 1.0)
                    else:
                        frac = min((folders_done_count + 1) / max(total_listed, 1), 1.0)

                    self._queue_progress(
                        folders_done=folders_done_count,
                        folders_total=total_listed,
                        files_done=files_done_count,
                        files_total=total_files_all,
                        bytes_done=bytes_done_count,
                        bytes_total=total_bytes_all,
                        status_text=f"Transfert en cours : {dossier_name} — {file_path.name}",
                        fraction=frac,
                    )

                def _on_root_links_updated(
                    partial_imgs: List[str],
                    partial_vids: List[str],
                    file_idx_in_folder: int,
                    total_in_folder: int,
                ) -> None:
                    ExcelManager.upsert_output_result(
                        output_path=output_excel,
                        dossier_name=dossier_name,
                        image_links=partial_imgs,
                        video_links=partial_vids,
                    )

                root_images, root_videos, uploaded_count, _ = uploader.upload_folder(
                    folder_path=folder_path,
                    local_journal_files=local_journal_files,
                    on_file_uploaded=_on_file_done,
                    on_root_links_updated=_on_root_links_updated,
                )

                ExcelManager.upsert_output_result(
                    output_path=output_excel,
                    dossier_name=dossier_name,
                    image_links=root_images,
                    video_links=root_videos,
                )
                self._queue_log(
                    f"  📊 Excel de sortie consolidé pour '{dossier_name}' "
                    f"({len(root_images)} lien(s) image racine, {len(root_videos)} lien(s) vidéo racine).",
                    "INFO",
                )

                ExcelManager.update_source_status(
                    source_path=source_excel,
                    row_index=row_index,
                    dossier_name=dossier_name,
                    status="Terminé",
                )
                folders_done_count += 1

                frac = (
                    min(bytes_done_count / total_bytes_all, 1.0)
                    if total_bytes_all > 0
                    else min(folders_done_count / max(total_listed, 1), 1.0)
                )
                self._queue_progress(
                    folders_done=folders_done_count,
                    folders_total=total_listed,
                    files_done=files_done_count,
                    files_total=total_files_all,
                    bytes_done=bytes_done_count,
                    bytes_total=total_bytes_all,
                    status_text=f"✔ Dossier sauvegardé : {dossier_name}",
                    fraction=frac,
                )
                self._queue_log(
                    f"✔ Dossier '{dossier_name}' sauvegardé avec succès et marqué 'Terminé'.",
                    "SUCCESS",
                )

            self.config_manager.mark_session_active(False, current_folder="")

            self._queue_progress(
                folders_done=folders_done_count,
                folders_total=total_listed,
                files_done=files_done_count,
                files_total=total_files_all,
                bytes_done=bytes_done_count,
                bytes_total=total_bytes_all,
                status_text="✔ Sauvegarde terminée — Vos données ont été sauvegardées avec succès sur Backblaze B2.",
                fraction=1.0,
            )
            self._queue_log("=" * 72, "INFO")
            self._queue_log("✔ SAUVEGARDE TERMINÉE : TOUS LES DOSSIERS SONT SYNCHRONISÉS ET VÉRIFIABLES.", "SUCCESS")
            self.ui_queue.put({"type": "finished", "cancelled": False, "error": None})

        except UploadCancelledError as e:
            self._queue_log(f"⏹ {e}", "WARN")
            self._queue_status("⏹ Sauvegarde interrompue — Reprise automatique prévue au prochain lancement.")
            self.ui_queue.put({"type": "finished", "cancelled": True, "error": None})

        except ExcelPermissionError as e:
            self._queue_log(f"✖ Erreur d'accès Excel : {e}", "ERROR")
            self._queue_status("✖ Action requise : Fermez le fichier Excel ouvert puis réessayez.")
            self.ui_queue.put({"type": "finished", "cancelled": False, "error": str(e)})

        except Exception as e:
            self._queue_log(f"✖ Erreur : {e}", "ERROR")
            self._queue_status(f"✖ Connexion ou opération interrompue : {e}")
            self.ui_queue.put({"type": "finished", "cancelled": False, "error": str(e)})

    # =========================================================================
    # Communication Thread -> Interface Graphique (Queue UI)
    # =========================================================================
    def _queue_log(self, message: str, level: str = "INFO") -> None:
        self.ui_queue.put({"type": "log", "message": message, "level": level})

    def _queue_status(self, status_text: str) -> None:
        self.ui_queue.put({"type": "status", "text": status_text})

    def _queue_progress(
        self,
        folders_done: int,
        folders_total: int,
        files_done: int,
        files_total: int,
        bytes_done: int,
        bytes_total: int,
        status_text: str,
        fraction: float,
    ) -> None:
        self.ui_queue.put(
            {
                "type": "progress",
                "folders_done": folders_done,
                "folders_total": folders_total,
                "files_done": files_done,
                "files_total": files_total,
                "bytes_done": bytes_done,
                "bytes_total": bytes_total,
                "status_text": status_text,
                "fraction": fraction,
            }
        )

    def _process_ui_queue(self) -> None:
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                msg_type = msg.get("type")

                if msg_type == "log":
                    self._append_log(msg["message"], msg.get("level", "INFO"))

                elif msg_type == "status":
                    txt = msg["text"]
                    self.status_label.configure(text=txt)
                    if "Connexion Internet coupée" in txt or "En attente" in txt:
                        self._update_global_badge("Attente Internet", "warning")
                    elif "rétablie" in txt:
                        self._update_global_badge("Sauvegarde en cours", "running")

                elif msg_type == "progress":
                    frac = max(0.0, min(1.0, float(msg["fraction"])))
                    self.progress_bar.set(frac)
                    self.percent_label.configure(text=f"{frac * 100:.1f} %")
                    if self.pause_event.is_set():
                        self.status_label.configure(text=msg["status_text"])

                    vol_str = f"{format_size(msg['bytes_done'])} / {format_size(msg['bytes_total'])}"
                    fold_str = f"{msg['folders_done']} / {msg['folders_total']} dossiers"
                    files_str = f"{msg['files_done']} / {msg['files_total']} fichiers"

                    self.kpi_volume_val.configure(text=vol_str)
                    self.kpi_folders_val.configure(text=fold_str)
                    self.kpi_files_val.configure(text=files_str)

                    self.Counters_label.configure(
                        text=(
                            f"Dossiers : {msg['folders_done']} / {msg['folders_total']}   •   "
                            f"Fichiers : {msg['files_done']} / {msg['files_total']}   •   "
                            f"Volume transféré : {vol_str}"
                        )
                    )

                elif msg_type == "conn_test_result":
                    self.test_conn_btn.configure(state="normal", text="Tester B2")
                    if msg["success"]:
                        self._append_log(f"✔ {msg['msg']}", "SUCCESS")
                        messagebox.showinfo("Connexion Backblaze B2", msg["msg"])
                    else:
                        self._append_log(f"✖ {msg['msg']}", "ERROR")
                        messagebox.showerror("Connexion au stockage indisponible", msg["msg"])

                elif msg_type == "update_check_result":
                    self.update_btn.configure(state="normal", text="🔄 Vérifier mise à jour")
                    res = msg["result"]
                    silent = msg["silent"]
                    if res.get("update_available"):
                        remote_v = res["remote_version"]
                        desc = res.get("description", "")
                        self._append_log(
                            f"✨ Nouvelle version v{remote_v} disponible sur GitHub ! ({desc})",
                            "SUCCESS",
                        )
                        if not self.is_running:
                            self._apply_remote_update_now(res["files"], remote_v)
                        else:
                            self.update_btn.configure(
                                text=f"⬆ Installer v{remote_v}",
                                fg_color=COLOR_PRIMARY,
                                command=lambda: self._apply_remote_update_now(res["files"], remote_v),
                            )
                    elif not silent:
                        if res.get("error"):
                            messagebox.showwarning(
                                "Vérification de mise à jour",
                                f"Impossible de contacter GitHub pour vérifier les mises à jour :\n{res['error']}",
                            )
                        else:
                            messagebox.showinfo(
                                "Application à jour",
                                f"Vous utilisez déjà la dernière version (v{res['local_version']}).",
                            )

                elif msg_type == "update_applied":
                    self.update_btn.configure(state="normal", text="🔄 Vérifier mise à jour")
                    if msg["success"]:
                        rv = msg["remote_version"]
                        self._append_log(
                            f"✔ Mise à jour v{rv} installée avec succès ! Redémarrage de l'application...",
                            "SUCCESS",
                        )
                        self.after(800, self._restart_application)
                    else:
                        self._append_log(f"✖ {msg['message']}", "ERROR")
                        messagebox.showerror("Erreur de mise à jour", msg["message"])

                elif msg_type == "finished":
                    self.is_running = False
                    self._set_controls_state(running=False, paused=False)
                    error_msg = msg.get("error")
                    cancelled = msg.get("cancelled", False)
                    if error_msg:
                        self._update_global_badge("Échec / Action requise", "error")
                        messagebox.showerror("Action requise", error_msg)
                    elif cancelled:
                        self._update_global_badge("Interrompu", "warning")
                    else:
                        self._update_global_badge("Sauvegarde réussie", "success")
                        messagebox.showinfo(
                            "Sauvegarde réussie !",
                            "Vos données ont été sauvegardées avec succès sur Backblaze B2 et le fichier Excel de résultats est à jour.",
                        )

        except queue.Empty:
            pass
        finally:
            self.after(80, self._process_ui_queue)

    def _append_log(self, message: str, level: str = "INFO") -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] [{level}] {message}\n"

        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", formatted)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(formatted)
        except Exception:
            pass

    def _clear_logs(self) -> None:
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def _on_close(self) -> None:
        self._save_current_config()
        if self.is_running:
            resp = messagebox.askyesno(
                "Sauvegarde en cours",
                "Une sauvegarde est actuellement en cours.\n\n"
                "Voulez-vous vraiment quitter l'application ?\n"
                "(La progression actuelle est enregistrée et reprendra automatiquement au prochain lancement.)",
            )
            if not resp:
                return
            self.stop_event.set()
            self.pause_event.set()
        self.destroy()


if __name__ == "__main__":
    app = BackblazeUploaderApp()
    app.mainloop()
