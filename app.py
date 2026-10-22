import os
import sys
import time
import queue
import ctypes
import subprocess
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional, List, Dict, Any
from PIL import Image, ImageTk

# Identifiant Windows unique (AppUserModelID) indispensable pour que la barre des tâches Windows
# affiche l'icône officielle de l'application au lieu de l'icône Python par défaut :
APP_USER_MODEL_ID = "BackblazeB2.SauvegardeAutomatisee.DesktopApp.v1"
if sys.platform == "win32":
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass

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
    resolve_matching_folder,
    is_tav_folder,
)
from auto_updater import AutoUpdater


# =============================================================================
# CHARTE GRAPHIQUE OFFICIELLE — SAUVEGARDE AUTOMATISÉE BACKBLAZE B2 (v1.0)
# =============================================================================
COLOR_NAVY = "#102A5E"
COLOR_NAVY_LIGHT = "#193B7A"
COLOR_PRIMARY = "#0878F9"
COLOR_PRIMARY_HOVER = "#0662CE"
COLOR_CYAN = "#12B7F5"
COLOR_ACCENT_RED = "#E51B23"
COLOR_ACCENT_RED_HOVER = "#C4141B"
COLOR_BACKGROUND = "#FFFFFF"
COLOR_SURFACE = "#F2F6FC"
COLOR_TEXT_SECONDARY = "#526078"
COLOR_CARD_BORDER = "#DCE6F5"
COLOR_INPUT_BORDER = "#CBD5E1"
COLOR_SUCCESS = "#16A34A"
COLOR_SUCCESS_BG = "#DCFCE7"
COLOR_WARNING = "#D97706"
COLOR_WARNING_BG = "#FEF3C7"
COLOR_DANGER_BG = "#FEE2E2"

RADIUS_CARD = 10
FONT_FAMILY = "Segoe UI"

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
    Utilise un fichier Excel unique dans lequel sont mis à jour les statuts
    et les liens des images et vidéos dans les colonnes suivantes.
    """

    def __init__(self) -> None:
        super().__init__()

        self.updater = AutoUpdater()
        self.current_version = self.updater.get_local_version()

        self.title(f"Sauvegarde Automatisée Backblaze B2 — v{self.current_version}")
        self.geometry("1220x840")
        self.minsize(1060, 740)
        self.configure(fg_color=COLOR_SURFACE)

        # Chargement et application forcée de l'icône officielle partout
        self._tk_icons: List[ImageTk.PhotoImage] = []
        self._apply_window_icon()
        self.after(210, self._apply_window_icon)
        self.after(550, self._apply_window_icon)
        self.after(1200, self._apply_window_icon)

        self.config_manager = ConfigManager()
        self.config_manager.ensure_windows_run_registry()
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

        # Images de marque
        self.sidebar_icon_img: Optional[ctk.CTkImage] = None
        self.header_logo_img: Optional[ctk.CTkImage] = None
        self._load_brand_images()

        # Construction de l'interface graphique
        self._build_ui()
        self._load_config_into_ui()

        # Boucle de rafraîchissement de la file UI
        self.after(100, self._process_ui_queue)

        # Vérification prioritaire de reprise automatique après coupure brusque de courant
        self.after(500, self._check_power_loss_auto_resume)

        # Vérification silencieuse des mises à jour à distance au démarrage
        self.after(2200, lambda: self._check_remote_updates(silent=True))

        # Fermeture propre de l'application
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _apply_window_icon(self) -> None:
        ico_path = find_asset_path("icone_application_b2.ico") or find_asset_path("app_icon.ico")
        if ico_path:
            try:
                self.iconbitmap(default=str(ico_path))
                self.iconbitmap(str(ico_path))
            except Exception:
                pass

        if not self._tk_icons:
            for png_name in ("icone_256.png", "icone_128.png", "icone_64.png", "icone_32.png"):
                png_path = find_asset_path(png_name)
                if png_path:
                    try:
                        pil_img = Image.open(png_path).convert("RGBA")
                        self._tk_icons.append(ImageTk.PhotoImage(pil_img))
                    except Exception:
                        pass

        if self._tk_icons:
            try:
                self.wm_iconphoto(True, *self._tk_icons)
            except Exception:
                pass

        if sys.platform == "win32" and ico_path:
            try:
                self.update_idletasks()
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id()) or self.winfo_id()
                IMAGE_ICON = 1
                LR_LOADFROMFILE = 0x00000010
                LR_DEFAULTSIZE = 0x00000040
                WM_SETICON = 0x0080
                ICON_SMALL = 0
                ICON_BIG = 1

                hicon_big = ctypes.windll.user32.LoadImageW(
                    None, str(ico_path), IMAGE_ICON, 256, 256, LR_LOADFROMFILE
                ) or ctypes.windll.user32.LoadImageW(
                    None, str(ico_path), IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE
                )
                hicon_small = ctypes.windll.user32.LoadImageW(
                    None, str(ico_path), IMAGE_ICON, 32, 32, LR_LOADFROMFILE
                )
                if hicon_big:
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_big)
                if hicon_small:
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
            except Exception:
                pass

    def _load_brand_images(self) -> None:
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

        sep = ctk.CTkFrame(self.sidebar, height=1, fg_color="#23427D")
        sep.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="ew")

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

        info_card = ctk.CTkFrame(self.sidebar, fg_color="#163570", corner_radius=10)
        info_card.grid(row=7, column=0, padx=14, pady=8, sticky="ew")

        ctk.CTkLabel(
            info_card,
            text="🛡️ Fichier Excel Unique",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_CYAN,
        ).pack(anchor="w", padx=12, pady=(10, 2))

        ctk.CTkLabel(
            info_card,
            text="• Liens inscrits dans l'Excel initial\n• Préservation des données existantes\n• 'Dossier vide' colorié en rouge\n• Mise à jour différentielle B2",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color="#D0E1FD",
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 10))

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

        # 0. Barre supérieure
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
            text="Mise à jour directe des statuts et des liens Images/Vidéos dans votre fichier Excel initial.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=COLOR_TEXT_SECONDARY,
        )
        welcome_sub.grid(row=1, column=1, padx=4, pady=(0, 10), sticky="w")

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

        # 1. Rangée des 4 Cartes KPI
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
            self.kpi_row, 3, "🛡️", "#E0F7FA", COLOR_CYAN, "Fichiers vérifiés / envoyés", "0 / 0 fichiers"
        )

        # 2. Grille centrale : Carte Configuration B2 + Carte Sélection du Fichier Excel Unique
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

        # --- CARTE B : Sélection du Répertoire Parent, Fichier Excel Unique & Image Standard TAV ---
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
            text="📂  Répertoire Parent, Fichier Excel & Image Standard TAV",
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

        # Fichier Excel Unique (dossiers.xlsx)
        ctk.CTkLabel(
            self.card_files,
            text="Fichier Excel",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=2, column=0, padx=(16, 8), pady=5, sticky="w")
        self.source_excel_var = ctk.StringVar()
        self.source_excel_entry = ctk.CTkEntry(
            self.card_files,
            textvariable=self.source_excel_var,
            placeholder_text="Chemin du fichier Excel (dossiers + statuts + liens Images/Vidéos)...",
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

        # Image Standard TAV (fournie une fois pour toutes pour les dossiers vides commençant par TAV)
        ctk.CTkLabel(
            self.card_files,
            text="Image Standard TAV",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLOR_NAVY,
        ).grid(row=3, column=0, padx=(16, 8), pady=(5, 14), sticky="w")
        self.tav_image_var = ctk.StringVar()
        self.tav_image_entry = ctk.CTkEntry(
            self.card_files,
            textvariable=self.tav_image_var,
            placeholder_text="Image standard (fournie 1 fois pour toutes) pour les dossiers TAV vides...",
            fg_color=COLOR_SURFACE,
            border_color=COLOR_INPUT_BORDER,
            text_color=COLOR_NAVY,
            height=32,
        )
        self.tav_image_entry.grid(row=3, column=1, columnspan=2, padx=(0, 6), pady=(5, 14), sticky="ew")
        self.browse_tav_btn = ctk.CTkButton(
            self.card_files,
            text="Choisir image...",
            width=100,
            height=32,
            corner_radius=8,
            fg_color=COLOR_BACKGROUND,
            border_width=1,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            command=self._browse_tav_standard_image,
        )
        self.browse_tav_btn.grid(row=3, column=3, padx=(0, 16), pady=(5, 14), sticky="e")

        # 3. Carte de Contrôle, Barre de Progression & Actions
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
            command=lambda: self._start_backup(is_power_loss_resume=False),
        )
        self.start_btn.pack(side="left", padx=(0, 10))

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

        self.open_excel_btn = ctk.CTkButton(
            btn_bar,
            text="📊  Ouvrir le Fichier Excel",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            height=40,
            width=195,
            corner_radius=20,
            fg_color=COLOR_BACKGROUND,
            border_width=1.5,
            border_color=COLOR_PRIMARY,
            text_color=COLOR_PRIMARY,
            hover_color=COLOR_SURFACE,
            command=self._open_source_excel,
        )
        self.open_excel_btn.pack(side="right", padx=(6, 0))

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

        # 4. Carte Journal d'Exécution & Historique des Opérations
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
            self.tav_image_var,
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
        self.key_id_var.set(self.saved_config.get("b2_key_id", ""))
        self.app_key_var.set(self.saved_config.get("b2_app_key", ""))
        self.bucket_var.set(self.saved_config.get("bucket_name", ""))
        self.parent_dir_var.set(self.saved_config.get("parent_dir", ""))
        self.source_excel_var.set(self.saved_config.get("source_excel", ""))

        tav_img = self.saved_config.get("tav_standard_image_path", "")
        if not tav_img:
            default_tav = find_asset_path("image_standard_tav.png")
            if default_tav:
                tav_img = str(default_tav)
        self.tav_image_var.set(tav_img)

        self._append_log(
            f"Application v{self.current_version} prête. Les liens Images et Vidéos seront inscrits et mis à jour directement dans le fichier Excel initial.",
            "INFO",
        )

    def _save_current_config(self) -> None:
        data = {
            "b2_key_id": self.key_id_var.get().strip(),
            "b2_app_key": self.app_key_var.get().strip(),
            "bucket_name": self.bucket_var.get().strip(),
            "parent_dir": self.parent_dir_var.get().strip(),
            "source_excel": self.source_excel_var.get().strip(),
            "tav_standard_image_path": self.tav_image_var.get().strip(),
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

        if key_id and app_key and bucket and parent_dir and source_excel:
            last_folder = state.get("current_folder", "")
            self._append_log(
                f"⚡ Rallumage après coupure de courant détecté (dernier dossier actif : '{last_folder or 'Initialisation'}'). "
                "Lancement et reprise automatique là où l'application s'était arrêtée...",
                "WARN",
            )
            self._update_global_badge("Reprise auto", "warning")
            self.status_label.configure(text="⚡ Reprise automatique suite à une coupure de courant...")
            self.after(800, lambda: self._start_backup(is_power_loss_resume=True))

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
            exe_candidate = get_app_dir() / "dist" / "BackblazeB2Uploader" / "BackblazeB2Uploader.exe"
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable] + sys.argv[1:], cwd=str(get_app_dir()))
            elif exe_candidate.exists():
                subprocess.Popen([str(exe_candidate)], cwd=str(get_app_dir()))
            else:
                app_py = get_app_dir() / "app.py"
                pythonw = Path(sys.executable).with_name("pythonw.exe")
                runner = str(pythonw) if pythonw.exists() else sys.executable
                subprocess.Popen([runner, str(app_py)], cwd=str(get_app_dir()))
        except Exception:
            pass
        self.destroy()

    # =========================================================================
    # Sélecteurs de fichiers / répertoires / image standard TAV
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
            title="Sélectionner le fichier Excel (dossiers.xlsx)",
            initialdir=initial_dir,
            filetypes=[("Fichiers Excel", "*.xlsx")],
        )
        if chosen:
            self.source_excel_var.set(os.path.normpath(chosen))

    def _browse_tav_standard_image(self) -> None:
        initial_file = self.tav_image_var.get().strip()
        initial_dir = str(Path(initial_file).parent) if initial_file and os.path.exists(initial_file) else str(get_app_dir())
        chosen = filedialog.askopenfilename(
            title="Sélectionner l'Image Standard pour les dossiers TAV vides (fournie une fois pour toutes)",
            initialdir=initial_dir,
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.gif")],
        )
        if chosen:
            norm_chosen = os.path.normpath(chosen)
            self.tav_image_var.set(norm_chosen)
            cfg = self.config_manager.load()
            cfg["tav_standard_image_path"] = norm_chosen
            cfg["tav_standard_image_url"] = ""
            self.config_manager.save(cfg)
            self._append_log(f"Image standard TAV enregistrée une fois pour toutes : {norm_chosen}", "SUCCESS")

    def _generate_template_excel(self) -> None:
        parent_dir = self.parent_dir_var.get().strip()
        target_excel = self.source_excel_var.get().strip()
        if not target_excel:
            target_excel = str(get_app_dir() / "dossiers.xlsx")

        if os.path.exists(target_excel):
            resp = messagebox.askyesno(
                "Confirmation",
                f"Le fichier '{Path(target_excel).name}' existe déjà.\n\n"
                "Voulez-vous le remplacer par un nouveau modèle basé sur les sous-dossiers du Répertoire Parent ?",
            )
            if not resp:
                return

        try:
            count = ExcelManager.create_source_from_directory(target_excel, parent_dir if parent_dir else None)
            self.source_excel_var.set(os.path.normpath(target_excel))
            if count > 0:
                self._append_log(
                    f"Fichier Excel '{target_excel}' créé avec {count} sous-dossier(s) et les colonnes Statut / Images / Vidéos.",
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

    def _open_source_excel(self) -> None:
        excel_path = self.source_excel_var.get().strip()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showinfo(
                "Fichier non disponible",
                "Veuillez sélectionner ou créer un fichier Excel.",
            )
            return
        try:
            os.startfile(excel_path)
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
    def _start_backup(self, is_power_loss_resume: bool = False) -> None:
        if self.is_running:
            return

        key_id = self.key_id_var.get().strip()
        app_key = self.app_key_var.get().strip()
        bucket_name = self.bucket_var.get().strip()
        parent_dir = self.parent_dir_var.get().strip()
        source_excel = self.source_excel_var.get().strip()
        tav_std_img = self.tav_image_var.get().strip()

        if not key_id or not app_key or not bucket_name:
            messagebox.showwarning(
                "Identifiants Backblaze B2 requis",
                "Veuillez renseigner vos identifiants Backblaze B2 (Key ID, Application Key et Bucket Name).",
            )
            return

        if not parent_dir or (not is_power_loss_resume and not os.path.isdir(parent_dir)):
            messagebox.showwarning(
                "Répertoire Parent invalide",
                "Veuillez sélectionner un Répertoire Parent existant sur votre ordinateur.",
            )
            return

        if not source_excel or (not is_power_loss_resume and not os.path.isfile(source_excel)):
            messagebox.showwarning(
                "Fichier Excel introuvable",
                "Veuillez sélectionner un fichier Excel existant.",
            )
            return

        self._save_current_config()

        if not is_power_loss_resume:
            self.config_manager.mark_session_active(True, current_folder="", last_completed_row=0)
        else:
            self.config_manager.mark_session_active(True)

        self.is_running = True
        self.pause_event.set()
        self.stop_event.clear()

        self._set_controls_state(running=True, paused=False)
        self._update_global_badge("Sauvegarde en cours", "running")
        self.status_label.configure(text="Initialisation du transfert et vérification des dossiers...")

        self.worker_thread = threading.Thread(
            target=self._backup_worker,
            args=(
                key_id,
                app_key,
                bucket_name,
                parent_dir,
                source_excel,
                tav_std_img,
                is_power_loss_resume,
            ),
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
            "Tous les fichiers déjà envoyés sont enregistrés dans votre fichier Excel et la sauvegarde reprendra là où elle s'est arrêtée au prochain lancement.",
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
            self.browse_tav_btn.configure(state="normal")
            self.gen_source_btn.configure(state="normal")
        else:
            self.start_btn.configure(state="disabled")
            self.test_conn_btn.configure(state="disabled")
            self.browse_parent_btn.configure(state="disabled")
            self.browse_source_btn.configure(state="disabled")
            self.browse_tav_btn.configure(state="disabled")
            self.gen_source_btn.configure(state="disabled")
            self.stop_btn.configure(state="normal")
            if paused:
                self.pause_btn.configure(state="disabled")
                self.resume_btn.configure(state="normal")
            else:
                self.pause_btn.configure(state="normal")
                self.resume_btn.configure(state="disabled")

    # =========================================================================
    # Thread principal de traitement et téléversement (Fichier Excel Unique)
    # =========================================================================
    def _backup_worker(
        self,
        key_id: str,
        app_key: str,
        bucket_name: str,
        parent_dir: str,
        source_excel: str,
        tav_std_img: str,
        is_power_loss_resume: bool = False,
    ) -> None:
        try:
            self._queue_log("=" * 72, "INFO")
            self._queue_log("DÉMARRAGE DE LA SAUVEGARDE & MISE À JOUR DU FICHIER EXCEL", "INFO")
            self._queue_log(f"Répertoire Parent : {parent_dir}", "INFO")
            self._queue_log(f"Fichier Excel     : {source_excel}", "INFO")

            # Si on redémarre automatiquement après une coupure de courant et que le disque externe / dossier
            # met quelques secondes à se monter sous Windows, on attend qu'il soit prêt :
            if is_power_loss_resume and not os.path.isdir(parent_dir):
                self._queue_log(
                    f"⏳ En attente de l'initialisation du disque '{parent_dir}' après rallumage de la machine...",
                    "WARN",
                )
                self._queue_status("⏳ Attente de la disponibilité du disque après rallumage...")
                waited_disk = 0
                while not os.path.isdir(parent_dir) and waited_disk < 90:
                    if self.stop_event.is_set():
                        raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")
                    time.sleep(2.0)
                    waited_disk += 2

            folders_info = ExcelManager.read_source_folders(source_excel)
            total_listed = len(folders_info)

            if total_listed == 0:
                self.config_manager.mark_session_active(False)
                self._queue_log("Le fichier Excel ne contient aucun nom de dossier.", "WARN")
                self.ui_queue.put({"type": "finished", "cancelled": False, "error": None})
                return

            state_journal = self.config_manager.load_state_journal()
            last_row_done = int(state_journal.get("last_completed_row_in_run", 0)) if is_power_loss_resume else 0

            # 1. Résolution des dossiers (correspondance exacte ou par préfixe)
            # Les lignes dont le dossier n'est pas dans ce Répertoire Parent ne sont PAS modifiées
            # afin de préserver intégralement les informations existantes dans le fichier Excel.
            self._queue_status("Résolution des dossiers (préfixe inclus) et analyse du contenu...")
            parent_path = Path(parent_dir)
            total_files_all = 0
            total_bytes_all = 0
            resolved_items: List[Dict[str, Any]] = []

            for item in folders_info:
                if self.stop_event.is_set():
                    raise UploadCancelledError("Sauvegarde interrompue par l'utilisateur.")

                excel_name = item["dossier"]
                matched_folder = resolve_matching_folder(parent_path, excel_name)

                if matched_folder is None:
                    self._queue_log(
                        f"⏭ Dossier '{excel_name}' absent de '{parent_dir}' : ses informations existantes dans l'Excel sont conservées.",
                        "INFO",
                    )
                    continue

                if matched_folder.name != excel_name:
                    self._queue_log(
                        f"🔎 Dossier Excel '{excel_name}' associé par préfixe au dossier : '{matched_folder.name}'",
                        "INFO",
                    )

                files_in_folder = B2UploaderService.scan_folder_files(matched_folder)
                folder_bytes = sum(fp.stat().st_size for fp, _, _ in files_in_folder)

                item["folder_path"] = matched_folder
                item["files_count"] = len(files_in_folder)
                item["folder_bytes"] = folder_bytes
                total_files_all += len(files_in_folder)
                total_bytes_all += folder_bytes
                resolved_items.append(item)

            if not resolved_items:
                self.config_manager.mark_session_active(False)
                self._queue_log(
                    "Aucun dossier correspondant aux lignes du fichier Excel n'a été trouvé dans le Répertoire Parent.",
                    "WARN",
                )
                self.ui_queue.put(
                    {
                        "type": "finished",
                        "cancelled": False,
                        "error": "Aucun dossier correspondant trouvé dans le Répertoire Parent.",
                    }
                )
                return

            self._queue_log(
                f"Analyse terminée : {len(resolved_items)} dossier(s) résolu(s), "
                f"{total_files_all} fichier(s) éligible(s), volume total : {format_size(total_bytes_all)}.",
                "INFO",
            )

            # 2. Connexion à Backblaze B2
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

            cached_tav_url: str = self.config_manager.load().get("tav_standard_image_url", "").strip()

            folders_done_count = 0
            files_done_count = 0
            bytes_done_count = 0
            local_journal_files: Dict[str, Any] = state_journal.get("uploaded_files", {})

            self._queue_progress(
                folders_done=0,
                folders_total=len(resolved_items),
                files_done=0,
                files_total=total_files_all,
                bytes_done=0,
                bytes_total=total_bytes_all,
                status_text="Traitement et mise à jour du fichier Excel en cours...",
                fraction=0.0,
            )

            # 3. Traitement de chaque dossier résolu et mise à jour directe dans le fichier Excel initial
            resumed_logged = False
            for idx, item in enumerate(resolved_items, start=1):
                excel_dossier_name = item["dossier"]
                row_index = item["row_index"]
                folder_path: Path = item["folder_path"]
                files_count: int = item["files_count"]
                folder_bytes: int = item["folder_bytes"]

                if is_power_loss_resume and last_row_done > 0 and row_index <= last_row_done:
                    folders_done_count += 1
                    files_done_count += files_count
                    bytes_done_count += folder_bytes
                    self._queue_log(
                        f"⏭ [{idx}/{len(resolved_items)}] Dossier '{excel_dossier_name}' déjà finalisé avant la coupure de courant : ignoré.",
                        "INFO",
                    )
                    continue

                if is_power_loss_resume and not resumed_logged:
                    resumed_logged = True
                    self._queue_log(
                        f"⚡ Reprise exacte au point d'arrêt sur le dossier : '{excel_dossier_name}' (ligne Excel {row_index})",
                        "WARN",
                    )

                self.config_manager.mark_session_active(
                    True, current_folder=excel_dossier_name, last_completed_row=last_row_done
                )

                self._queue_log("-" * 64, "INFO")
                self._queue_log(
                    f"📁 [{idx}/{len(resolved_items)}] Dossier Excel : '{excel_dossier_name}' "
                    f"(Dossier disque : '{folder_path.name}' — {files_count} fichier(s), {format_size(folder_bytes)})",
                    "INFO",
                )
                self._queue_status(f"Traitement : {excel_dossier_name} ({idx}/{len(resolved_items)})")

                # =============================================================
                # CAS 1 : LE DOSSIER À SAUVEGARDER EST VIDE (0 fichier éligible)
                # =============================================================
                if files_count == 0:
                    if is_tav_folder(excel_dossier_name, folder_path):
                        # Dossier vide commençant par "TAV" :
                        # -> Ne pas téléverser le dossier sur B2
                        # -> Mettre le lien unique de l'image standard dans Image_1 (et vider Image_2..10, Video_1..5)
                        # -> Écrire "Dossier vide" colorié en rouge
                        if not cached_tav_url:
                            img_src = tav_std_img
                            if not img_src:
                                default_asset = find_asset_path("image_standard_tav.png")
                                if default_asset:
                                    img_src = str(default_asset)
                            cached_tav_url = uploader.ensure_tav_standard_image_url(img_src)
                            cfg = self.config_manager.load()
                            cfg["tav_standard_image_url"] = cached_tav_url
                            self.config_manager.save(cfg)

                        ExcelManager.update_row_links(
                            excel_path=source_excel,
                            row_index=row_index,
                            dossier_name=excel_dossier_name,
                            image_links=[cached_tav_url],
                            video_links=[],
                            status="Dossier vide",
                        )
                        self._queue_log(
                            f"  🖼 Dossier '{excel_dossier_name}' vide (préfixe 'TAV') : "
                            "lien de l'image standard inscrit et statut 'Dossier vide' colorié en rouge.",
                            "WARN",
                        )
                    else:
                        # Dossier vide ne commençant PAS par "TAV" :
                        # -> Ne pas téléverser le dossier sur B2
                        # -> Ne rien mettre dans les colonnes de liens (effacer d'éventuels anciens liens obsolètes)
                        # -> Écrire "Dossier vide" colorié en rouge
                        ExcelManager.update_row_links(
                            excel_path=source_excel,
                            row_index=row_index,
                            dossier_name=excel_dossier_name,
                            image_links=[],
                            video_links=[],
                            status="Dossier vide",
                        )
                        self._queue_log(
                            f"  ⚪ Dossier '{excel_dossier_name}' vide (hors 'TAV') : "
                            "aucun lien inscrit et statut 'Dossier vide' colorié en rouge.",
                            "WARN",
                        )

                    folders_done_count += 1
                    last_row_done = row_index
                    self.config_manager.mark_session_active(
                        True, current_folder=excel_dossier_name, last_completed_row=last_row_done
                    )

                    frac = (
                        min(bytes_done_count / total_bytes_all, 1.0)
                        if total_bytes_all > 0
                        else min(folders_done_count / max(len(resolved_items), 1), 1.0)
                    )
                    self._queue_progress(
                        folders_done=folders_done_count,
                        folders_total=len(resolved_items),
                        files_done=files_done_count,
                        files_total=total_files_all,
                        bytes_done=bytes_done_count,
                        bytes_total=total_bytes_all,
                        status_text=f"Dossier vide traité : {excel_dossier_name}",
                        fraction=frac,
                    )
                    continue

                # =============================================================
                # CAS 2 : DOSSIER NON VIDE (NOUVEAU OU DÉJÀ SAUVEGARDÉ SUR B2)
                # =============================================================
                ExcelManager.update_source_status(
                    source_path=source_excel,
                    row_index=row_index,
                    dossier_name=excel_dossier_name,
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
                        frac = min((folders_done_count + 1) / max(len(resolved_items), 1), 1.0)

                    self._queue_progress(
                        folders_done=folders_done_count,
                        folders_total=len(resolved_items),
                        files_done=files_done_count,
                        files_total=total_files_all,
                        bytes_done=bytes_done_count,
                        bytes_total=total_bytes_all,
                        status_text=f"Traitement : {excel_dossier_name} — {file_path.name}",
                        fraction=frac,
                    )

                def _on_root_links_updated(
                    partial_imgs: List[str],
                    partial_vids: List[str],
                    file_idx_in_folder: int,
                    total_in_folder: int,
                ) -> None:
                    # Enregistrement au fil de l'eau dans le fichier Excel initial
                    ExcelManager.update_row_links(
                        excel_path=source_excel,
                        row_index=row_index,
                        dossier_name=excel_dossier_name,
                        image_links=partial_imgs,
                        video_links=partial_vids,
                    )

                root_images, root_videos, uploaded_count, _, sync_state = uploader.upload_folder(
                    folder_path=folder_path,
                    local_journal_files=local_journal_files,
                    on_file_uploaded=_on_file_done,
                    on_root_links_updated=_on_root_links_updated,
                )

                # Mise à jour finale des liens (ajout, modification ou suppression selon l'état actuel du dossier source)
                # et du statut 'Terminé' dans le fichier Excel initial
                ExcelManager.update_row_links(
                    excel_path=source_excel,
                    row_index=row_index,
                    dossier_name=excel_dossier_name,
                    image_links=root_images,
                    video_links=root_videos,
                    status="Terminé",
                )
                self._queue_log(
                    f"  📊 Fichier Excel mis à jour pour '{excel_dossier_name}' "
                    f"({len(root_images)} image(s) racine, {len(root_videos)} vidéo(s) racine — état B2 : {sync_state}).",
                    "INFO",
                )

                folders_done_count += 1
                last_row_done = row_index
                self.config_manager.mark_session_active(
                    True, current_folder=excel_dossier_name, last_completed_row=last_row_done
                )

                frac = (
                    min(bytes_done_count / total_bytes_all, 1.0)
                    if total_bytes_all > 0
                    else min(folders_done_count / max(len(resolved_items), 1), 1.0)
                )
                self._queue_progress(
                    folders_done=folders_done_count,
                    folders_total=len(resolved_items),
                    files_done=files_done_count,
                    files_total=total_files_all,
                    bytes_done=bytes_done_count,
                    bytes_total=total_bytes_all,
                    status_text=f"✔ Dossier traité : {excel_dossier_name}",
                    fraction=frac,
                )
                self._queue_log(
                    f"✔ Dossier '{excel_dossier_name}' finalisé et marqué 'Terminé'.",
                    "SUCCESS",
                )

            self.config_manager.mark_session_active(False, current_folder="", last_completed_row=0)

            self._queue_progress(
                folders_done=folders_done_count,
                folders_total=len(resolved_items),
                files_done=files_done_count,
                files_total=total_files_all,
                bytes_done=bytes_done_count,
                bytes_total=total_bytes_all,
                status_text="✔ Sauvegarde terminée — Le fichier Excel a été mis à jour avec tous les liens.",
                fraction=1.0,
            )
            self._queue_log("=" * 72, "INFO")
            self._queue_log("✔ OPÉRATION TERMINÉE : TOUS LES DOSSIERS ET LIENS SONT À JOUR DANS LE FICHIER EXCEL.", "SUCCESS")
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
                            "Sauvegarde terminée !",
                            "Tous les dossiers ont été sauvegardés et les liens Images/Vidéos ont été mis à jour dans votre fichier Excel.",
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
    if getattr(sys, "frozen", False) and not os.environ.get("_B2_EXTERNAL_LOADED"):
        _ext_app = get_app_dir() / "app.py"
        if _ext_app.exists():
            os.environ["_B2_EXTERNAL_LOADED"] = "1"
            import runpy
            runpy.run_path(str(_ext_app), run_name="__main__")
            sys.exit(0)

    # Si l'application est lancée automatiquement par Windows au rallumage du PC (--auto-boot) :
    # - Si une sauvegarde était en cours lors d'une coupure de courant (active_session == True),
    #   l'application s'ouvre toute seule et reprend immédiatement là où elle s'était arrêtée.
    # - Si aucune sauvegarde n'était en cours (arrêt normal), elle ne dérange pas l'utilisateur.
    if "--auto-boot" in sys.argv:
        try:
            _boot_cm = ConfigManager()
            _boot_state = _boot_cm.load_state_journal()
            if not _boot_state.get("active_session", False):
                sys.exit(0)
        except Exception:
            pass

    # Protection instance unique sous Windows (évite qu'une double fenêtre ne s'ouvre au démarrage)
    if sys.platform == "win32":
        try:
            _mutex = ctypes.windll.kernel32.CreateMutexW(
                None, False, "Local\\BackblazeB2Uploader_SingleInstance_Mutex"
            )
            if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
                sys.exit(0)
            sys._b2_single_instance_mutex = _mutex  # type: ignore[attr-defined]
        except Exception:
            pass

    app = BackblazeUploaderApp()
    app.mainloop()
