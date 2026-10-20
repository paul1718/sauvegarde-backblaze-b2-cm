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


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class BackblazeUploaderApp(ctk.CTk):
    """Application Desktop de Sauvegarde Automatisée vers Backblaze B2."""

    def __init__(self) -> None:
        super().__init__()

        self.updater = AutoUpdater()
        self.current_version = self.updater.get_local_version()

        self.title(f"Sauvegarde Automatisée Backblaze B2 — v{self.current_version}")
        self.geometry("1000x800")
        self.minsize(900, 720)

        # Chargement de l'icône de l'application
        icon_path = get_app_dir() / "assets" / "app_icon.ico"
        if not icon_path.exists():
            icon_path = Path(__file__).resolve().parent / "assets" / "app_icon.ico"
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass

        self.config_manager = ConfigManager()
        self.saved_config = self.config_manager.load()

        # File d'attente thread-safe pour les mises à jour de l'interface
        self.ui_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()

        # Événements de contrôle du thread de téléversement
        self.pause_event = threading.Event()
        self.pause_event.set()  # set = actif, clear = en pause
        self.stop_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self.is_running = False

        # Fichier log disque
        self.log_file_path = get_app_dir() / "sauvegarde_b2.log"

        # Construction de l'interface graphique
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

    def _build_ui(self) -> None:
        """Construit les 3 sections de l'interface utilisateur."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # En-tête de l'application + bouton de mise à jour à distance
        header_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=("gray88", "gray17"))
        header_frame.grid(row=0, column=0, padx=16, pady=(14, 6), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)

        title_lbl = ctk.CTkLabel(
            header_frame,
            text=f"☁️  Sauvegarde Automatisée Backblaze B2  (v{self.current_version})",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        title_lbl.grid(row=0, column=0, padx=16, pady=(10, 2), sticky="w")

        subtitle_lbl = ctk.CTkLabel(
            header_frame,
            text="Protection anti-coupure de courant • Pause/Reprise auto sur coupure Internet • Sauvegarde Excel en temps réel",
            font=ctk.CTkFont(size=12),
            text_color=("gray35", "gray70"),
        )
        subtitle_lbl.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        self.update_btn = ctk.CTkButton(
            header_frame,
            text="🔄 Vérifier mise à jour",
            width=165,
            height=32,
            fg_color=("gray75", "gray28"),
            text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray38"),
            command=lambda: self._check_remote_updates(silent=False),
        )
        self.update_btn.grid(row=0, column=1, rowspan=2, padx=16, pady=10, sticky="e")

        # =========================================================================
        # SECTION 1 : Configuration / Identifiants Backblaze B2
        # =========================================================================
        sec1 = ctk.CTkFrame(self, corner_radius=10)
        sec1.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
        sec1.grid_columnconfigure((1, 3), weight=1)

        sec1_title = ctk.CTkLabel(
            sec1,
            text="1. Configuration & Identifiants Backblaze B2",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        sec1_title.grid(row=0, column=0, columnspan=5, padx=14, pady=(10, 6), sticky="w")

        # B2 Key ID
        ctk.CTkLabel(sec1, text="B2 Key ID :").grid(row=1, column=0, padx=(14, 6), pady=6, sticky="w")
        self.key_id_var = ctk.StringVar()
        self.key_id_entry = ctk.CTkEntry(
            sec1, textvariable=self.key_id_var, placeholder_text="Ex: 005a1b2c3d4e5f0000000001"
        )
        self.key_id_entry.grid(row=1, column=1, padx=6, pady=6, sticky="ew")

        # Bucket Name
        ctk.CTkLabel(sec1, text="Bucket Name :").grid(row=1, column=2, padx=(14, 6), pady=6, sticky="w")
        self.bucket_var = ctk.StringVar()
        self.bucket_entry = ctk.CTkEntry(
            sec1, textvariable=self.bucket_var, placeholder_text="Nom exact du Bucket B2"
        )
        self.bucket_entry.grid(row=1, column=3, columnspan=2, padx=(6, 14), pady=6, sticky="ew")

        # B2 Application Key
        ctk.CTkLabel(sec1, text="B2 Application Key :").grid(row=2, column=0, padx=(14, 6), pady=(6, 12), sticky="w")
        self.app_key_var = ctk.StringVar()
        self.app_key_entry = ctk.CTkEntry(
            sec1, textvariable=self.app_key_var, show="•", placeholder_text="Clé secrète d'application B2"
        )
        self.app_key_entry.grid(row=2, column=1, columnspan=2, padx=6, pady=(6, 12), sticky="ew")

        self.show_key_btn = ctk.CTkButton(
            sec1,
            text="Afficher",
            width=85,
            fg_color=("gray75", "gray30"),
            text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray40"),
            command=self._toggle_app_key_visibility,
        )
        self.show_key_btn.grid(row=2, column=3, padx=6, pady=(6, 12), sticky="w")

        self.test_conn_btn = ctk.CTkButton(
            sec1,
            text="Tester la connexion",
            width=150,
            command=self._test_b2_connection,
        )
        self.test_conn_btn.grid(row=2, column=4, padx=(6, 14), pady=(6, 12), sticky="e")

        # =========================================================================
        # SECTION 2 : Sélection des Fichiers et Répertoires
        # =========================================================================
        sec2 = ctk.CTkFrame(self, corner_radius=10)
        sec2.grid(row=2, column=0, padx=16, pady=6, sticky="ew")
        sec2.grid_columnconfigure(1, weight=1)

        sec2_title = ctk.CTkLabel(
            sec2,
            text="2. Sélection des Fichiers et Répertoires",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        sec2_title.grid(row=0, column=0, columnspan=4, padx=14, pady=(10, 6), sticky="w")

        # Répertoire Parent
        ctk.CTkLabel(sec2, text="Répertoire Parent :").grid(row=1, column=0, padx=(14, 6), pady=5, sticky="w")
        self.parent_dir_var = ctk.StringVar()
        self.parent_dir_entry = ctk.CTkEntry(
            sec2,
            textvariable=self.parent_dir_var,
            placeholder_text="Chemin du dossier contenant les sous-dossiers à sauvegarder...",
        )
        self.parent_dir_entry.grid(row=1, column=1, columnspan=2, padx=6, pady=5, sticky="ew")
        self.browse_parent_btn = ctk.CTkButton(
            sec2, text="Parcourir...", width=120, command=self._browse_parent_dir
        )
        self.browse_parent_btn.grid(row=1, column=3, padx=(6, 14), pady=5, sticky="e")

        # Fichier Excel Source (dossiers.xlsx)
        ctk.CTkLabel(sec2, text="Excel Source (dossiers.xlsx) :").grid(
            row=2, column=0, padx=(14, 6), pady=5, sticky="w"
        )
        self.source_excel_var = ctk.StringVar()
        self.source_excel_entry = ctk.CTkEntry(
            sec2,
            textvariable=self.source_excel_var,
            placeholder_text="Chemin du fichier Excel contenant la colonne 'Dossier' et 'Statut'...",
        )
        self.source_excel_entry.grid(row=2, column=1, padx=6, pady=5, sticky="ew")

        self.gen_source_btn = ctk.CTkButton(
            sec2,
            text="Créer modèle",
            width=110,
            fg_color=("gray75", "gray30"),
            text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray40"),
            command=self._generate_template_excel,
        )
        self.gen_source_btn.grid(row=2, column=2, padx=4, pady=5, sticky="e")

        self.browse_source_btn = ctk.CTkButton(
            sec2, text="Parcourir...", width=120, command=self._browse_source_excel
        )
        self.browse_source_btn.grid(row=2, column=3, padx=(6, 14), pady=5, sticky="e")

        # Fichier Excel de Sortie (resultats_backblaze.xlsx)
        ctk.CTkLabel(sec2, text="Excel Sortie (resultats_backblaze.xlsx) :").grid(
            row=3, column=0, padx=(14, 6), pady=(5, 12), sticky="w"
        )
        self.output_excel_var = ctk.StringVar()
        self.output_excel_entry = ctk.CTkEntry(
            sec2,
            textvariable=self.output_excel_var,
            placeholder_text="Emplacement du fichier Excel de résultats...",
        )
        self.output_excel_entry.grid(row=3, column=1, columnspan=2, padx=6, pady=(5, 12), sticky="ew")
        self.browse_output_btn = ctk.CTkButton(
            sec2, text="Parcourir...", width=120, command=self._browse_output_excel
        )
        self.browse_output_btn.grid(row=3, column=3, padx=(6, 14), pady=(5, 12), sticky="e")

        # =========================================================================
        # SECTION 3 : Contrôle, Progression & Journal d'Exécution
        # =========================================================================
        sec3 = ctk.CTkFrame(self, corner_radius=10)
        sec3.grid(row=3, column=0, padx=16, pady=(6, 14), sticky="nsew")
        sec3.grid_columnconfigure(0, weight=1)
        sec3.grid_rowconfigure(4, weight=1)

        sec3_title = ctk.CTkLabel(
            sec3,
            text="3. Contrôle & Progression de la Sauvegarde",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        sec3_title.grid(row=0, column=0, padx=14, pady=(10, 6), sticky="w")

        # Barre de boutons de contrôle
        btn_frame = ctk.CTkFrame(sec3, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=14, pady=4, sticky="ew")

        self.start_btn = ctk.CTkButton(
            btn_frame,
            text="▶  Lancer la Sauvegarde",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=38,
            width=210,
            fg_color="#16A34A",
            hover_color="#15803D",
            command=self._start_backup,
        )
        self.start_btn.pack(side="left", padx=(0, 10))

        self.pause_btn = ctk.CTkButton(
            btn_frame,
            text="⏸  Pause",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=38,
            width=130,
            fg_color="#D97706",
            hover_color="#B45309",
            state="disabled",
            command=self._pause_backup,
        )
        self.pause_btn.pack(side="left", padx=6)

        self.resume_btn = ctk.CTkButton(
            btn_frame,
            text="⏯  Reprendre",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=38,
            width=140,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            state="disabled",
            command=self._resume_backup,
        )
        self.resume_btn.pack(side="left", padx=6)

        self.stop_btn = ctk.CTkButton(
            btn_frame,
            text="⏹  Interrompre",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=38,
            width=140,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            state="disabled",
            command=self._stop_backup,
        )
        self.stop_btn.pack(side="left", padx=6)

        self.open_output_btn = ctk.CTkButton(
            btn_frame,
            text="📊 Ouvrir Résultats Excel",
            height=38,
            width=175,
            fg_color=("gray75", "gray30"),
            text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray40"),
            command=self._open_output_excel,
        )
        self.open_output_btn.pack(side="right", padx=(6, 0))

        # Statistiques & Barre de progression globale
        prog_frame = ctk.CTkFrame(sec3, fg_color=("gray90", "gray20"), corner_radius=8)
        prog_frame.grid(row=2, column=0, padx=14, pady=8, sticky="ew")
        prog_frame.grid_columnconfigure(0, weight=1)

        stats_top = ctk.CTkFrame(prog_frame, fg_color="transparent")
        stats_top.grid(row=0, column=0, padx=12, pady=(8, 2), sticky="ew")
        stats_top.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            stats_top,
            text="Prêt à démarrer.",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.percent_label = ctk.CTkLabel(
            stats_top,
            text="0.0 %",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.percent_label.grid(row=0, column=1, sticky="e")

        self.progress_bar = ctk.CTkProgressBar(prog_frame, height=16)
        self.progress_bar.grid(row=1, column=0, padx=12, pady=6, sticky="ew")
        self.progress_bar.set(0.0)

        self.Counters_label = ctk.CTkLabel(
            prog_frame,
            text="Dossiers : 0 / 0   |   Fichiers : 0 / 0   |   Volume transféré : 0 o / 0 o",
            font=ctk.CTkFont(size=12),
            text_color=("gray30", "gray75"),
        )
        self.Counters_label.grid(row=2, column=0, padx=12, pady=(0, 8), sticky="w")

        # Journal d'exécution (Logs)
        log_header_frame = ctk.CTkFrame(sec3, fg_color="transparent")
        log_header_frame.grid(row=3, column=0, padx=14, pady=(2, 2), sticky="ew")
        ctk.CTkLabel(
            log_header_frame,
            text="Journal d'exécution en temps réel :",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left")

        ctk.CTkButton(
            log_header_frame,
            text="Effacer le journal",
            width=120,
            height=24,
            font=ctk.CTkFont(size=11),
            fg_color=("gray75", "gray30"),
            text_color=("gray10", "gray90"),
            hover_color=("gray65", "gray40"),
            command=self._clear_logs,
        ).pack(side="right")

        self.log_textbox = ctk.CTkTextbox(
            sec3,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
            state="disabled",
        )
        self.log_textbox.grid(row=4, column=0, padx=14, pady=(2, 14), sticky="nsew")

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

    def _load_config_into_ui(self) -> None:
        """Pré-remplit les champs avec les dernières valeurs mémorisées."""
        self.key_id_var.set(self.saved_config.get("b2_key_id", ""))
        self.app_key_var.set(self.saved_config.get("b2_app_key", ""))
        self.bucket_var.set(self.saved_config.get("bucket_name", ""))
        self.parent_dir_var.set(self.saved_config.get("parent_dir", ""))
        self.source_excel_var.set(self.saved_config.get("source_excel", ""))
        self.output_excel_var.set(self.saved_config.get("output_excel", ""))
        self._append_log(
            f"Application v{self.current_version} initialisée. Configuration et journal anti-coupure chargés.",
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
        """
        Vérifie si une session était active lors d'une coupure brusque de courant.
        Si oui, reprend automatiquement la sauvegarde là où elle s'était arrêtée.
        """
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
                f"⚡ Détection d'une interruption brusque (coupure de courant) sur la session précédente "
                f"(dernier dossier en cours : '{last_folder}'). Reprise automatique immédiate !",
                "WARN",
            )
            self.status_label.configure(
                text="⚡ Reprise automatique suite à une coupure de courant..."
            )
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
        """Relance proprement l'application après une mise à jour à distance."""
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
        """Génère un fichier Excel source `dossiers.xlsx` à partir des sous-dossiers du Répertoire Parent."""
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
                    f"Fichier source '{target_excel}' créé avec {count} sous-dossier(s) détecté(s) dans le Répertoire Parent.",
                    "SUCCESS",
                )
                messagebox.showinfo(
                    "Modèle Excel créé",
                    f"Le fichier '{Path(target_excel).name}' a été créé avec succès ({count} dossiers ajoutés).",
                )
            else:
                self._append_log(
                    f"Modèle Excel vide créé : '{target_excel}'. Vous pouvez y saisir les noms de dossiers.",
                    "INFO",
                )
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
                "Champs manquants",
                "Veuillez renseigner le 'B2 Key ID', la 'B2 Application Key' et le 'Bucket Name'.",
            )
            return

        self._save_current_config()
        self.test_conn_btn.configure(state="disabled", text="Test en cours...")
        self._append_log(f"Test de connexion au bucket Backblaze B2 '{bucket}'...", "INFO")

        def _worker() -> None:
            try:
                service = B2UploaderService(key_id, app_key, bucket)
                service.connect(max_auth_retries=1)
                self.ui_queue.put(
                    {
                        "type": "conn_test_result",
                        "success": True,
                        "msg": f"Connexion réussie au bucket '{bucket}' !",
                    }
                )
            except Exception as e:
                self.ui_queue.put(
                    {
                        "type": "conn_test_result",
                        "success": False,
                        "msg": f"Échec de connexion à Backblaze B2 : {e}",
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

        # Marquage de session active pour la reprise automatique en cas de coupure de courant
        self.config_manager.mark_session_active(True)

        self.is_running = True
        self.pause_event.set()
        self.stop_event.clear()

        self._set_controls_state(running=True, paused=False)
        self.status_label.configure(text="Initialisation de la sauvegarde...")

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
        self.status_label.configure(text="⏸ Sauvegarde en pause (cliquez sur 'Reprendre' pour continuer)")
        self._append_log("Pause demandée par l'utilisateur. Le transfert se met en pause...", "WARN")

    def _resume_backup(self) -> None:
        if not self.is_running or self.pause_event.is_set():
            return
        self.pause_event.set()
        self._set_controls_state(running=True, paused=False)
        self.status_label.configure(text="▶ Sauvegarde en cours...")
        self._append_log("Reprise de la sauvegarde.", "INFO")

    def _stop_backup(self) -> None:
        if not self.is_running:
            return
        resp = messagebox.askyesno(
            "Interrompre la sauvegarde",
            "Voulez-vous interrompre la sauvegarde en cours ?\n\n"
            "Tous les fichiers déjà envoyés sont enregistrés et la sauvegarde reprendra exactement là où elle s'est arrêtée au prochain lancement.",
        )
        if not resp:
            return
        self.config_manager.mark_session_active(False)
        self.stop_event.set()
        self.pause_event.set()
        self.stop_btn.configure(state="disabled")
        self.status_label.configure(text="Arrêt en cours...")
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
            self._queue_log("DÉMARRAGE DE LA SESSION DE SAUVEGARDE BACKBLAZE B2", "INFO")
            self._queue_log(f"Répertoire Parent : {parent_dir}", "INFO")
            self._queue_log(f"Fichier Source    : {source_excel}", "INFO")
            self._queue_log(f"Fichier Sortie    : {output_excel}", "INFO")

            # 1. Lecture sécurisée du fichier Excel source (avec auto-réparation .bak si nécessaire)
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
                    status_text="Tous les dossiers sont déjà marqués 'Terminé' !",
                    fraction=1.0,
                )
                self._queue_log("Aucun dossier restant à traiter. Sauvegarde terminée !", "SUCCESS")
                self.ui_queue.put({"type": "finished", "cancelled": False, "error": None})
                return

            # 2. Pré-calcul du nombre total de fichiers et du volume à transférer
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
                        f"⚠ Attention : Le dossier '{item['dossier']}' est introuvable dans '{parent_dir}'. Il sera ignoré.",
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

            # 3. Connexion à Backblaze B2 (avec attente automatique si Internet est coupé)
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

            # 4. Boucle sur chaque dossier éligible
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
                status_text="Téléversement en cours...",
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
                self._queue_status(f"Téléversement du dossier : {dossier_name} ({idx}/{len(valid_pending)})")

                # Marquer "En cours" dans l'Excel source de manière atomique
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

                    # Enregistrement immédiat dans le journal anti-coupure de courant
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
                        status_text=f"En cours : {dossier_name} — {file_path.name}",
                        fraction=frac,
                    )

                def _on_root_links_updated(
                    partial_imgs: List[str],
                    partial_vids: List[str],
                    file_idx_in_folder: int,
                    total_in_folder: int,
                ) -> None:
                    # Enregistrement en temps réel (au fil de l'eau) du fichier Excel final de manière atomique
                    ExcelManager.upsert_output_result(
                        output_path=output_excel,
                        dossier_name=dossier_name,
                        image_links=partial_imgs,
                        video_links=partial_vids,
                    )

                # Téléversement complet du dossier et de ses sous-dossiers
                root_images, root_videos, uploaded_count, _ = uploader.upload_folder(
                    folder_path=folder_path,
                    local_journal_files=local_journal_files,
                    on_file_uploaded=_on_file_done,
                    on_root_links_updated=_on_root_links_updated,
                )

                # Consolidation finale de la ligne dans l'Excel de sortie
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

                # Mise à jour immédiate et atomique du statut 'Terminé' dans l'Excel source
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
                    status_text=f"Dossier terminé : {dossier_name}",
                    fraction=frac,
                )
                self._queue_log(
                    f"✅ Dossier '{dossier_name}' intégralement téléversé et marqué 'Terminé' dans l'Excel source.",
                    "SUCCESS",
                )

            # Fin normale de tous les dossiers : désactivation du drapeau de coupure de courant
            self.config_manager.mark_session_active(False, current_folder="")

            self._queue_progress(
                folders_done=folders_done_count,
                folders_total=total_listed,
                files_done=files_done_count,
                files_total=total_files_all,
                bytes_done=bytes_done_count,
                bytes_total=total_bytes_all,
                status_text="🎉 Sauvegarde intégralement terminée avec succès !",
                fraction=1.0,
            )
            self._queue_log("=" * 72, "INFO")
            self._queue_log("🎉 TOUS LES DOSSIERS ONT ÉTÉ SAUVEGARDÉS ET CONSIGNÉS AVEC SUCCÈS !", "SUCCESS")
            self.ui_queue.put({"type": "finished", "cancelled": False, "error": None})

        except UploadCancelledError as e:
            self._queue_log(f"⏹ {e}", "WARN")
            self._queue_status("⏹ Sauvegarde interrompue. Elle reprendra automatiquement au prochain lancement.")
            self.ui_queue.put({"type": "finished", "cancelled": True, "error": None})

        except ExcelPermissionError as e:
            self._queue_log(f"❌ Erreur d'accès Excel : {e}", "ERROR")
            self._queue_status("❌ Erreur : Fichier Excel ouvert dans une autre application.")
            self.ui_queue.put({"type": "finished", "cancelled": False, "error": str(e)})

        except Exception as e:
            self._queue_log(f"❌ Erreur critique : {e}", "ERROR")
            self._queue_status(f"❌ Erreur : {e}")
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
        """Traite les messages envoyés par le thread de sauvegarde sans bloquer l'interface."""
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                msg_type = msg.get("type")

                if msg_type == "log":
                    self._append_log(msg["message"], msg.get("level", "INFO"))

                elif msg_type == "status":
                    self.status_label.configure(text=msg["text"])

                elif msg_type == "progress":
                    frac = max(0.0, min(1.0, float(msg["fraction"])))
                    self.progress_bar.set(frac)
                    self.percent_label.configure(text=f"{frac * 100:.1f} %")
                    if self.pause_event.is_set():
                        self.status_label.configure(text=msg["status_text"])
                    self.Counters_label.configure(
                        text=(
                            f"Dossiers : {msg['folders_done']} / {msg['folders_total']}   |   "
                            f"Fichiers : {msg['files_done']} / {msg['files_total']}   |   "
                            f"Volume transféré : {format_size(msg['bytes_done'])} / {format_size(msg['bytes_total'])}"
                        )
                    )

                elif msg_type == "conn_test_result":
                    self.test_conn_btn.configure(state="normal", text="Tester la connexion")
                    if msg["success"]:
                        self._append_log(f"✅ {msg['msg']}", "SUCCESS")
                        messagebox.showinfo("Connexion Backblaze B2", msg["msg"])
                    else:
                        self._append_log(f"❌ {msg['msg']}", "ERROR")
                        messagebox.showerror("Erreur de connexion Backblaze B2", msg["msg"])

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
                                fg_color="#2563EB",
                                text_color="white",
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
                            f"✅ Mise à jour v{rv} installée avec succès ! Redémarrage de l'application...",
                            "SUCCESS",
                        )
                        self.after(800, self._restart_application)
                    else:
                        self._append_log(f"❌ {msg['message']}", "ERROR")
                        messagebox.showerror("Erreur de mise à jour", msg["message"])

                elif msg_type == "finished":
                    self.is_running = False
                    self._set_controls_state(running=False, paused=False)
                    error_msg = msg.get("error")
                    cancelled = msg.get("cancelled", False)
                    if error_msg:
                        messagebox.showerror("Erreur pendant la sauvegarde", error_msg)
                    elif not cancelled:
                        messagebox.showinfo(
                            "Sauvegarde terminée",
                            "La sauvegarde vers Backblaze B2 et la mise à jour des fichiers Excel sont terminées avec succès !",
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
