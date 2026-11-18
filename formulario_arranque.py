"""Form-based tool for calculating motor start-up parameters."""

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import math
from decimal import Decimal, ROUND_HALF_UP
from typing import Any
# Optional dependencies: pandas for Excel lookup, fpdf for PDF export, Pillow for images
try:
    import pandas as pd
except Exception:
    pd = None
    print("[PANDAS] pandas not installed. MLFB lookup will be disabled.")

try:
    from fpdf import FPDF
except Exception:
    FPDF = None
    print("[FPDF] fpdf not installed. PDF export will be disabled.")

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None
    print("[PIL] Pillow not installed. Header images will be skipped.")
import os
import sys
import subprocess
from tkinter import scrolledtext

from style_utils import ABB_COLORS, aplicar_colorimetria
import logging

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
debug_enabled = True

# ---------------------------------------------------------------------------
# Engineering thresholds and UI state palettes
# Values below are business defaults used in thermal/start-up evaluations.
# ---------------------------------------------------------------------------
TEMP_MAX_FRIO = 120.0
TEMP_MAX_CALIENTE = 200.0
TEMP_INICIAL_CALIENTE = 120.0
TX_CALIENTE = 130.0

ESTADO_UMBRAL = 0.9
COLOR_ESTADO = {
    "OK": "#2e7d32",
    "Cerca": "#d4a300",
    "Excede": "#c00000",
}
BANDA_OK = "#eaf5ee"
BANDA_ALERTA = "#fff3cd"
BANDA_PELIGRO = "#fdecea"

PDF_COLOR_ESTADO = {
    "OK": (46, 125, 50),
    "Cerca": (212, 163, 0),
    "Excede": (192, 0, 0),
}
PDF_BANDA_OK = (234, 245, 238)
PDF_BANDA_ALERTA = (255, 243, 205)
PDF_BANDA_PELIGRO = (253, 236, 234)

UI_THEME = {
    "bg": "#eef2f6",
    "panel": "#ffffff",
    "panel_alt": "#f8fafc",
    "text": "#1f2937",
    "muted": "#6b7280",
    "border": "#d7dde5",
    "accent": "#d81e05",
    "accent_hover": "#b71805",
    "ok_soft": "#eaf5ee",
    "warn_soft": "#fff4d6",
    "bad_soft": "#fdecea",
    "info_soft": "#ecf5ff",
}

UI_FONTS = {
    "hero": ("Segoe UI Semibold", 18),
    "subtitle": ("Segoe UI", 10),
    "title": ("Segoe UI Semibold", 12),
    "body": ("Segoe UI", 10),
    "button": ("Segoe UI Semibold", 10),
}


# ---------------------------------------------------------------------------
# UI theme setup
# Configure a consistent ttk theme so all widgets share visual semantics.
# ---------------------------------------------------------------------------
def configurar_estilo_moderno(widget: tk.Misc) -> None:
    """Aplica un tema visual moderno sin alterar comportamiento."""
    style = ttk.Style(widget)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    def _cfg(style_name: str, **kwargs) -> None:
        try:
            style.configure(style_name, **kwargs)
        except tk.TclError:
            pass

    def _map(style_name: str, **kwargs) -> None:
        try:
            style.map(style_name, **kwargs)
        except tk.TclError:
            pass

    _cfg(
        ".",
        background=UI_THEME["bg"],
        foreground=UI_THEME["text"],
        font=UI_FONTS["body"],
    )
    _cfg(
        "TLabel",
        background=UI_THEME["bg"],
        foreground=UI_THEME["text"],
        font=UI_FONTS["body"],
    )
    _cfg(
        "Hero.TFrame",
        background=UI_THEME["panel"],
        bordercolor=UI_THEME["border"],
        borderwidth=1,
        relief="solid",
    )
    _cfg(
        "HeroTitle.TLabel",
        background=UI_THEME["panel"],
        foreground="#111827",
        font=UI_FONTS["hero"],
    )
    _cfg(
        "HeroSubtitle.TLabel",
        background=UI_THEME["panel"],
        foreground=UI_THEME["muted"],
        font=UI_FONTS["subtitle"],
    )
    _cfg(
        "TLabelframe",
        background=UI_THEME["panel"],
        bordercolor=UI_THEME["border"],
        borderwidth=1,
        relief="solid",
    )
    _cfg(
        "TLabelframe.Label",
        background=UI_THEME["panel"],
        foreground=UI_THEME["text"],
        font=UI_FONTS["title"],
    )
    _cfg(
        "Card.TLabelframe",
        background=UI_THEME["panel_alt"],
        bordercolor=UI_THEME["border"],
        borderwidth=1,
        relief="solid",
    )
    _cfg(
        "Card.TLabelframe.Label",
        background=UI_THEME["panel_alt"],
        foreground=UI_THEME["text"],
        font=UI_FONTS["title"],
    )
    _cfg(
        "Status.TLabel",
        background=UI_THEME["panel_alt"],
        foreground=UI_THEME["text"],
        font=("Segoe UI Semibold", 11),
        padding=(12, 10),
    )
    _cfg(
        "StatusOk.TLabel",
        background=UI_THEME["ok_soft"],
        foreground="#1f6f3f",
        font=("Segoe UI Semibold", 11),
        padding=(12, 10),
    )
    _cfg(
        "StatusWarn.TLabel",
        background=UI_THEME["warn_soft"],
        foreground="#8a5a00",
        font=("Segoe UI Semibold", 11),
        padding=(12, 10),
    )
    _cfg(
        "StatusBad.TLabel",
        background=UI_THEME["bad_soft"],
        foreground="#9a1a1a",
        font=("Segoe UI Semibold", 11),
        padding=(12, 10),
    )
    _cfg(
        "StatusInfo.TLabel",
        background=UI_THEME["info_soft"],
        foreground="#0b5cad",
        font=("Segoe UI Semibold", 11),
        padding=(12, 10),
    )
    _cfg(
        "FieldLabel.TLabel",
        background=UI_THEME["panel_alt"],
        foreground="#111827",
        font=("Segoe UI Semibold", 10),
    )
    _cfg(
        "Muted.TLabel",
        background=UI_THEME["panel_alt"],
        foreground=UI_THEME["muted"],
        font=("Segoe UI", 9),
    )
    _cfg(
        "TButton",
        background="#e5e7eb",
        foreground=UI_THEME["text"],
        font=UI_FONTS["button"],
        borderwidth=0,
        padding=(10, 7),
    )
    _map(
        "TButton",
        background=[("active", "#d1d5db"), ("pressed", "#c7ccd3")],
    )
    _cfg(
        "Accent.TButton",
        background=UI_THEME["accent"],
        foreground="#ffffff",
        font=UI_FONTS["button"],
        borderwidth=0,
        padding=(12, 8),
    )
    _map(
        "Accent.TButton",
        background=[("active", UI_THEME["accent_hover"]), ("pressed", "#9d1404")],
        foreground=[("disabled", "#f3f4f6")],
    )
    _cfg(
        "TEntry",
        fieldbackground="#ffffff",
        foreground=UI_THEME["text"],
        bordercolor=UI_THEME["border"],
        lightcolor=UI_THEME["border"],
        darkcolor=UI_THEME["border"],
        insertcolor=UI_THEME["text"],
        padding=4,
    )
    _cfg(
        "ReadonlyDark.TEntry",
        fieldbackground="#f3f4f6",
        foreground="#111827",
        bordercolor=UI_THEME["border"],
        lightcolor=UI_THEME["border"],
        darkcolor=UI_THEME["border"],
        padding=4,
    )
    _cfg(
        "TCombobox",
        fieldbackground="#ffffff",
        foreground=UI_THEME["text"],
        bordercolor=UI_THEME["border"],
        lightcolor=UI_THEME["border"],
        darkcolor=UI_THEME["border"],
        padding=3,
    )
    _cfg(
        "TCheckbutton",
        background=UI_THEME["bg"],
        foreground=UI_THEME["text"],
        font=UI_FONTS["body"],
    )
    _cfg(
        "Modern.Treeview",
        background="#ffffff",
        foreground=UI_THEME["text"],
        fieldbackground="#ffffff",
        bordercolor=UI_THEME["border"],
        rowheight=24,
    )
    _cfg(
        "Modern.Treeview.Heading",
        background="#eef2f7",
        foreground="#111827",
        font=UI_FONTS["button"],
        relief="flat",
    )
    _map("Modern.Treeview.Heading", background=[("active", "#e5eaf1")])
    _cfg("TPanedwindow", background=UI_THEME["bg"])
    _cfg("Sash", background=UI_THEME["border"], sashthickness=8)


def obtener_estado_temperatura(max_temp: float, max_val: float) -> str:
    if max_temp <= 0:
        return "OK"
    ratio = max_val / max_temp
    if ratio >= 1.0:
        return "Excede"
    if ratio >= ESTADO_UMBRAL:
        return "Cerca"
    return "OK"


class ToolTip:
    """Simple tooltip para widgets de Tkinter."""

    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tipwindow: tk.Toplevel | None = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None) -> None:
        if self.tipwindow or not self.text:
            return
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        try:
            tw.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        label = tk.Label(
            tw,
            text=f"Ayuda:\n{self.text}",
            justify="left",
            background="#111827",
            foreground="#f9fafb",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9, "normal"),
            padx=8,
            pady=5,
            wraplength=360,
        )
        label.pack(ipadx=1)
        tw.update_idletasks()

        pointer_x = self.widget.winfo_pointerx()
        pointer_y = self.widget.winfo_pointery()
        tip_w = tw.winfo_reqwidth()
        tip_h = tw.winfo_reqheight()
        vroot_x = self.widget.winfo_vrootx()
        vroot_y = self.widget.winfo_vrooty()
        vroot_w = self.widget.winfo_vrootwidth()
        vroot_h = self.widget.winfo_vrootheight()
        right_limit = vroot_x + vroot_w
        bottom_limit = vroot_y + vroot_h
        margin = 8
        offset = 12

        # Anclar al cursor. Si no entra a la derecha, mover al lado izquierdo del cursor.
        x = pointer_x + offset
        y = pointer_y + offset
        if x + tip_w > right_limit - margin:
            x = pointer_x - tip_w - offset
        if x < vroot_x + margin:
            x = vroot_x + margin

        # Mantener visible en vertical dentro del monitor virtual activo.
        if y + tip_h > bottom_limit - margin:
            y = pointer_y - tip_h - offset
        if y < vroot_y + margin:
            y = vroot_y + margin

        tw.wm_geometry(f"+{x}+{y}")

    def hide(self, event=None) -> None:
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

def abrir_ruta(ruta: str) -> None:
    """Abre un archivo o carpeta utilizando el explorador predeterminado."""
    try:
        if sys.platform.startswith("darwin"):
            subprocess.Popen(["open", ruta])
        elif os.name == "nt":
            os.startfile(ruta)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", ruta])
    except Exception as e:  # pragma: no cover - solo utilitario
        print(f"[ERROR ABRIR] {e}")

def truncate(number, decimals=3):
    """Return *number* truncated to ``decimals`` decimal places."""

    factor = 10 ** decimals
    return math.trunc(number * factor) / factor

campo_resaltado_actual: tk.Widget | None = None
arrow_label: ttk.Label | None = None
pregunta_respondida = False
form_toplevel: tk.Toplevel | None = None
auto_calc_after_id: str | None = None
torque_syncing = False
stator_widgets: dict[str, tuple[ttk.Label, ttk.Entry]] = {}
consecutive_check_labels: dict[str, ttk.Label] = {}


# ---------------------------------------------------------------------------
# Input guidance helpers
# These helpers drive visual focus, completion marks, and lock indicators.
# ---------------------------------------------------------------------------
def resaltar_campo(widget: tk.Widget | None) -> None:
    """Resalta un widget y coloca una flecha de llenado a su lado."""
    global campo_resaltado_actual, arrow_label

    if campo_resaltado_actual is not None:
        try:
            if isinstance(campo_resaltado_actual, ttk.Entry):
                if getattr(campo_resaltado_actual, "_completed", False):
                    campo_resaltado_actual.config(style="Completed.TEntry")
                else:
                    campo_resaltado_actual.config(style="TEntry")
        except tk.TclError:
            pass

    if arrow_label is not None:
        try:
            arrow_label.destroy()
        except tk.TclError:
            pass
        arrow_label = None

    if widget is not None:
        try:
            if isinstance(widget, ttk.Entry):
                widget.config(style="Highlight.TEntry")
            widget.update_idletasks()
            parent = widget.master
            info = widget.grid_info()
            row = info.get("row")
            arrow_label = ttk.Label(parent, text="->", foreground="red")
            arrow_label.grid(row=row, column=4, padx=(5, 0))
        except Exception:
            pass

    campo_resaltado_actual = widget


def marcar_completado(entry: ttk.Entry, label_check: ttk.Label | None = None) -> None:
    """Aplica el estilo de completado y muestra una marca de verificado."""
    try:
        entry.config(style="Completed.TEntry")
        setattr(entry, "_completed", True)
    except tk.TclError:
        pass
    if label_check is not None:
        label_check.config(text="OK", foreground="#1f6f3f")


def bloquear_si_lleno(
    entry: ttk.Entry,
    label_check: ttk.Label | None = None,
    label_lock: ttk.Label | None = None,
) -> None:
    """Bloquea el campo si tiene valor."""

    if entry.get().strip():
        marcar_completado(entry, label_check)
        try:
            entry.config(state="readonly", style="ReadonlyDark.TEntry")
        except tk.TclError:
            pass
        if label_lock is not None:
            label_lock.config(text="LOCK")



def obtener_ruta_configuracion() -> str:
    """Devuelve una ruta de config por usuario, estable para cualquier equipo."""
    appdata = os.getenv("APPDATA")
    if appdata:
        config_dir = os.path.join(appdata, "ABB", "MotorCalc")
    else:
        config_dir = os.path.join(os.path.expanduser("~"), ".abb_motorcalc")
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "mlfb_config.txt")


# Ruta donde se guarda la config (por usuario).
config_file = obtener_ruta_configuracion()
legacy_config_file = "mlfb_config.txt"


# ---------------------------------------------------------------------------
# Excel source discovery and persistence
# Supports: bundled resource (PyInstaller), project paths, and user override.
# ---------------------------------------------------------------------------
def obtener_excel_predeterminado() -> str:
    """Resuelve la ruta por defecto del Excel de MLFB."""
    excel_name = "PMD NEMA V46 ADAPTED APPLICACION V1.xlsx"
    candidates: list[str] = []

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, excel_name))  # type: ignore[attr-defined]

    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(base_dir, excel_name))
    candidates.append(os.path.join(base_dir, "cm anailisis electrico", excel_name))
    candidates.append(os.path.join(os.path.dirname(base_dir), excel_name))
    candidates.append(os.path.join(os.path.dirname(base_dir), "cm anailisis electrico", excel_name))

    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        candidates.append(os.path.join(exe_dir, excel_name))
        candidates.append(os.path.join(exe_dir, "cm anailisis electrico", excel_name))

    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return candidates[0]


# Cargar configuraciÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n al iniciar
def cargar_configuracion():
    excel_default = obtener_excel_predeterminado()
    config_a_cargar = config_file if os.path.exists(config_file) else legacy_config_file
    if os.path.exists(config_a_cargar):
        with open(config_a_cargar, "r") as f:
            lineas = f.read().splitlines()
            if len(lineas) >= 7:
                archivo = lineas[0].strip()
                if not archivo or not os.path.exists(archivo):
                    archivo = excel_default
                return {
                    "archivo": archivo,
                    "col_mlfb": lineas[1],
                    "col_hp": lineas[2],
                    "col_inercia": lineas[3],
                    "col_curvas": lineas[4],
                    "col_hot": lineas[5],
                    "col_cold": lineas[6],
                }
            elif len(lineas) >= 5:
                archivo = lineas[0].strip()
                if not archivo or not os.path.exists(archivo):
                    archivo = excel_default
                return {
                    "archivo": archivo,
                    "col_mlfb": lineas[1],
                    "col_hp": lineas[2],
                    "col_inercia": lineas[3],
                    "col_curvas": lineas[4],
                    "col_hot": "AN",
                    "col_cold": "AO",
                }
    return {
        "archivo": excel_default,
        "col_mlfb": "A",
        "col_hp": "C",
        "col_inercia": "AS",
        "col_curvas": "BT",
        "col_hot": "AN",
        "col_cold": "AO",
    }



def guardar_configuracion(conf):
    with open(config_file, "w") as f:
        f.write(
            f"{conf['archivo']}\n{conf['col_mlfb']}\n{conf['col_hp']}\n{conf['col_inercia']}\n"
            f"{conf['col_curvas']}\n{conf['col_hot']}\n{conf['col_cold']}"
        )



mlfb_config = cargar_configuracion()


# ---------------------------------------------------------------------------
# Parsing and normalization utilities
# Safe converters avoid hard crashes from partial or malformed user inputs.
# ---------------------------------------------------------------------------
def safe_float(value):
    """Convierte un texto a `float devolviendo None si falla."""
    try:
        if isinstance(value, str):
            value = value.strip().replace(",", ".")
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_int(value, default: int | None = None) -> int | None:
    """Convierte un texto a `int` devolviendo default si falla."""

    valor = safe_float(value)
    if valor is None:
        return default
    try:
        return int(round(valor))
    except (TypeError, ValueError):
        return default


def clamp_torque_pct(valor: float) -> float:
    """Clamp torque percentage to the [0, 100] range."""

    return max(0.0, min(100.0, valor))


def obtener_factor_estator() -> tuple[float | None, str]:
    """Devuelve el factor del estator activo y su tipo."""
    tipo = "std"
    if "var_stator_tipo" in globals():
        try:
            tipo = str(var_stator_tipo.get() or "std")
        except Exception:
            tipo = "std"
    entry_map = {
        "std": "entry_std",
        "ranura_profunda": "entry_ranura_profunda",
        "doble_cuadro": "entry_doble_cuadro",
    }
    entry_name = entry_map.get(tipo, "entry_std")
    valor = None
    if entry_name in globals():
        try:
            valor = safe_float(globals()[entry_name].get())
        except Exception:
            valor = None
    return valor, tipo


def actualizar_selector_estator(*_args) -> None:
    """Muestra solo el campo del tipo de estator seleccionado."""
    if "var_stator_tipo" not in globals():
        return
    tipo = str(var_stator_tipo.get() or "std")
    if "stator_widgets" in globals() and stator_widgets:
        for clave, widgets in stator_widgets.items():
            try:
                etiqueta, entry = widgets
            except Exception:
                continue
            if clave == tipo:
                try:
                    etiqueta.grid()
                    entry.grid()
                    entry.config(state="normal", style="TEntry")
                    if clave in consecutive_check_labels:
                        consecutive_check_labels[clave].grid()
                except Exception:
                    pass
            else:
                try:
                    etiqueta.grid_remove()
                    entry.grid_remove()
                    if clave in consecutive_check_labels:
                        consecutive_check_labels[clave].grid_remove()
                except Exception:
                    pass
    else:
        entry_map = {
            "std": "entry_std",
            "ranura_profunda": "entry_ranura_profunda",
            "doble_cuadro": "entry_doble_cuadro",
        }
        for clave, nombre in entry_map.items():
            if nombre not in globals():
                continue
            entry = globals()[nombre]
            try:
                if clave == tipo:
                    entry.config(state="normal", style="TEntry")
                else:
                    entry.config(state="disabled", style="ReadonlyDark.TEntry")
            except Exception:
                pass
    if "var_consecutivos_activo" in globals() and var_consecutivos_activo.get():
        if "actualizar_estado_consecutivos" in globals():
            actualizar_estado_consecutivos()
        programar_autocalculo()


def _crear_fila_factor_estator(
    frame: ttk.Frame,
    texto: str,
    row: int,
    var_tipo: str,
) -> tuple[ttk.Label, ttk.Entry]:
    """Crea una fila (label+entry) para factor de estator."""
    etiqueta = ttk.Label(frame, text=texto, style="FieldLabel.TLabel")
    etiqueta.grid(row=row, column=0, padx=10, pady=5, sticky="w")
    entry = crear_entry(frame, row)
    if "stator_widgets" in globals():
        stator_widgets[var_tipo] = (etiqueta, entry)
    return etiqueta, entry


def _crear_check_consecutivo(frame: ttk.Frame, row: int, key: str) -> ttk.Label:
    """Crea y registra una etiqueta de estado OK para campos consecutivos."""
    lbl = ttk.Label(frame, text="")
    lbl.grid(row=row, column=3, padx=5)
    consecutive_check_labels[key] = lbl
    return lbl


def actualizar_estado_consecutivos() -> None:
    """Aplica check/verde y resalta el siguiente campo pendiente en consecutivos."""
    if "var_consecutivos_activo" not in globals() or not var_consecutivos_activo.get():
        return

    selected_stator = str(var_stator_tipo.get() or "std")
    hidden_factor_keys = {
        "std",
        "ranura_profunda",
        "doble_cuadro",
    } - {selected_stator}

    ordered_required = [
        "entry_f",
        "entry_gfe",
        "entry_diam_rotor",
        "entry_long_paquete",
        "entry_num_arranques",
        "entry_arranques_hora",
    ]
    selected_factor = {
        "std": "entry_std",
        "ranura_profunda": "entry_ranura_profunda",
        "doble_cuadro": "entry_doble_cuadro",
    }.get(str(var_stator_tipo.get() or "std"), "entry_std")
    ordered_required.insert(4, selected_factor)

    first_pending: ttk.Entry | None = None

    for key, lbl in consecutive_check_labels.items():
        entry = globals().get(key)
        if not isinstance(entry, ttk.Entry):
            continue
        # Solo ocultamos checks para factores de estator no seleccionados.
        # El resto de campos puede no estar "viewable" instantaneamente al hacer pack,
        # pero igual deben participar para mantener el resaltado guiado.
        if key in hidden_factor_keys:
            lbl.config(text="")
            continue

        text = entry.get().strip()
        if text:
            marcar_completado(entry, lbl)
        else:
            try:
                entry.config(style="TEntry")
            except Exception:
                pass
            setattr(entry, "_completed", False)
            lbl.config(text="", foreground=UI_THEME["text"])
            if first_pending is None and key in ordered_required:
                first_pending = entry

    if first_pending is not None:
        resaltar_campo(first_pending)


def obtener_factor_torque() -> tuple[float, float]:
    """Return torque percentage and its squared factor."""

    valor = None
    if "var_torque_pct" in globals():
        try:
            valor = float(var_torque_pct.get())
        except Exception:
            valor = None
    if valor is None and "entry_torque_pct" in globals():
        valor = safe_float(entry_torque_pct.get())
    if valor is None:
        valor = 100.0
    valor = clamp_torque_pct(valor)
    factor = (valor / 100.0) ** 2
    return valor, factor


def establecer_torque_pct(valor: float, marcar: bool = False) -> None:
    """Sync torque percentage widgets and optional checkmark."""

    global torque_syncing
    valor = clamp_torque_pct(valor)
    valor_int = int(round(valor))
    torque_syncing = True
    try:
        if "var_torque_pct" in globals():
            try:
                var_torque_pct.set(valor_int)
            except Exception:
                pass
        if "var_torque_pct_str" in globals():
            try:
                var_torque_pct_str.set(str(valor_int))
            except Exception:
                pass
        if "label_check_torque_pct" in globals() and "entry_torque_pct" in globals():
            try:
                if marcar:
                    marcar_completado(entry_torque_pct, label_check_torque_pct)
                else:
                    label_check_torque_pct.config(text="")
                    setattr(entry_torque_pct, "_completed", False)
            except Exception:
                pass
    finally:
        torque_syncing = False


def programar_autocalculo(event=None) -> None:
    """Schedule a silent recalculation after a short delay."""

    global auto_calc_after_id
    if form_toplevel is None:
        return
    if auto_calc_after_id is not None:
        try:
            form_toplevel.after_cancel(auto_calc_after_id)
        except Exception:
            pass
    auto_calc_after_id = form_toplevel.after(
        300, lambda: calcular_datos(silencioso=True, actualizar_botones=False)
    )


def actualizar_estado_arranque(texto: str, color: str = "#444444") -> None:
    """Update the on-screen start-up evaluation label."""

    if "var_eval_text" not in globals() or "label_eval" not in globals():
        return
    var_eval_text.set(texto)
    try:
        color_l = (color or "").lower()
        style_name = "Status.TLabel"
        if color_l in {"#2e7d32", "#1f6f3f"}:
            style_name = "StatusOk.TLabel"
        elif color_l in {"#d4a300", "#8a5a00"}:
            style_name = "StatusWarn.TLabel"
        elif color_l in {"#c00000", "#9a1a1a"}:
            style_name = "StatusBad.TLabel"
        elif color_l in {"#0074c2", "#0b5cad"}:
            style_name = "StatusInfo.TLabel"
        label_eval.config(style=style_name)
    except Exception:
        pass


def limpiar_entry(campo: ttk.Entry) -> None:
    """Clear an entry and leave it read-only."""

    campo.config(state="normal")
    campo.delete(0, tk.END)
    campo.config(state="readonly", style="ReadonlyDark.TEntry")


def generar_series_arranques(
    num_arranques: int,
    v2a_ceiling: float,
    v2aus_ceiling: float | None,
    ta_val: float,
    taus_ceiling: float,
    std_val: float,
    max_temp: float,
) -> tuple[list[tuple[str, float, float, str]], list[float]]:
    """Genera filas de arranque y lista de temperaturas para la grÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡fica."""

    filas: list[tuple[str, float, float, str]] = []
    temps: list[float] = []
    prev_cool: float | None = None

    for i in range(1, num_arranques + 1):
        if i == 1:
            temp_heat = v2a_ceiling
        else:
            temp_heat = (prev_cool or 0.0) + v2a_ceiling
        filas.append((str(i), temp_heat, ta_val, "calentamiento"))
        temps.append(temp_heat)

        if i == 1 and v2aus_ceiling is not None:
            temp_cool = v2aus_ceiling
        else:
            temp_cool = temp_heat * std_val
        filas.append(("", temp_cool, taus_ceiling, "enfriamiento"))
        temps.append(temp_cool)
        prev_cool = temp_cool

    return filas, temps


def estados_progresivos(filas: list[tuple[str, float, float, str]], max_temp: float) -> list[str]:
    """Asigna estado progresivo: pasa de OK a Cerca y luego a Excede sin volver atrÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡s."""
    estados: list[str] = []
    estado_actual = "OK"
    for _, temp, _, _ in filas:
        if max_temp > 0:
            ratio = temp / max_temp
            if ratio >= 1.0:
                estado_actual = "Excede"
            elif ratio >= ESTADO_UMBRAL and estado_actual == "OK":
                estado_actual = "Cerca"
        estados.append(estado_actual)
    return estados


def calcular_arranques_por_hora() -> tuple[list[tuple[str, float, float]], float, float, float, float]:
    """Devuelve filas (arr, t, temp) y mÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©tricas de ciclo para arranques por hora."""
    if "entry_arranques_hora" not in globals():
        return [], 0, 0, 0, 0
    n_hora = max(1, safe_int(entry_arranques_hora.get(), default=4))
    ta_c = safe_float(entry_ta_ceiling.get())
    taus_c = safe_float(entry_taus_ceiling.get())
    v2a_ceiling_val = safe_float(entry_v2a_ceiling.get())
    factor_estator, tipo_estator = obtener_factor_estator()
    if None in (ta_c, taus_c, v2a_ceiling_val, factor_estator):
        return [], 0, 0, 0, 0

    ciclo = ta_c + taus_c
    resto = max(0, 3600 - ciclo * n_hora)
    reposo_ciclo = resto / n_hora if n_hora else 0
    b5_ref = ta_c + taus_c + reposo_ciclo

    filas: list[tuple[str, float, float]] = []
    # Fila inicial (B2=0, C2=0)
    filas.append(("", 0.0, 0.0))
    prev_rest_t = 0.0
    prev_rest_temp = 0.0
    debug_print("arranques_hora", "valor usuario", n_hora)
    debug_print("ciclo_arr", f"{ta_c} + {taus_c}", ciclo)
    debug_print("resto_reposo", f"3600 - ({ciclo} * {n_hora})", resto)
    debug_print("reposo_ciclo", f"{resto} / {n_hora}", reposo_ciclo)

    for i in range(1, n_hora + 1):
        # Calentamiento: B3/B6/B9..., C3/C6/C9...
        t_heat = prev_rest_t + ta_c
        temp_heat = v2a_ceiling_val + prev_rest_temp
        debug_print(f"arr{i}_t_heat", f"{prev_rest_t} + {ta_c}", t_heat)
        debug_print(f"arr{i}_temp_heat", f"{prev_rest_temp} + {v2a_ceiling_val}", temp_heat)
        filas.append((str(i), t_heat, temp_heat))

        # Enfriamiento: B4/B7/B10..., C4/C7/C10...
        t_cool = t_heat + taus_c
        temp_cool = temp_heat * factor_estator
        debug_print(f"arr{i}_t_cool", f"{t_heat} + {taus_c}", t_cool)
        debug_print(
            f"arr{i}_temp_cool",
            f"{temp_heat} * {factor_estator} ({tipo_estator})",
            temp_cool,
        )
        filas.append(("", t_cool, temp_cool))

        # Reposo: B5/B8/B11..., C5/C8/C11...
        t_rest_end = t_cool + reposo_ciclo
        # Coincide con hoja: descuento termico fijo usando referencia $B$5/60
        temp_rest = max(0.0, temp_cool - (b5_ref / 60.0))
        debug_print(f"arr{i}_t_reposo", f"{t_cool} + {reposo_ciclo}", t_rest_end)
        debug_print(f"arr{i}_temp_reposo", f"{temp_cool} - ({b5_ref}/60)", temp_rest)
        filas.append(("", t_rest_end, temp_rest))

        prev_rest_t = t_rest_end
        prev_rest_temp = temp_rest

    return filas, n_hora, ciclo, resto, reposo_ciclo


def generar_series_arranques_caliente(
    num_arranques: int,
    v2a_ceiling: float,
    v2aus_ceiling: float,
    ta_val: float,
    taus_ceiling: float,
    temp_inicial: float,
    tx_val: float,
    max_temp: float,
) -> tuple[list[tuple[str, float, float, str]], list[float]]:
    """Genera filas de arranques en caliente y lista de temperaturas."""

    filas: list[tuple[str, float, float, str]] = []
    temps: list[float] = []

    # En frio (referencia inicial)
    filas.append(("En frio", v2a_ceiling, ta_val, "calentamiento"))
    temps.append(v2a_ceiling)
    filas.append(("", v2aus_ceiling, taus_ceiling, "enfriamiento"))
    temps.append(v2aus_ceiling)

    # En caliente (fase inicial)
    temp_hot = temp_inicial + v2a_ceiling
    filas.append(("En caliente", temp_hot, tx_val, "calentamiento"))
    temps.append(temp_hot)
    temp_cool = temp_hot - v2aus_ceiling
    filas.append(("", temp_cool, taus_ceiling, "enfriamiento"))
    temps.append(temp_cool)

    prev_cool = temp_cool
    for i in range(1, num_arranques + 1):
        temp_heat = prev_cool + v2a_ceiling
        filas.append((str(i), temp_heat, ta_val, "calentamiento"))
        temps.append(temp_heat)
        temp_cool = temp_heat - v2aus_ceiling
        filas.append(("", temp_cool, taus_ceiling, "enfriamiento"))
        temps.append(temp_cool)
        prev_cool = temp_cool

    return filas, temps


def actualizar_tabla_arranques(
    tree: ttk.Treeview,
    filas: list[tuple[str, float, float, str]],
    max_temp: float,
) -> None:
    """Actualiza una tabla con filas de arranques, aplicando color por estado."""

    tree.tag_configure("OK", background=BANDA_OK)
    tree.tag_configure("Cerca", background=BANDA_ALERTA)
    tree.tag_configure("Excede", background=BANDA_PELIGRO)

    tree.delete(*tree.get_children())
    estados = estados_progresivos(filas, max_temp)
    debug_print(
        "tabla_arranques_total",
        f"filas={len(filas)} max_temp={formato_valor_pdf(max_temp)}",
        len(filas),
    )
    for (arr, temp, seg, tipo), estado in zip(filas, estados):
        debug_print(
            "tabla_arranques_fila",
            f"arr={arr or '-'} temp={formato_valor_pdf(temp)} seg={formato_valor_pdf(seg)} tipo={tipo} estado={estado}",
            temp,
        )
        tree.insert(
            "",
            "end",
            values=(
                arr,
                formato_valor_pdf(temp),
                formato_valor_pdf(seg),
                tipo,
            ),
            tags=(estado,),
        )


def dibujar_grafica_arranques() -> None:
    """Dibuja la grafica de arranques en el canvas."""

    if "canvas_arranques" not in globals():
        return

    canvas = canvas_arranques
    canvas.delete("all")

    if not var_consecutivos_activo.get():
        return

    v2a_ceiling_val = safe_float(entry_v2a_ceiling.get())
    v2aus_ceiling_val = safe_float(entry_v2aus_ceiling.get())
    ta_val = safe_float(entry_ta.get())
    taus_ceiling_val = safe_float(entry_taus_ceiling.get())
    factor_estator, tipo_estator = obtener_factor_estator()

    num_arranques = safe_int(entry_num_arranques.get(), default=10) if "entry_num_arranques" in globals() else 10
    num_arranques = max(1, num_arranques)
    max_temp = TEMP_MAX_FRIO

    if None in (v2a_ceiling_val, ta_val, taus_ceiling_val, factor_estator):
        canvas.create_text(10, 10, anchor="nw", text="Datos insuficientes para graficar / Insufficient data to plot.")
        return

    filas, temps = generar_series_arranques(
        num_arranques,
        v2a_ceiling_val,
        v2aus_ceiling_val,
        ta_val,
        taus_ceiling_val,
        factor_estator,
        max_temp,
    )

    if not temps:
        return

    # cortar la serie en la primera superaciÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n de lÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â­mite
    temps_plot = temps
    estados_plot = estados_progresivos([(str(i), t, 0.0, "") for i, t in enumerate(temps)], max_temp)
    for idx, est in enumerate(estados_plot):
        if est == "Excede":
            temps_plot = temps[: idx + 1]
            break

    max_val = max(temps)
    estado = obtener_estado_temperatura(max_temp, max_val)
    color_line = COLOR_ESTADO.get(estado, "#0b4f8a")
    debug_print(
        "grafica_caliente_entrada",
        f"num_arr={num_arranques}, v2a={formato_valor_pdf(v2a_ceiling_val)}, v2aus={formato_valor_pdf(v2aus_ceiling_val)}, ta={formato_valor_pdf(ta_val)}, taus={formato_valor_pdf(taus_ceiling_val)}",
        len(temps_plot),
    )
    debug_print(
        "grafica_frio_entrada",
        f"num_arr={num_arranques}, v2a={formato_valor_pdf(v2a_ceiling_val)}, v2aus={formato_valor_pdf(v2aus_ceiling_val)}, ta={formato_valor_pdf(ta_val)}, taus={formato_valor_pdf(taus_ceiling_val)}, factor={formato_valor_pdf(factor_estator)} ({tipo_estator})",
        len(temps_plot),
    )

    canvas.update_idletasks()
    width = max(1, canvas.winfo_width())
    height = max(1, canvas.winfo_height())
    left, right, top, bottom = 55, 20, 30, 35
    plot_w = width - left - right
    plot_h = height - top - bottom
    if plot_w <= 0 or plot_h <= 0:
        return

    y_max = max(max_temp, max(temps_plot)) * 1.1
    if y_max <= 0:
        y_max = 1
    y_min = 0.0
    debug_print("grafica_caliente_ymax", f"MAX({formato_valor_pdf(max_temp)}, MAX(temps_plot)) * 1.1", y_max)
    debug_print("grafica_frio_ymax", f"MAX({formato_valor_pdf(max_temp)}, MAX(temps_plot)) * 1.1", y_max)

    def to_xy(idx: int, temp: float) -> tuple[float, float]:
        x = left + (plot_w * idx / max(1, len(temps_plot) - 1))
        y = height - bottom - ((temp - y_min) / (y_max - y_min)) * plot_h
        return x, y

    # Zonas (seguro/alerta/peligro)
    if max_temp > 0:
        y_warn = height - bottom - ((max_temp * ESTADO_UMBRAL - y_min) / (y_max - y_min)) * plot_h
        y_limit = height - bottom - ((max_temp - y_min) / (y_max - y_min)) * plot_h
        canvas.create_rectangle(left, top, width - right, y_limit, fill=BANDA_PELIGRO, outline="")
        canvas.create_rectangle(left, y_limit, width - right, y_warn, fill=BANDA_ALERTA, outline="")
        canvas.create_rectangle(left, y_warn, width - right, height - bottom, fill=BANDA_OK, outline="")

    # Fondo y borde
    canvas.create_rectangle(left, top, width - right, height - bottom, fill="", outline="#bbbbbb")
    canvas.create_text(width // 2, top - 12, text="Arranques en Frio / Cold Starts", fill="#333333")

    # Grid y etiquetas Y
    y_ticks = 5
    for i in range(y_ticks + 1):
        y_val = y_min + (y_max - y_min) * i / y_ticks
        y = height - bottom - ((y_val - y_min) / (y_max - y_min)) * plot_h
        canvas.create_line(left, y, width - right, y, fill="#e6e6e6")
        canvas.create_text(left - 6, y, text=formato_valor_pdf(round(y_val, 1)), anchor="e", fill="#555555")

    # Grid y etiquetas X (arranques)
    step = max(1, math.ceil(num_arranques / 6))
    for arr in range(1, num_arranques + 1, step):
        idx = 2 * (arr - 1)
        x = left + (plot_w * idx / max(1, len(temps_plot) - 1))
        canvas.create_line(x, top, x, height - bottom, fill="#f0f0f0")
        canvas.create_text(x, height - bottom + 12, text=str(arr), anchor="n", fill="#555555")

    # Ejes
    canvas.create_line(left, top, left, height - bottom, fill="#333333")
    canvas.create_line(left, height - bottom, width - right, height - bottom, fill="#333333")

    # Linea de temp maxima
    y_line = height - bottom - ((max_temp - y_min) / (y_max - y_min)) * plot_h
    debug_print(
        "grafica_caliente_linea_max",
        f"{height} - {bottom} - (({max_temp} - {y_min}) / ({y_max} - {y_min})) * {plot_h}",
        y_line,
    )
    debug_print(
        "grafica_frio_linea_max",
        f"{height} - {bottom} - (({max_temp} - {y_min}) / ({y_max} - {y_min})) * {plot_h}",
        y_line,
    )
    canvas.create_line(left, y_line, width - right, y_line, fill="#c00000", dash=(4, 2))
    canvas.create_text(width - right, y_line - 8, anchor="ne", text=f"{formato_valor_pdf(max_temp)} degC", fill="#c00000")

    # Serie
    points = [to_xy(i, t) for i, t in enumerate(temps_plot)]
    debug_print("grafica_caliente_puntos", f"len(points)={len(points)}", len(points))
    if points:
        debug_print("grafica_caliente_punto_inicio", f"point0={points[0]}", points[0][1])
        debug_print("grafica_caliente_punto_fin", f"pointN={points[-1]}", points[-1][1])
    debug_print("grafica_frio_puntos", f"len(points)={len(points)}", len(points))
    if points:
        debug_print("grafica_frio_punto_inicio", f"point0={points[0]}", points[0][1])
        debug_print("grafica_frio_punto_fin", f"pointN={points[-1]}", points[-1][1])
    for i in range(len(points) - 1):
        canvas.create_line(points[i][0], points[i][1], points[i + 1][0], points[i + 1][1], fill=color_line, width=2)
        canvas.create_oval(points[i][0] - 2, points[i][1] - 2, points[i][0] + 2, points[i][1] + 2, fill=color_line, outline="")
    canvas.create_oval(points[-1][0] - 2, points[-1][1] - 2, points[-1][0] + 2, points[-1][1] + 2, fill=color_line, outline="")

    # Leyenda debajo de la grÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡fica
    canvas.create_text(width // 2, height - 26, text=f"Estado/Status: {estado}", fill=color_line)

    canvas.create_text(width // 2, height - 12, text="arranques / starts")
    canvas.create_text(12, top, text="degC", anchor="nw")


def dibujar_grafica_arranques_caliente() -> None:
    """Dibuja la grafica de arranques en caliente en el canvas."""

    if "canvas_arranques_caliente" not in globals():
        return

    canvas = canvas_arranques_caliente
    canvas.delete("all")

    if not var_consecutivos_activo.get():
        return

    v2a_ceiling_val = safe_float(entry_v2a_ceiling.get())
    v2aus_ceiling_val = safe_float(entry_v2aus_ceiling.get())
    ta_val = safe_float(entry_ta.get())
    taus_ceiling_val = safe_float(entry_taus_ceiling.get())

    num_arranques = safe_int(entry_num_arranques.get(), default=10) if "entry_num_arranques" in globals() else 10
    num_arranques = max(1, num_arranques)
    max_temp = TEMP_MAX_CALIENTE
    temp_inicial = TEMP_INICIAL_CALIENTE
    tx_val = TX_CALIENTE

    if None in (v2a_ceiling_val, v2aus_ceiling_val, ta_val, taus_ceiling_val):
        canvas.create_text(10, 10, anchor="nw", text="Datos insuficientes para graficar / Insufficient data to plot.")
        return

    filas, temps = generar_series_arranques_caliente(
        num_arranques,
        v2a_ceiling_val,
        v2aus_ceiling_val,
        ta_val,
        taus_ceiling_val,
        temp_inicial,
        tx_val,
        max_temp,
    )

    if not temps:
        return

    temps_plot = temps
    estados_plot = estados_progresivos([(str(i), t, 0.0, "") for i, t in enumerate(temps)], max_temp)
    for idx, est in enumerate(estados_plot):
        if est == "Excede":
            temps_plot = temps[: idx + 1]
            break

    max_val = max(temps)
    estado = obtener_estado_temperatura(max_temp, max_val)
    color_line = COLOR_ESTADO.get(estado, "#0b4f8a")

    canvas.update_idletasks()
    width = max(1, canvas.winfo_width())
    height = max(1, canvas.winfo_height())
    left, right, top, bottom = 55, 20, 30, 35
    plot_w = width - left - right
    plot_h = height - top - bottom
    if plot_w <= 0 or plot_h <= 0:
        return

    y_max = max(max_temp, max(temps_plot)) * 1.1
    if y_max <= 0:
        y_max = 1
    y_min = 0.0

    def to_xy(idx: int, temp: float) -> tuple[float, float]:
        x = left + (plot_w * idx / max(1, len(temps_plot) - 1))
        y = height - bottom - ((temp - y_min) / (y_max - y_min)) * plot_h
        return x, y

    # Zonas (seguro/alerta/peligro)
    if max_temp > 0:
        y_warn = height - bottom - ((max_temp * ESTADO_UMBRAL - y_min) / (y_max - y_min)) * plot_h
        y_limit = height - bottom - ((max_temp - y_min) / (y_max - y_min)) * plot_h
        canvas.create_rectangle(left, top, width - right, y_limit, fill=BANDA_PELIGRO, outline="")
        canvas.create_rectangle(left, y_limit, width - right, y_warn, fill=BANDA_ALERTA, outline="")
        canvas.create_rectangle(left, y_warn, width - right, height - bottom, fill=BANDA_OK, outline="")

    # Fondo y borde
    canvas.create_rectangle(left, top, width - right, height - bottom, fill="", outline="#bbbbbb")
    canvas.create_text(width // 2, top - 12, text="Arranques en Caliente / Hot Starts", fill="#333333")

    # Grid y etiquetas Y
    y_ticks = 5
    for i in range(y_ticks + 1):
        y_val = y_min + (y_max - y_min) * i / y_ticks
        y = height - bottom - ((y_val - y_min) / (y_max - y_min)) * plot_h
        canvas.create_line(left, y, width - right, y, fill="#e6e6e6")
        canvas.create_text(left - 6, y, text=formato_valor_pdf(round(y_val, 1)), anchor="e", fill="#555555")

    # Grid y etiquetas X (tiempo)
    step = max(1, math.ceil(len(temps_plot) / 6))
    for idx in range(0, len(temps_plot), step):
        x = left + (plot_w * idx / max(1, len(temps_plot) - 1))
        canvas.create_line(x, top, x, height - bottom, fill="#f0f0f0")
        canvas.create_text(x, height - bottom + 12, text=str(idx + 1), anchor="n", fill="#555555")

    # Ejes
    canvas.create_line(left, top, left, height - bottom, fill="#333333")
    canvas.create_line(left, height - bottom, width - right, height - bottom, fill="#333333")

    # Linea de temp maxima
    y_line = height - bottom - ((max_temp - y_min) / (y_max - y_min)) * plot_h
    canvas.create_line(left, y_line, width - right, y_line, fill="#c00000", dash=(4, 2))
    canvas.create_text(width - right, y_line - 8, anchor="ne", text=f"{formato_valor_pdf(max_temp)} degC", fill="#c00000")

    # Serie
    points = [to_xy(i, t) for i, t in enumerate(temps_plot)]
    for i in range(len(points) - 1):
        canvas.create_line(points[i][0], points[i][1], points[i + 1][0], points[i + 1][1], fill=color_line, width=2)
        canvas.create_oval(points[i][0] - 2, points[i][1] - 2, points[i][0] + 2, points[i][1] + 2, fill=color_line, outline="")
    canvas.create_oval(points[-1][0] - 2, points[-1][1] - 2, points[-1][0] + 2, points[-1][1] + 2, fill=color_line, outline="")

    # Leyenda
    legend_x = width - right - 120
    legend_y = top + 6
    # Leyenda debajo de la grÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡fica
    canvas.create_text(width // 2, height - 26, text=f"Estado/Status: {estado}", fill=color_line)

    canvas.create_text(width // 2, height - 12, text="tiempo / time")
    canvas.create_text(12, top, text="degC", anchor="nw")


def dibujar_grafica_arranques_hora() -> None:
    if "canvas_arranques_hora" not in globals():
        return
    canvas = canvas_arranques_hora
    canvas.delete("all")

    filas, n_hora, ciclo, resto, reposo_ciclo = calcular_arranques_por_hora()
    if not filas:
        canvas.create_text(10, 10, anchor="nw", text="Datos insuficientes para graficar / Insufficient data to plot.")
        return

    temps = [float(t) for _, _, t in filas]
    xs = [float(t) for _, t, _ in filas]
    max_temp = max(temps) if temps else 1
    min_temp = min(temps) if temps else 0

    width = max(1, canvas.winfo_width())
    height = max(1, canvas.winfo_height())
    left, right, top, bottom = 55, 20, 20, 30
    plot_w = width - left - right
    plot_h = height - top - bottom
    if plot_w <= 0 or plot_h <= 0:
        return

    y_max = max_temp * 1.1 if max_temp > 0 else 1
    y_min = min_temp * 0.9 if min_temp > 0 else 0
    x_max = max(xs) if xs else 3600
    debug_print(
        "grafica_hora_escalas",
        f"n_hora={n_hora}, ciclo={formato_valor_pdf(ciclo)}, resto={formato_valor_pdf(resto)}, reposo_ciclo={formato_valor_pdf(reposo_ciclo)}",
        x_max,
    )
    debug_print(
        "grafica_hora_rangos",
        f"x_max={formato_valor_pdf(x_max)}, y_min={formato_valor_pdf(y_min)}, y_max={formato_valor_pdf(y_max)}",
        y_max,
    )

    def to_xy(t_val: float, temp: float) -> tuple[float, float]:
        x = left + (plot_w * t_val / max(1, x_max))
        y = height - bottom - ((temp - y_min) / max(1e-6, (y_max - y_min))) * plot_h
        return x, y

    # ejes y fondo
    canvas.create_rectangle(left, top, width - right, height - bottom, outline="#bbbbbb")
    # y ticks
    y_ticks = 4
    for i in range(y_ticks + 1):
        val = y_min + (y_max - y_min) * i / y_ticks
        y = height - bottom - ((val - y_min) / (y_max - y_min)) * plot_h
        canvas.create_line(left, y, width - right, y, fill="#e6e6e6")
        canvas.create_text(left - 6, y, text=formato_valor_pdf(round(val, 1)), anchor="e", fill="#555")

    # x ticks (0..3600)
    step = max(1, int(x_max // 6))
    for xv in range(0, int(x_max) + 1, step):
        x = left + (plot_w * xv / max(1, x_max))
        canvas.create_line(x, top, x, height - bottom, fill="#f0f0f0")
        canvas.create_text(x, height - bottom + 12, text=str(xv), anchor="n", fill="#555")

    # serie
    pts = [to_xy(t, temp) for _, t, temp in filas]
    debug_print("grafica_hora_puntos", f"len(points)={len(pts)}", len(pts))
    if pts:
        debug_print("grafica_hora_punto_inicio", f"point0={pts[0]}", pts[0][1])
        debug_print("grafica_hora_punto_fin", f"pointN={pts[-1]}", pts[-1][1])
    if len(pts) >= 2:
        for i in range(len(pts) - 1):
            canvas.create_line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], fill="#0b4f8a", width=2)
    for x, y in pts:
        canvas.create_oval(x - 2, y - 2, x + 2, y + 2, fill="#0b4f8a", outline="")

    canvas.create_text(width // 2, top - 6, text=f"Grafica para {n_hora} arranques por hora / Chart for {n_hora} starts per hour", fill="#333")
    canvas.create_text(width // 2, height - 12, text="Tiempo (seg) / Time (sec)", fill="#333")
    canvas.create_text(12, top, text="degC", anchor="nw", fill="#333")
def calcular_arranque_consecutivo() -> None:
    """Compute optional consecutive start fields when enabled."""

    if "var_consecutivos_activo" not in globals() or not var_consecutivos_activo.get():
        return

    diam = safe_float(entry_diam_rotor.get())
    long_paquete = safe_float(entry_long_paquete.get())
    gfe = safe_float(entry_gfe.get())
    f_val = safe_float(entry_f.get())
    mmv_nmot = safe_float(entry_mmv_nmot.get())
    ngd2 = safe_float(entry_ngd2.get())
    rpm_motor = safe_float(entry_rpm_motor.get())
    mmm_val = safe_float(entry_mmm.get())
    ta_val = safe_float(entry_ta.get())
    factor_estator, tipo_estator = obtener_factor_estator()

    if diam is not None and long_paquete is not None:
        diamxlong = (diam - 0.1) * 3.1416 * 0.1 * 7.8 * long_paquete * 0.945
        debug_print(
            "diamxlongitud",
            f"({diam} - 0.1) * 3.1416 * 0.1 * 7.8 * {long_paquete} * 0.945",
            diamxlong,
        )
        actualizar_campo(entry_diamxlongitud, diamxlong)
    else:
        limpiar_entry(entry_diamxlongitud)

    if (
        ngd2 is not None
        and rpm_motor is not None
        and mmm_val is not None
        and f_val is not None
        and gfe is not None
        and gfe != 0
        and mmv_nmot is not None
        and (mmm_val - mmv_nmot) != 0
    ):
        v2a_raw = (
            (ngd2 * (rpm_motor**2) * mmm_val * f_val)
            / (730 * (mmm_val - mmv_nmot) * 460 * gfe)
        )
        v2a = round(v2a_raw, 2)
        debug_print(
            "V'2A",
            f"({ngd2} * ({rpm_motor}^2) * {mmm_val} * {f_val}) / "
            f"(730 * ({mmm_val} - {mmv_nmot}) * 460 * {gfe})",
            v2a,
        )
        actualizar_campo(entry_v2a, v2a)
        v2a_ceiling = round(v2a, 0)
        debug_print("V'2A Ceiling", f"ROUND({v2a}, 0)", v2a_ceiling)
        actualizar_campo(entry_v2a_ceiling, v2a_ceiling)
    else:
        limpiar_entry(entry_v2a)
        limpiar_entry(entry_v2a_ceiling)

    if ngd2 is not None and rpm_motor is not None and mmv_nmot is not None and mmv_nmot != 0:
        taus = round((ngd2 * rpm_motor) / (375 * mmv_nmot), 2)
        debug_print(
            "taus",
            f"({ngd2} * {rpm_motor}) / (375 * {mmv_nmot})",
            taus,
        )
        actualizar_campo(entry_taus, taus)
        taus_ceiling = round(taus, 0)
        debug_print("taus ceiling", f"ROUND({taus}, 0)", taus_ceiling)
        actualizar_campo(entry_taus_ceiling, taus_ceiling)
    else:
        limpiar_entry(entry_taus)
        limpiar_entry(entry_taus_ceiling)

    if (
        safe_float(entry_taus_ceiling.get()) is not None
        and ta_val is not None
        and ta_val != 0
    ):
        taus_ta = safe_float(entry_taus_ceiling.get()) / ta_val
        debug_print("Taus/Ta", f"{entry_taus_ceiling.get()} / {ta_val}", taus_ta)
        actualizar_campo(entry_taus_ta, taus_ta)
    else:
        limpiar_entry(entry_taus_ta)

    v2a_ceiling_val = safe_float(entry_v2a_ceiling.get())
    if factor_estator is not None and v2a_ceiling_val is not None:
        v2aus = factor_estator * v2a_ceiling_val
        debug_print("V2aus", f"{factor_estator} * {v2a_ceiling_val} ({tipo_estator})", v2aus)
        actualizar_campo(entry_v2aus, v2aus)
        v2aus_ceiling = math.floor(v2aus)
        debug_print("V2aus ceiling", f"FLOOR({v2aus}, 1)", v2aus_ceiling)
        actualizar_campo(entry_v2aus_ceiling, v2aus_ceiling)
    else:
        limpiar_entry(entry_v2aus)
        limpiar_entry(entry_v2aus_ceiling)

    # Tablas en pantalla
    if "tree_arranques" in globals():
        if None not in (
            v2a_ceiling_val,
            ta_val,
            safe_float(entry_taus_ceiling.get()),
            factor_estator,
        ):
            filas_tabla, _ = generar_series_arranques(
                max(1, safe_int(entry_num_arranques.get(), default=10)),
                v2a_ceiling_val,
                safe_float(entry_v2aus_ceiling.get()),
                ta_val,
                safe_float(entry_taus_ceiling.get()),
                factor_estator,
                TEMP_MAX_FRIO,
            )
            actualizar_tabla_arranques(tree_arranques, filas_tabla, TEMP_MAX_FRIO)
        else:
            tree_arranques.delete(*tree_arranques.get_children())

    if "tree_arranques_caliente" in globals():
        v2aus_ceiling_val = safe_float(entry_v2aus_ceiling.get())
        taus_ceiling_val = safe_float(entry_taus_ceiling.get())
        num_arr_hot = max(1, safe_int(entry_num_arranques.get(), default=10))
        temp_inicial = TEMP_INICIAL_CALIENTE
        tx_val = TX_CALIENTE
        if None not in (v2a_ceiling_val, v2aus_ceiling_val, ta_val, taus_ceiling_val):
            filas_hot, _ = generar_series_arranques_caliente(
                num_arr_hot,
                v2a_ceiling_val,
                v2aus_ceiling_val,
                ta_val,
                taus_ceiling_val,
                temp_inicial,
                tx_val,
                TEMP_MAX_CALIENTE,
            )
            actualizar_tabla_arranques(tree_arranques_caliente, filas_hot, TEMP_MAX_CALIENTE)
        else:
            tree_arranques_caliente.delete(*tree_arranques_caliente.get_children())

    # Arranques por hora
    if "tree_arranques_hora" in globals():
        filas_hora, n_hora, ciclo, resto, reposo_ciclo = calcular_arranques_por_hora()
        tree_arranques_hora.delete(*tree_arranques_hora.get_children())
        debug_print("tabla_hora_total", f"filas={len(filas_hora)}", len(filas_hora))
        for arr, t_abs, temp_val in filas_hora:
            debug_print(
                "tabla_hora_fila",
                f"arr={arr} t={formato_valor_pdf(t_abs)} temp={formato_valor_pdf(temp_val)}",
                temp_val,
            )
            tree_arranques_hora.insert(
                "",
                "end",
                values=(arr, formato_valor_pdf(t_abs), formato_valor_pdf(temp_val)),
            )
        if "entry_ciclo_arr" in globals():
            entry_ciclo_arr.config(state="normal")
            entry_ciclo_arr.delete(0, tk.END)
            entry_ciclo_arr.insert(0, formato_valor_pdf(ciclo))
            entry_ciclo_arr.config(state="readonly", style="ReadonlyDark.TEntry")
        if "entry_resto_arr" in globals():
            entry_resto_arr.config(state="normal")
            entry_resto_arr.delete(0, tk.END)
            entry_resto_arr.insert(0, formato_valor_pdf(resto))
            entry_resto_arr.config(state="readonly", style="ReadonlyDark.TEntry")
        if "entry_reposo_ciclo" in globals():
            entry_reposo_ciclo.config(state="normal")
            entry_reposo_ciclo.delete(0, tk.END)
            entry_reposo_ciclo.insert(0, formato_valor_pdf(reposo_ciclo))
            entry_reposo_ciclo.config(state="readonly", style="ReadonlyDark.TEntry")

    if "canvas_arranques_hora" in globals():
        dibujar_grafica_arranques_hora()

    if "actualizar_estado_consecutivos" in globals():
        actualizar_estado_consecutivos()

    dibujar_grafica_arranques()
    dibujar_grafica_arranques_caliente()


def sync_torque_from_str(*_args) -> None:
    """Sync scale when the user types the torque percentage."""

    global torque_syncing
    if torque_syncing:
        return
    if "var_torque_pct_str" not in globals():
        return
    texto = var_torque_pct_str.get().strip()
    if not texto:
        return
    texto = texto.replace(",", ".")
    try:
        valor = safe_float(texto)
        if valor is None:
            return None
    except ValueError:
        return
    establecer_torque_pct(valor, marcar=False)


def calcular_campos(vars_dict: dict) -> None:
    """Calcula los valores de adaptaciÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n para un motor y actualiza los
    ``StringVar`` del diccionario.

    ``vars_dict`` debe incluir las claves ``rpm_eq``, ``wr2``, ``mnv``,
    ``mmv``, ``jv``, ``rpm_motor``, ``jm`` y ``mmm`` ademÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡s de ``is_pump``
    (``BooleanVar``). TambiÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©n debe contener ``mmv_nmot``, ``gd2v``,
    ``gd2v_nmot``, ``ngd2`` y ``ta`` para almacenar los resultados. Se
    lanza ``ValueError`` si falta algÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Âºn dato requerido o si un valor no es
    numÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©rico.
    """

    rpm_eq = safe_float(vars_dict["rpm_eq"].get())
    rpm_motor = safe_float(vars_dict["rpm_motor"].get())
    wr2 = safe_float(vars_dict["wr2"].get())
    mnv = safe_float(vars_dict["mnv"].get())
    jv = safe_float(vars_dict["jv"].get())
    jm = safe_float(vars_dict["jm"].get())
    mmm = safe_float(vars_dict["mmm"].get())

    if None in (rpm_eq, rpm_motor, wr2, mnv, jv, jm, mmm):
        raise ValueError("Ingrese todos los datos numÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©ricos")

    if vars_dict["is_pump"].get():
        mmv = mnv * 0.33
        vars_dict["mmv"].set(f"{mmv:.7f}")
    else:
        mmv = safe_float(vars_dict["mmv"].get())
        if mmv is None:
            raise ValueError("Ingrese MMV")

    gd2m = 4 * jm
    mmv_nmot_raw = mmv * (rpm_eq / rpm_motor)
    mmv_nmot = round(mmv_nmot_raw, 2)
    vars_dict["mmv_nmot"].set(f"{mmv_nmot:.7f}")
    gd2v = 4 * jv
    gd2v_nmot_raw = gd2v * (rpm_eq / rpm_motor) ** 2
    gd2v_nmot = round(gd2v_nmot_raw, 2)
    ngd2 = gd2v_nmot + gd2m

    try:
        ta = (ngd2 * rpm_motor) / (375 * (mmm - mmv_nmot))
    except ZeroDivisionError:
        ta = None

    vars_dict["mmv_nmot"].set(f"{mmv_nmot:.7f}")
    vars_dict["gd2v"].set(f"{gd2v:.7f}")
    vars_dict["gd2v_nmot"].set(f"{gd2v_nmot:.7f}")
    vars_dict["ngd2"].set(f"{ngd2:.7f}")
    vars_dict["ta"].set(f"{ta:.7f}")


def debug_print(etiqueta, operacion, valor):
    """Log a formatted debug line both to stdout and the log file."""

    if not debug_enabled:
        return
    linea = f"[DEBUG] {etiqueta}: {operacion} = {formato_valor_pdf(valor)}"
    print(linea)
    logger.debug(linea)


def leer_flotante(campo):
    """Return ``float`` from an Entry widget or ``None`` if empty."""

    texto = campo.get().strip()
    return safe_float(texto) if texto else None

def leer_valor_unidad(campo, unidad):
    """Lee el valor desde un Entry y lo convierte a sistema mÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©trico si es necesario."""
    texto = campo.get().strip()
    try:
        valor = safe_float(texto)
        if valor is None:
            return None
        if unidad == "lbft2":
            return valor / 23.731
        elif unidad == "lbft":
            return valor / 7.233
        else:
            return valor  # ya estÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ en sistema mÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©trico
    except ValueError:
        return None



def calcular_datos(
    *, silencioso: bool = False, actualizar_botones: bool = True, debug_always: bool = False
):
    """Compute all derived motor parameters and display results."""

    global btn_calcular, btn_debug, debug_enabled
    # Allow debug windows to force logging without triggering interactive dialogs.
    prev_debug_enabled = debug_enabled
    debug_enabled = debug_always or not silencioso
    try:
        global var_ventilador_bomba, entry_factor_mmv, label_check_mmv, label_check_factor_mmv
        if debug_enabled:
            logger.handlers.clear()
            handler = logging.FileHandler("debug_log.txt", mode="w", encoding="utf-8")
            handler.setFormatter(logging.Formatter("%(message)s"))
            logger.addHandler(handler)
            logger.debug("==== NUEVO C??LCULO ====")

        rpm_eq = leer_flotante(entry_rpm_eq)
        hp_motor = leer_flotante(entry_hp)
        rpm_motor = leer_flotante(entry_rpm_motor)

        # Normalize user-entered units before any derived calculation.
        wr2 = leer_valor_unidad(entry_wr2, unidad_wr2.get())
        mnv = leer_valor_unidad(entry_mnv, unidad_mnv.get())

        if wr2 is None or mnv is None:
            if silencioso:
                return
            raise ValueError

        if not silencioso:
            actualizar_campo(entry_wr2, wr2, bloquear=False)
            marcar_completado(entry_wr2, label_check_wr2)
            unidad_wr2.set("kgm2")

            actualizar_campo(entry_mnv, mnv, bloquear=False)
            marcar_completado(entry_mnv, label_check_mnv)
            unidad_mnv.set("kgm")

        debug_print("wr2", "manual", wr2)
        debug_print("mnv", "manual", mnv)

        torque_pct, torque_factor = obtener_factor_torque()
        debug_print("torque_pct", "manual", torque_pct)
        debug_print("torque_factor", f"({torque_pct} / 100) ^ 2", torque_factor)

        if var_ventilador_bomba.get():
            factor_mmv = 0.33
            debug_print("factor_mmv", "bomba/ventilador", factor_mmv)
        else:
            factor_mmv = leer_flotante(entry_factor_mmv)
            if factor_mmv is None:
                if silencioso:
                    return
                raise ValueError
            debug_print("factor_mmv", "manual", factor_mmv)
            marcar_completado(entry_factor_mmv, label_check_factor_mmv)

        mmv = mnv * factor_mmv
        debug_print("mmv", f"{mnv} * {factor_mmv}", mmv)

        jv = wr2
        debug_print("jv", f"{wr2}", jv)

        actualizar_campo(entry_mmv, mmv)
        marcar_completado(entry_mmv, label_check_mmv)
        actualizar_campo(entry_jv, jv)
        marcar_completado(entry_jv, None)

        mnm = jm = gd2m = mmm = None
        if hp_motor and rpm_motor:
            # Nominal torque at motor shaft from power/rpm relation.
            mnm_base = (hp_motor * 726.15) / rpm_motor
            debug_print("mnm_base", f"({hp_motor} * 726.15) / {rpm_motor}", mnm_base)
            mnm = round(mnm_base * torque_factor, 2)
            debug_print("mnm", f"{mnm_base} * {torque_factor}", mnm)

            jm = leer_flotante(entry_jm)
            if jm is None:
                if silencioso:
                    return
                raise ValueError
            debug_print("jm", "manual", jm)

            gd2m = 4 * jm
            debug_print("gd2m", f"4 * {jm}", gd2m)

        if mnm is not None and "factor_mmm" in globals():
            mmm = mnm * factor_mmm
            debug_print("mmm", f"{mnm} * {factor_mmm}", mmm)

            actualizar_campo(entry_mnm, mnm, bloquear=True)
            marcar_completado(entry_mnm, label_check_mnm)
            marcar_completado(entry_jm, label_check_jm)
            actualizar_campo(entry_gd2m, gd2m)
            marcar_completado(entry_gd2m, label_check_gd2m)
            actualizar_campo(entry_mmm, mmm)
            marcar_completado(entry_mmm, label_check_mmm)

        mmv_nmot = gd2v = gd2v_nmot = ngd2 = None
        if rpm_eq and rpm_motor:
            # MMV nMOT redondeado solo a 2 decimales
            mmv_nmot_raw = mmv * (rpm_eq / rpm_motor)
            mmv_nmot = round(mmv_nmot_raw, 2)
            debug_print("mmv_nmot", f"{mmv} * ({rpm_eq} / {rpm_motor})", mmv_nmot)

            # GD2V sin redondeo
            gd2v = jv * 4
            debug_print("gd2v", f"{jv} * 4", gd2v)

            # GD2V nMOT redondeado solo a 2 decimales
            gd2v_nmot_raw = gd2v * (rpm_eq / rpm_motor) ** 2
            gd2v_nmot = round(gd2v_nmot_raw, 2)
            debug_print("gd2v_nmot", f"{gd2v} * ({rpm_eq} / {rpm_motor})??", gd2v_nmot)

            # nGD2 sin redondeo
            ngd2 = gd2v_nmot + gd2m if gd2m is not None else None
            debug_print("ngd2", f"{gd2v_nmot} + {gd2m}", ngd2)

            # Mostrar campos
            actualizar_campo(entry_mmv_nmot, mmv_nmot)
            marcar_completado(entry_mmv_nmot, label_check_mmv_nmot)
            actualizar_campo(entry_gd2v, gd2v)
            marcar_completado(entry_gd2v, label_check_gd2v)
            actualizar_campo(entry_gd2v_nmot, gd2v_nmot)
            marcar_completado(entry_gd2v_nmot, label_check_gd2v_nmot)
            if ngd2 is not None:
                actualizar_campo(entry_ngd2, ngd2)
                marcar_completado(entry_ngd2, label_check_ngd2)

            if rpm_motor and mmm and mmv_nmot and ngd2:
                # Primary start-up model used by the tool:
                # ta = (nGD2 * rpm_motor) / (375 * (MMM - MMV_nMot))
                ta_raw = (ngd2 * rpm_motor) / (375 * (mmm - mmv_nmot))
                ta = round(ta_raw, 2)
                ta_ceiling = math.ceil(ta_raw)

                debug_print("ta", f"({ngd2} * {rpm_motor}) / (375 * ({mmm} - {mmv_nmot}))", ta)
                debug_print("ta_ceiling", f"CEILING({ta_raw})", ta_ceiling)

                actualizar_campo(entry_ta, ta)
                marcar_completado(entry_ta, label_check_ta)
                actualizar_campo(entry_ta_ceiling, ta_ceiling)
                marcar_completado(entry_ta_ceiling, label_check_ta_ceiling)

        # Leer valores desde los campos para evaluacion segura
        ta_val = safe_float(entry_ta.get().strip()) if entry_ta.get().strip() else None
        stall_hot = safe_float(entry_stall_hot.get().strip()) if entry_stall_hot.get().strip() else None
        stall_cold = safe_float(entry_stall_cold.get().strip()) if entry_stall_cold.get().strip() else None
        mmm_val = safe_float(entry_mmm.get().strip()) if entry_mmm.get().strip() else None
        mmv_nmot_val = safe_float(entry_mmv_nmot.get().strip()) if entry_mmv_nmot.get().strip() else None

        mensaje = ""
        estado_texto = ""
        estado_color = "#666666"
        fmt = lambda v: formato_valor_pdf(v) if v is not None else "N/A"

        # Si el par medio del motor es menor o igual que el par medio de la carga,
        # el arranque no es posible (resultado fisicamente invalido aunque ta sea negativo).
        if mmm_val is not None and mmv_nmot_val is not None and mmm_val <= mmv_nmot_val:
            mensaje = (
                "Diagnostico/Diagnosis: NO SOPORTA ARRANQUE / DOES NOT SUPPORT START-UP.\n"
                f"Motivo/Reason: MMM ({fmt(mmm_val)}) <= MMV nMot ({fmt(mmv_nmot_val)}).\n"
                "Explicacion/Explanation: el par del motor no supera al de la carga, por lo que no acelera / "
                "motor torque does not exceed load torque, so it cannot accelerate."
            )
            estado_texto = mensaje
            estado_color = "#c00000"
        # Si no hay datos suficientes, no mostrar mensaje automatico
        elif ta_val is None or stall_hot is None or stall_cold is None:
            mensaje = ""
            estado_texto = (
                "Diagnostico/Diagnosis: DATOS INSUFICIENTES / INSUFFICIENT DATA.\n"
                "Explicacion/Explanation: complete ta, Stall HOT y Stall COLD para evaluar correctamente / "
                "complete ta, Stall HOT and Stall COLD for a valid evaluation."
            )
            estado_color = "#666666"
        else:
            # ta valido: si es 0 o negativo tratar como fallo
            if ta_val <= 0:
                mensaje = (
                    "Diagnostico/Diagnosis: NO SOPORTA ARRANQUE / DOES NOT SUPPORT START-UP.\n"
                    f"Motivo/Reason: ta ({fmt(ta_val)}) <= 0.\n"
                    "Explicacion/Explanation: valor no fisico o datos inconsistentes; verifique entradas / "
                    "non-physical value or inconsistent data; review inputs."
                )
                estado_texto = mensaje
                estado_color = "#c00000"
            elif ta_val < stall_hot and ta_val < stall_cold:
                mensaje = (
                    "Diagnostico/Diagnosis: SOPORTA ARRANQUE EN FRIO Y CALIENTE / SUPPORTS COLD AND HOT START-UP.\n"
                    f"Motivo/Reason: ta ({fmt(ta_val)}) < Stall HOT ({fmt(stall_hot)}) y/and < Stall COLD ({fmt(stall_cold)}).\n"
                    "Explicacion/Explanation: hay margen termico en ambos escenarios / there is thermal margin in both scenarios."
                )
                estado_texto = mensaje
                estado_color = "#007a00"
            elif ta_val < stall_cold:
                mensaje = (
                    "Diagnostico/Diagnosis: SOPORTE PARCIAL (SOLO FRIO) / PARTIAL SUPPORT (COLD ONLY).\n"
                    f"Motivo/Reason: ta ({fmt(ta_val)}) < Stall COLD ({fmt(stall_cold)}) y/and >= Stall HOT ({fmt(stall_hot)}).\n"
                    "Explicacion/Explanation: arranca en frio, pero en caliente excede el limite / "
                    "it starts in cold condition, but exceeds the limit in hot condition."
                )
                estado_texto = mensaje
                estado_color = "#d48300"
            elif ta_val < stall_hot:
                mensaje = (
                    "Diagnostico/Diagnosis: SOPORTE PARCIAL (SOLO CALIENTE) / PARTIAL SUPPORT (HOT ONLY).\n"
                    f"Motivo/Reason: ta ({fmt(ta_val)}) < Stall HOT ({fmt(stall_hot)}) y/and >= Stall COLD ({fmt(stall_cold)}).\n"
                    "Explicacion/Explanation: arranca en caliente, pero en frio excede el limite / "
                    "it starts in hot condition, but exceeds the limit in cold condition."
                )
                estado_texto = mensaje
                estado_color = "#0074c2"
            else:
                mensaje = (
                    "Diagnostico/Diagnosis: NO SOPORTA ARRANQUE / DOES NOT SUPPORT START-UP.\n"
                    f"Motivo/Reason: ta ({fmt(ta_val)}) >= Stall HOT ({fmt(stall_hot)}) y/and >= Stall COLD ({fmt(stall_cold)}).\n"
                    "Explicacion/Explanation: excede el tiempo maximo permitido en ambas condiciones / "
                    "it exceeds the maximum allowed time in both conditions."
                )
                estado_texto = mensaje
                estado_color = "#c00000"

        if mensaje and not silencioso:
            messagebox.showinfo("Evaluacion de Arranque / Start-Up Evaluation", mensaje, parent=form_toplevel)

        if estado_texto:
            actualizar_estado_arranque(estado_texto, estado_color)

        calcular_arranque_consecutivo()

        if entry_ta.get().strip() and "btn_consecutivos" in globals():
            try:
                if not btn_consecutivos.winfo_ismapped():
                    btn_consecutivos.pack(
                        pady=6,
                        padx=20,
                        fill="x",
                        before=frame_boton if "frame_boton" in globals() else None,
                    )
            except Exception:
                pass

        if actualizar_botones:
            btn_debug.grid()
            resaltar_campo(None)

    except ValueError:
        if not silencioso:
            messagebox.showerror(
                "Error / Error",
                "Ingresa todos los datos requeridos correctamente. / Enter all required data correctly.",
                parent=form_toplevel,
            )
        return
    except Exception as e:
        if not silencioso:
            messagebox.showerror(
                "Error / Error", f"Ocurrio un error inesperado: {e} / An unexpected error occurred: {e}", parent=form_toplevel
            )
    finally:
        debug_enabled = prev_debug_enabled

def actualizar_estado_factor_mmv(marcar: bool = True) -> None:
    """Actualiza el campo del factor MMV segÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Âºn el tipo de carga."""
    global var_ventilador_bomba, entry_factor_mmv, label_check_factor_mmv
    if var_ventilador_bomba.get():
        entry_factor_mmv.delete(0, tk.END)
        entry_factor_mmv.insert(0, "0.33")
        entry_factor_mmv.config(state="readonly", style="ReadonlyDark.TEntry")
        if marcar:
            marcar_completado(entry_factor_mmv, label_check_factor_mmv)
        else:
            setattr(entry_factor_mmv, "_completed", False)
            label_check_factor_mmv.config(text="")
    else:
        entry_factor_mmv.config(state="normal", style="TEntry")
        if not marcar:
            setattr(entry_factor_mmv, "_completed", False)
            label_check_factor_mmv.config(text="")
    programar_autocalculo()


def activar_arranque_consecutivo() -> None:
    """Show and enable the optional consecutive start section."""

    if "frame_consecutivos" not in globals():
        return
    if not frame_consecutivos.winfo_ismapped():
        frame_consecutivos.pack(
            padx=20, pady=10, fill="x", before=frame_boton if "frame_boton" in globals() else None
        )
    var_consecutivos_activo.set(True)
    try:
        btn_consecutivos.config(state="disabled")
    except Exception:
        pass
    actualizar_estado_consecutivos()
    if form_toplevel is not None:
        form_toplevel.after(30, actualizar_estado_consecutivos)


def preguntar_ventilador():
    """Pregunta si es bomba o ventilador antes de ingresar RPM."""
    global pregunta_respondida, paso_actual_var
    if pregunta_respondida:
        return
    respuesta = messagebox.askyesno(
        "Verificacion / Verification", "Es bomba o ventilador? / Is it a pump or a fan?", parent=form_toplevel
    )
    var_ventilador_bomba.set(respuesta)
    actualizar_estado_factor_mmv()
    pregunta_respondida = True
    if respuesta:
        paso_actual_var.set(1)
        entry_rpm_eq.config(state="normal", style="TEntry")
        resaltar_campo(entry_rpm_eq)
    else:
        paso_actual_var.set(0)
        entry_factor_mmv.config(state="normal", style="TEntry")
        resaltar_campo(entry_factor_mmv)


def reiniciar_calculos():
    """Limpia todas las entradas y restablece el formulario."""
    global pregunta_respondida, btn_calcular, btn_debug
    pregunta_respondida = False
    paso_actual_var.set(0)
    resaltar_campo(None)

    # Restablecer variable de bomba/ventilador y factor MMV
    var_ventilador_bomba.set(True)
    actualizar_estado_factor_mmv(marcar=False)

    # Campos de entrada principales
    entradas_desactivar = [entry_rpm_eq, entry_wr2, entry_mnv]
    for e in entradas_desactivar:
        e.config(state="normal", style="TEntry")
        e.delete(0, tk.END)
        setattr(e, "_completed", False)
        e.config(state="disabled", style="ReadonlyDark.TEntry")

    entry_mlfb.config(state="normal")
    entry_mlfb.delete(0, tk.END)
    btn_buscar.config(state="normal")

    for entrada in [
        entry_mmv,
        entry_jv,
        entry_hp,
        entry_rpm_motor,
        entry_mnm,
        entry_jm,
        entry_gd2m,
        entry_mmm,
        entry_mmv_nmot,
        entry_gd2v,
        entry_gd2v_nmot,
        entry_ngd2,
        entry_ta,
        entry_ta_ceiling,
        entry_stall_hot,
        entry_stall_cold,
    ]:
        entrada.config(state="normal", style="TEntry")
        entrada.delete(0, tk.END)
        setattr(entrada, "_completed", False)

    # Restaurar estados readonly
    readonly_entries = [
        entry_mmv,
        entry_jv,
        entry_mnm,
        entry_gd2m,
        entry_mmm,
        entry_mmv_nmot,
        entry_gd2v,
        entry_gd2v_nmot,
        entry_ngd2,
        entry_ta,
        entry_ta_ceiling,
    ]
    for e in readonly_entries:
        e.config(state="readonly", style="ReadonlyDark.TEntry")
        setattr(e, "_completed", False)

    # Factor MMV por defecto
    entry_factor_mmv.config(state="normal")
    entry_factor_mmv.delete(0, tk.END)
    entry_factor_mmv.insert(0, "0.33")
    entry_factor_mmv.config(state="readonly", style="ReadonlyDark.TEntry")

    if "entry_torque_pct" in globals():
        entry_torque_pct.config(state="normal", style="TEntry")
    establecer_torque_pct(100, marcar=False)

    if "btn_consecutivos" in globals():
        try:
            btn_consecutivos.pack_forget()
            btn_consecutivos.config(state="normal")
        except Exception:
            pass
    if "frame_consecutivos" in globals():
        try:
            frame_consecutivos.pack_forget()
        except Exception:
            pass
    if "var_consecutivos_activo" in globals():
        var_consecutivos_activo.set(False)

    consecutivos_entries = []
    for nombre in [
        "entry_f",
        "entry_gfe",
        "entry_diam_rotor",
        "entry_long_paquete",
        "entry_dibujo",
        "entry_diamxlongitud",
        "entry_v2a",
        "entry_v2a_ceiling",
        "entry_taus",
        "entry_taus_ceiling",
        "entry_taus_ta",
        "entry_ranura_profunda",
        "entry_doble_cuadro",
        "entry_std",
        "entry_v2aus",
        "entry_v2aus_ceiling",
        "entry_num_arranques",
        "entry_arranques_hora",
    ]:
        if nombre in globals():
            consecutivos_entries.append(globals()[nombre])

    for entrada in consecutivos_entries:
        try:
            entrada.config(state="normal", style="TEntry")
            entrada.delete(0, tk.END)
            setattr(entrada, "_completed", False)
        except Exception:
            pass

    if "tree_arranques" in globals():
        tree_arranques.delete(*tree_arranques.get_children())
    if "tree_arranques_caliente" in globals():
        tree_arranques_caliente.delete(*tree_arranques_caliente.get_children())

    # Restaurar valores por defecto en consecutivos
    if "entry_num_arranques" in globals():
        entry_num_arranques.insert(0, "10")
    if "entry_arranques_hora" in globals():
        entry_arranques_hora.insert(0, "4")
    if "var_stator_tipo" in globals():
        var_stator_tipo.set("std")

    for nombre in ["entry_diamxlongitud", "entry_v2a", "entry_v2a_ceiling", "entry_taus", "entry_taus_ceiling", "entry_taus_ta", "entry_v2aus", "entry_v2aus_ceiling"]:
        if nombre in globals():
            try:
                globals()[nombre].config(state="readonly", style="ReadonlyDark.TEntry")
            except Exception:
                pass
    actualizar_selector_estator()

    # Limpiar marcas de verificaciÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n
    for lbl in [
        label_check_mlfb,
        label_check_rpm,
        label_check_wr2,
        label_check_mnv,
        label_check_mmv,
        label_check_hp,
        label_check_rpm_motor,
        label_check_stall_hot,
        label_check_stall_cold,
        label_check_factor_mmv,
        label_check_jm,
        label_check_gd2m,
        label_check_mnm,
        label_check_mmm,
        label_check_mmv_nmot,
        label_check_gd2v,
        label_check_gd2v_nmot,
        label_check_ngd2,
        label_check_ta,
        label_check_ta_ceiling,
        label_check_torque_pct,
    ]:
        lbl.config(text="")

    for lbl in [
        label_lock_hp,
        label_lock_rpm_motor,
        label_lock_jm,
        label_lock_stall_hot,
        label_lock_stall_cold,
    ]:
        try:
            lbl.config(text="")
        except Exception:
            pass
    for lbl in consecutive_check_labels.values():
        try:
            lbl.config(text="")
        except Exception:
            pass



    btn_debug.grid()
    resaltar_campo(entry_mlfb)
    actualizar_estado_arranque("")


def letra_a_indice(col):
    """Convert a spreadsheet column letter(s) to a zero-based index."""

    col = col.upper().strip()
    total = 0
    for c in col:
        if not c.isalpha():
            raise ValueError("Columna invalida")
        total = total * 26 + (ord(c) - ord("A") + 1)
    return total - 1


def buscar_datos_mlfb():
    """Load motor data from Excel using the MLFB code provided."""

    global factor_mmm
    if pd is None:
        messagebox.showerror(
            "Dependencia faltante / Missing dependency",
            "La busqueda por MLFB requiere la libreria 'pandas'.\n"
            "MLFB search requires the 'pandas' library.\n"
            "Instale/Install: pip install pandas openpyxl",
            parent=form_toplevel,
        )
        return
    codigo = entry_mlfb.get().strip().upper()
    if len(codigo) != 11:
        messagebox.showerror(
            "Codigo invalido / Invalid code",
            "El codigo MLFB debe tener 11 caracteres. / MLFB code must have 11 characters.",
            parent=form_toplevel,
        )
        return

    try:
        df = pd.read_excel(
            mlfb_config["archivo"],
            sheet_name="Core Level Data 60 Hz",
            header=2,
            usecols=None,
        )

        # Obtener ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Ândices de columnas
        idx_mlfb = letra_a_indice(mlfb_config["col_mlfb"])
        idx_hp = letra_a_indice(mlfb_config["col_hp"])
        idx_inercia = letra_a_indice(mlfb_config["col_inercia"])
        idx_curvas = letra_a_indice(mlfb_config["col_curvas"])
        idx_hot = letra_a_indice(mlfb_config["col_hot"])
        idx_cold = letra_a_indice(mlfb_config["col_cold"])


        # Validar que los ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Ândices no excedan el nÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Âºmero de columnas reales
        num_columnas = df.shape[1]
        for idx, nombre, col_letra in zip(
            [idx_mlfb, idx_hp, idx_inercia, idx_curvas],
            ["MLFB", "HP", "Inercia", "Curvas"],
            [
                mlfb_config["col_mlfb"],
                mlfb_config["col_hp"],
                mlfb_config["col_inercia"],
                mlfb_config["col_curvas"],
            ],
        ):

            if idx >= num_columnas:
                messagebox.showerror(
                    "Columna fuera de rango / Column out of range",
                    f"La columna '{nombre}' ({col_letra}) no existe en el archivo.\n"
                    f"Requested index: {idx}, available columns: {num_columnas - 1}",
                    parent=form_toplevel,
                )
                return

        # Buscar fila con MLFB
        fila = df[df.iloc[:, idx_mlfb].astype(str).str.strip() == codigo]
        if fila.empty:
            messagebox.showwarning(
                "No encontrado / Not found",
                "No se encontro el MLFB en el archivo. / MLFB was not found in the file.",
                parent=form_toplevel,
            )
            return

        # Extraer datos
        hp = fila.iloc[0, idx_hp]
        debug_print("hp", "valor desde Excel", hp)

        inercia = fila.iloc[0, idx_inercia]
        debug_print("inercia", "valor desde Excel", inercia)

        factor_mmm_excel = fila.iloc[0, idx_curvas]
        factor_mmm = float(factor_mmm_excel) / 100
        debug_print(
            "factor_mmm (desde Excel BT)", f"{factor_mmm_excel} / 100", factor_mmm
        )

        torque_pct, torque_factor = obtener_factor_torque()
        debug_print("torque_pct", "manual", torque_pct)
        debug_print("torque_factor", f"({torque_pct} / 100) ^ 2", torque_factor)

        stall_hot = fila.iloc[0, idx_hot]
        debug_print("stall_hot", "valor desde Excel", stall_hot)

        stall_cold = fila.iloc[0, idx_cold]
        debug_print("stall_cold", "valor desde Excel", stall_cold)

        entry_stall_hot.delete(0, tk.END)
        entry_stall_hot.insert(0, str(stall_hot))
        bloquear_si_lleno(entry_stall_hot, label_check_stall_hot, label_lock_stall_hot)

        entry_stall_cold.delete(0, tk.END)
        entry_stall_cold.insert(0, str(stall_cold))
        bloquear_si_lleno(entry_stall_cold, label_check_stall_cold, label_lock_stall_cold)

        # Inferir RPM a partir del MLFB
        letra_rpm = codigo[9].upper()
        rpm_map = {"A": 3600, "B": 1800, "C": 1200, "D": 900}
        rpm = rpm_map.get(letra_rpm, None)
        debug_print("rpm", f"letra {letra_rpm} ??? {rpm}", rpm)

        # Llenar campos
        entry_hp.delete(0, tk.END)
        entry_hp.insert(0, str(hp))
        bloquear_si_lleno(entry_hp, label_check_hp, label_lock_hp)

        entry_rpm_motor.delete(0, tk.END)
        entry_rpm_motor.insert(0, str(rpm if rpm else ""))
        bloquear_si_lleno(entry_rpm_motor, label_check_rpm_motor, label_lock_rpm_motor)

        jm = float(inercia) / 23.73
        entry_jm.config(state="normal")
        entry_jm.delete(0, tk.END)
        entry_jm.insert(0, f"{jm:.7f}")
        entry_jm.config(state="normal", style="TEntry")
        marcar_completado(entry_jm, label_check_jm)

        if rpm:
            mnm_base = (float(hp) * 726.15) / rpm
            debug_print("mnm_base", f"({hp} * 726.15) / {rpm}", mnm_base)
            mnm = round(mnm_base * torque_factor, 2)
            debug_print("mnm", f"{mnm_base} * {torque_factor}", mnm)
            gd2m = 4 * jm
            mmm = mnm * factor_mmm
            actualizar_campo(entry_mnm, mnm)
            marcar_completado(entry_mnm, label_check_mnm)
            actualizar_campo(entry_gd2m, gd2m)
            marcar_completado(entry_gd2m, label_check_gd2m)
            actualizar_campo(entry_mmm, mmm)
            marcar_completado(entry_mmm, label_check_mmm)

        messagebox.showinfo(
            "Autollenado / Auto-fill",
            "Datos del motor cargados correctamente. / Motor data loaded successfully.",
            parent=form_toplevel,
        )
        marcar_completado(entry_mlfb, label_check_mlfb)
        btn_buscar.config(state="disabled")
        preguntar_ventilador()
    except Exception as e:
        messagebox.showerror(
            "Error / Error",
            f"No se pudo procesar el archivo / Could not process file:\n{e}",
            parent=form_toplevel,
        )


def actualizar_campo(campo, valor, bloquear=True):
    """Actualiza un `tk.Entry` y lo deja siempre en modo de solo lectura."""

    campo.config(state="normal")
    campo.delete(0, tk.END)
    campo.insert(0, formato_valor_pdf(valor) if isinstance(valor, float) else str(valor))
    if bloquear:
        campo.config(state="readonly", style="ReadonlyDark.TEntry")
    else:
        campo.config(state="normal")


def abrir_formulario_arranque(root_principal, pos_x: int | None = None, pos_y: int | None = None):
    """Create the start-up calculation form as a child of *root_principal*.

    The window will appear on the same screen as ``root_principal`` so that it
    follows the user's preferred monitor.
    """
    global form_toplevel
    form = tk.Toplevel(root_principal)
    form_toplevel = form
    form.title("Calculo de Tiempo de Arranque")

    # Keep child window on the same monitor as launcher for smoother workflow.
    try:
        if pos_x is None or pos_y is None:
            root_principal.update_idletasks()
            pos_x = root_principal.winfo_rootx()
            pos_y = root_principal.winfo_rooty()
        form.geometry(f"1050x800+{pos_x}+{pos_y}")
    except Exception:
        form.geometry("1050x800")

    form.minsize(800, 600)
    form.resizable(True, True)

    aplicar_colorimetria(form)
    configurar_estilo_moderno(form)
    # Cargar imÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡genes y mostrarlas en la parte superior del formulario
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    motor_path = os.path.join(os.path.dirname(__file__), "motor_header.png")

    try:
        logo_img_raw = Image.open(logo_path)
        logo_img_raw.thumbnail((100, 100))
        form.logo_img = ImageTk.PhotoImage(logo_img_raw)
    except Exception as e:
        print(f"[LOGO] No se pudo cargar {logo_path}: {e}")
        form.logo_img = None

    try:
        motor_img_raw = Image.open(motor_path)
        motor_img_raw.thumbnail((100, 100))
        form.motor_img = ImageTk.PhotoImage(motor_img_raw)
    except Exception as e:
        print(f"[MOTOR] No se pudo cargar {motor_path}: {e}")
        form.motor_img = None
    header = ttk.Frame(form)
    header.pack(side="top", fill="x", pady=(10, 0))
    if Image is None or ImageTk is None:
        form.logo_img = None
        form.motor_img = None
        print("[PIL] Pillow not available; skipping header images.")
    else:
        if form.logo_img:
            ttk.Label(header, image=form.logo_img).pack(side="left", padx=10)
        if form.motor_img:
            ttk.Label(header, image=form.motor_img).pack(side="right", padx=10)
    global entry_mlfb, entry_excel_path, entry_col_mlfb, entry_col_hp, entry_col_inercia, var_ventilador_bomba, entry_factor_mmv
    global entry_rpm_eq, entry_wr2, entry_mnv, entry_mmv, entry_jv
    global entry_hp, entry_rpm_motor, entry_mnm, entry_jm, entry_gd2m, entry_mmm
    global entry_torque_pct, var_torque_pct, var_torque_pct_str, label_check_torque_pct
    global entry_mmv_nmot, entry_gd2v, entry_gd2v_nmot, entry_ngd2
    global entry_col_curvas
    global entry_stall_hot, entry_stall_cold
    global entry_col_hot, entry_col_cold
    global label_check_mlfb, label_check_rpm, label_check_wr2, label_check_mnv, label_check_mmv
    global label_check_hp, label_check_rpm_motor, label_check_stall_hot, label_check_stall_cold, label_check_factor_mmv
    global label_lock_hp, label_lock_rpm_motor, label_lock_jm, label_lock_stall_hot, label_lock_stall_cold
    global label_check_jm, label_check_gd2m, label_check_mnm, label_check_mmm, label_check_torque_pct
    global label_check_mmv_nmot, label_check_gd2v, label_check_gd2v_nmot
    global label_check_ngd2, label_check_ta, label_check_ta_ceiling
    global paso_actual_var, btn_buscar, btn_calcular, btn_debug
    global var_eval_text, label_eval
    global var_consecutivos_activo, btn_consecutivos, frame_consecutivos, frame_boton
    global var_stator_tipo
    global entry_f, entry_gfe, entry_diam_rotor, entry_long_paquete, entry_dibujo
    global entry_diamxlongitud, entry_v2a, entry_v2a_ceiling
    global entry_taus, entry_taus_ceiling, entry_taus_ta
    global entry_ranura_profunda, entry_doble_cuadro, entry_std, entry_v2aus, entry_v2aus_ceiling
    global entry_num_arranques, canvas_arranques, canvas_arranques_caliente
    global tree_arranques, tree_arranques_caliente
    global entry_arranques_hora, tree_arranques_hora, canvas_arranques_hora
    global entry_ciclo_arr, entry_resto_arr, entry_reposo_ciclo



    
    var_ventilador_bomba = tk.BooleanVar(form, value=True)  # ??? InicializaciÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n correcta
    paso_actual_var = tk.IntVar(value=0)
    var_torque_pct = tk.DoubleVar(form, value=100.0)
    var_torque_pct_str = tk.StringVar(form, value="100")
    var_torque_pct_str.trace_add("write", sync_torque_from_str)
    var_consecutivos_activo = tk.BooleanVar(form, value=False)
    var_stator_tipo = tk.StringVar(form, value="std")
    global unidad_wr2, unidad_mnv
    unidad_wr2 = tk.StringVar(value="kgm2")
    unidad_mnv = tk.StringVar(value="kgm")
    unidad_wr2.trace_add("write", lambda *_: programar_autocalculo())
    unidad_mnv.trace_add("write", lambda *_: programar_autocalculo())

    def on_rpm_entered(event=None):
        """When RPM is entered, enable the WR2 step."""

        if paso_actual_var.get() == 1 and entry_rpm_eq.get().strip():
            entry_wr2.config(state="normal", style="TEntry")
            marcar_completado(entry_rpm_eq, label_check_rpm)
            paso_actual_var.set(2)
            resaltar_campo(entry_wr2)
        programar_autocalculo()

    def on_wr2_entered(event=None):
        """Activate the MNV field once WR2 is provided."""

        if paso_actual_var.get() == 2 and entry_wr2.get().strip():
            entry_mnv.config(state="normal", style="TEntry")
            marcar_completado(entry_wr2, label_check_wr2)
            paso_actual_var.set(3)
            resaltar_campo(entry_mnv)
        programar_autocalculo()

    def on_mnv_entered(event=None):
        """Mark the MNV field complete and highlight the calculate button."""

        if paso_actual_var.get() == 3 and entry_mnv.get().strip():
            marcar_completado(entry_mnv, label_check_mnv)
            paso_actual_var.set(4)
            resaltar_campo(None)
        programar_autocalculo()

    def on_hp_entered(event=None):
        """Validate the HP entry when the user finishes editing."""

        if entry_hp.get().strip():
            marcar_completado(entry_hp, label_check_hp)
        programar_autocalculo()

    def on_rpm_motor_entered(event=None):
        """Update the check mark for motor RPM."""

        if entry_rpm_motor.get().strip():
            marcar_completado(entry_rpm_motor, label_check_rpm_motor)
        programar_autocalculo()

    def on_jm_entered(event=None):
        """Mark JM as entered and recalc."""

        if entry_jm.get().strip():
            marcar_completado(entry_jm, label_check_jm)
        programar_autocalculo()

    def on_stall_hot_entered(event=None):
        """Mark HOT stall current as entered."""

        if entry_stall_hot.get().strip():
            marcar_completado(entry_stall_hot, label_check_stall_hot)
        programar_autocalculo()

    def on_stall_cold_entered(event=None):
        """Mark COLD stall current as entered."""

        if entry_stall_cold.get().strip():
            marcar_completado(entry_stall_cold, label_check_stall_cold)
        programar_autocalculo()

    def on_factor_mmv_entered(event=None):
        """Validate the MMV factor field and move to the next step."""

        if not entry_factor_mmv.get().strip():
            return
        marcar_completado(entry_factor_mmv, label_check_factor_mmv)
        if paso_actual_var.get() == 0:
            paso_actual_var.set(1)
            entry_rpm_eq.config(state="normal", style="TEntry")
            resaltar_campo(entry_rpm_eq)
        programar_autocalculo()

    def on_torque_pct_entered(event=None):
        """Sync torque percentage from the entry."""

        texto = entry_torque_pct.get().strip()
        if not texto:
            return
        valor = safe_float(texto)
        if valor is None:
            return
        establecer_torque_pct(valor, marcar=True)
        programar_autocalculo()

    def on_torque_pct_scale(value=None):
        """Sync torque percentage from the scale."""

        try:
            valor = float(value) if value is not None else float(var_torque_pct.get())
        except (TypeError, ValueError):
            return
        establecer_torque_pct(valor, marcar=True)
        programar_autocalculo()

    def al_cerrar():
        """Close the form window and show the main menu again."""

        form.destroy()
        root_principal.deiconify()

    form.protocol("WM_DELETE_WINDOW", al_cerrar)

    contenedor = ttk.Frame(form)
    contenedor.pack(fill="both", expand=True)

    canvas = tk.Canvas(
        contenedor,
        background=UI_THEME["bg"],
        highlightthickness=0,
        bd=0,
    )
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")

    canvas.configure(yscrollcommand=scrollbar.set)

    frame_scroll = ttk.Frame(canvas)
    frame_scroll_id = canvas.create_window((0, 0), window=frame_scroll, anchor="nw")

    def on_canvas_configure(event):
        # Sync embedded frame width and scrollable region when resizing.
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(frame_scroll_id, width=event.width)

    canvas.bind("<Configure>", on_canvas_configure)
    frame_scroll.bind(
        "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    hero = ttk.Frame(frame_scroll, style="Hero.TFrame", padding=(18, 14))
    hero.pack(padx=20, pady=(12, 8), fill="x")
    ttk.Label(
        hero,
        text="Calculo de Tiempo de Arranque / Start-Up Time Calculation",
        style="HeroTitle.TLabel",
        anchor="w",
    ).pack(fill="x")
    ttk.Label(
        hero,
        text="Interfaz optimizada para evaluacion rapida, arranques consecutivos y exportacion tecnica. / "
        "Optimized interface for fast evaluation, consecutive starts, and technical export.",
        style="HeroSubtitle.TLabel",
        anchor="w",
    ).pack(fill="x", pady=(2, 0))

    var_eval_text = tk.StringVar(value="")
    frame_estado = ttk.Labelframe(
        frame_scroll, text="Evaluacion de Arranque / Start-Up Evaluation", padding=10, style="Card.TLabelframe"
    )
    frame_estado.pack(padx=20, pady=(6, 5), fill="x")
    frame_estado.columnconfigure(0, weight=1)
    label_eval = ttk.Label(
        frame_estado,
        textvariable=var_eval_text,
        justify="center",
        anchor="center",
        wraplength=900,
        style="Status.TLabel",
    )
    label_eval.grid(row=0, column=0, sticky="ew", padx=8, pady=6)

    frame_mlfb = ttk.Labelframe(
        frame_scroll, text="Buscar por Codigo MLFB / Search by MLFB Code", padding=12, style="Card.TLabelframe"
    )
    frame_mlfb.pack(padx=20, pady=8, fill="x")
    frame_mlfb.columnconfigure(1, weight=1)
    frame_mlfb.columnconfigure(4, weight=0)

    ttk.Label(frame_mlfb, text="Codigo MLFB (11 digitos) / MLFB Code (11 digits):").grid(
        row=0, column=0, padx=10, pady=5, sticky="w"
    )
    entry_mlfb = ttk.Entry(frame_mlfb)
    entry_mlfb.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
    resaltar_campo(entry_mlfb)

    btn_buscar = ttk.Button(
        frame_mlfb, text="Buscar / Search", command=buscar_datos_mlfb, style="Accent.TButton"
    )
    btn_buscar.grid(row=0, column=2, padx=10)
    label_check_mlfb = ttk.Label(frame_mlfb, text="")
    label_check_mlfb.grid(row=0, column=3, padx=5)


    frame_config_excel = ttk.Labelframe(
        frame_scroll, text="Configuracion de Archivo / File Configuration", padding=12, style="Card.TLabelframe"
    )
    frame_config_excel.pack(padx=20, pady=8, fill="x")
    frame_config_excel.columnconfigure(1, weight=1)

    ttk.Label(frame_config_excel, text="Archivo Excel / Excel File:").grid(
        row=0, column=0, padx=10, pady=5, sticky="w"
    )
    entry_excel_path = ttk.Entry(frame_config_excel, state="readonly", style="ReadonlyDark.TEntry")
    entry_excel_path.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
    entry_excel_path.insert(0, mlfb_config["archivo"])
    ttk.Button(
        frame_config_excel,
        text="Cambiar... / Change...",
        command=seleccionar_archivo,
        style="Accent.TButton",
    ).grid(
        row=0, column=2, padx=10
    )

    ttk.Label(frame_config_excel, text="Columna MLFB / MLFB Column:").grid(
        row=1, column=0, padx=10, pady=5
    )
    ttk.Label(frame_config_excel, text="Columna HP / HP Column:").grid(
        row=1, column=1, padx=10, pady=5
    )
    ttk.Label(frame_config_excel, text="Columna Inercia / Inertia Column:").grid(
        row=1, column=2, padx=10, pady=5
    )

    entry_col_mlfb = ttk.Entry(frame_config_excel)
    entry_col_mlfb.insert(0, mlfb_config["col_mlfb"])
    entry_col_mlfb.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

    entry_col_hp = ttk.Entry(frame_config_excel)
    entry_col_hp.insert(0, mlfb_config["col_hp"])
    entry_col_hp.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

    entry_col_inercia = ttk.Entry(frame_config_excel)
    entry_col_inercia.insert(0, mlfb_config["col_inercia"])
    entry_col_inercia.grid(row=2, column=2, padx=10, pady=5, sticky="ew")


    paned = ttk.Panedwindow(frame_scroll, orient=tk.HORIZONTAL)
    paned.pack(padx=20, pady=10, expand=True, fill="both")

    frame_eq = ttk.Labelframe(
        paned, text="Datos del Equipo Accionado / Driven Equipment Data", padding=12, style="Card.TLabelframe"
    )
    frame_motor = ttk.Labelframe(
        paned, text="Datos del Motor / Motor Data", padding=12, style="Card.TLabelframe"
    )
    for f in (frame_eq, frame_motor):
        f.columnconfigure(1, weight=1)
        f.columnconfigure(4, weight=0)
    paned.add(frame_eq, weight=1)
    paned.add(frame_motor, weight=1)


    crear_input(
        frame_eq,
        "RPM (Equipo) / RPM (Equipment):",
        0,
        help_text="Revoluciones por Minuto / Revolutions per Minute",
    )
    entry_rpm_eq = crear_entry(
        frame_eq,
        0,
        state="disabled",
        style="ReadonlyDark.TEntry",
    )
    entry_rpm_eq.bind("<KeyRelease>", on_rpm_entered)
    entry_rpm_eq.bind("<FocusIn>", lambda e: preguntar_ventilador())
    label_check_rpm = ttk.Label(frame_eq, text="")
    label_check_rpm.grid(row=0, column=3, padx=5)

    crear_input(
        frame_eq,
        "WR2:",
        1,
        help_text="Momento de Inercia / Inertia Moment",
        unit_text="kgm2",
    )
    entry_wr2 = crear_entry(
        frame_eq,
        1,
        state="disabled",
        style="ReadonlyDark.TEntry",
    )
    entry_wr2.bind("<KeyRelease>", on_wr2_entered)
    label_check_wr2 = ttk.Label(frame_eq, text="")
    label_check_wr2.grid(row=1, column=3, padx=5)
    ttk.Combobox(frame_eq, textvariable=unidad_wr2, values=["kgm2", "lbft2"], state="readonly").grid(row=1, column=2, padx=5, sticky="ew")

    crear_input(
        frame_eq,
        "MNV:",
        3,
        help_text="Momento Nominal del Ventilador (Mnv) / Fan Nominal Torque (Mnv).",
        unit_text="kgm",
    )
    entry_mnv = crear_entry(
        frame_eq,
        3,
        state="disabled",
        style="ReadonlyDark.TEntry",
    )
    entry_mnv.bind("<KeyRelease>", on_mnv_entered)
    label_check_mnv = ttk.Label(frame_eq, text="")
    label_check_mnv.grid(row=3, column=3, padx=5)
    ttk.Combobox(frame_eq, textvariable=unidad_mnv, values=["kgm", "lbft"], state="readonly").grid(row=3, column=2, padx=5, sticky="ew")


    crear_input(
        frame_eq,
        "MMV:",
        5,
        help_text="Momento Medio del Ventilador (segun curva) / Average Fan Torque (according to curve)",
        unit_text="kgm",
    )
    entry_mmv = crear_entry(frame_eq, 5, readonly=True)
    label_check_mmv = ttk.Label(frame_eq, text="")
    label_check_mmv.grid(row=5, column=3, padx=5)
    crear_input(frame_eq, "JV:", 6, unit_text="kgm2")
    entry_jv = crear_entry(frame_eq, 6, readonly=True)


    crear_input(
        frame_motor,
        "HP (Motor):",
        0,
        help_text="Horse Power del motor (HP) / Motor Horse Power (HP).",
    )
    entry_hp = crear_entry(
        frame_motor,
        0,
    )
    entry_hp.bind("<KeyRelease>", on_hp_entered)
    entry_hp.bind(
        "<FocusOut>", lambda e: bloquear_si_lleno(entry_hp, label_check_hp, label_lock_hp)
    )
    entry_hp.bind(
        "<Return>", lambda e: bloquear_si_lleno(entry_hp, label_check_hp, label_lock_hp)
    )
    label_check_hp = ttk.Label(frame_motor, text="")
    label_check_hp.grid(row=0, column=3, padx=5)
    label_lock_hp = ttk.Label(frame_motor, text="")
    label_lock_hp.grid(row=0, column=5, padx=5)
    crear_input(frame_motor, "RPM (Motor):", 1)
    entry_rpm_motor = crear_entry(
        frame_motor,
        1,
    )
    entry_rpm_motor.bind("<KeyRelease>", on_rpm_motor_entered)
    entry_rpm_motor.bind(
        "<FocusOut>",
        lambda e: bloquear_si_lleno(
            entry_rpm_motor, label_check_rpm_motor, label_lock_rpm_motor
        ),
    )
    entry_rpm_motor.bind(
        "<Return>",
        lambda e: bloquear_si_lleno(
            entry_rpm_motor, label_check_rpm_motor, label_lock_rpm_motor
        ),
    )
    label_check_rpm_motor = ttk.Label(frame_motor, text="")
    label_check_rpm_motor.grid(row=1, column=3, padx=5)
    label_lock_rpm_motor = ttk.Label(frame_motor, text="")
    label_lock_rpm_motor.grid(row=1, column=5, padx=5)
    crear_input(frame_motor, "MNM:", 2, unit_text="kgm")
    entry_mnm = crear_entry(frame_motor, 2, readonly=True)
    label_check_mnm = ttk.Label(frame_motor, text="")
    label_check_mnm.grid(row=2, column=3, padx=5)
    crear_input(
        frame_motor,
        "JM:",
        3,
        help_text="Momento de inercia del rotor del motor (JM) / Motor rotor inertia moment (JM).",
        unit_text="kgm2",
    )
    entry_jm = crear_entry(frame_motor, 3)
    entry_jm.bind("<KeyRelease>", on_jm_entered)
    entry_jm.bind(
        "<FocusOut>", lambda e: bloquear_si_lleno(entry_jm, label_check_jm, label_lock_jm)
    )
    entry_jm.bind(
        "<Return>", lambda e: bloquear_si_lleno(entry_jm, label_check_jm, label_lock_jm)
    )
    label_check_jm = ttk.Label(frame_motor, text="")
    label_check_jm.grid(row=3, column=3, padx=5)
    label_lock_jm = ttk.Label(frame_motor, text="")
    label_lock_jm.grid(row=3, column=5, padx=5)
    crear_input(
        frame_motor,
        "GD2M:",
        4,
        help_text="Momento cuarto de inercia del motor (GD2M) / Motor fourth inertia moment (GD2M).",
        unit_text="kgm2",
    )
    entry_gd2m = crear_entry(frame_motor, 4, readonly=True)
    label_check_gd2m = ttk.Label(frame_motor, text="")
    label_check_gd2m.grid(row=4, column=3, padx=5)
    crear_input(frame_motor, "MMM:", 5, unit_text="kgm")
    entry_mmm = crear_entry(frame_motor, 5, readonly=True)
    label_check_mmm = ttk.Label(frame_motor, text="")
    label_check_mmm.grid(row=5, column=3, padx=5)
    crear_input(frame_motor, "Stall Time HOT:", 6, unit_text="sec")
    entry_stall_hot = crear_entry(
        frame_motor,
        6,
    )
    entry_stall_hot.bind("<KeyRelease>", on_stall_hot_entered)
    entry_stall_hot.bind(
        "<FocusOut>",
        lambda e: bloquear_si_lleno(
            entry_stall_hot, label_check_stall_hot, label_lock_stall_hot
        ),
    )
    entry_stall_hot.bind(
        "<Return>",
        lambda e: bloquear_si_lleno(
            entry_stall_hot, label_check_stall_hot, label_lock_stall_hot
        ),
    )
    label_check_stall_hot = ttk.Label(frame_motor, text="")
    label_check_stall_hot.grid(row=6, column=3, padx=5)
    label_lock_stall_hot = ttk.Label(frame_motor, text="")
    label_lock_stall_hot.grid(row=6, column=5, padx=5)

    crear_input(frame_motor, "Stall Time COLD:", 7, unit_text="sec")
    entry_stall_cold = crear_entry(
        frame_motor,
        7,
    )
    entry_stall_cold.bind("<KeyRelease>", on_stall_cold_entered)
    entry_stall_cold.bind(
        "<FocusOut>",
        lambda e: bloquear_si_lleno(
            entry_stall_cold, label_check_stall_cold, label_lock_stall_cold
        ),
    )
    entry_stall_cold.bind(
        "<Return>",
        lambda e: bloquear_si_lleno(
            entry_stall_cold, label_check_stall_cold, label_lock_stall_cold
        ),
    )
    label_check_stall_cold = ttk.Label(frame_motor, text="")
    label_check_stall_cold.grid(row=7, column=3, padx=5)
    label_lock_stall_cold = ttk.Label(frame_motor, text="")
    label_lock_stall_cold.grid(row=7, column=5, padx=5)

    ttk.Label(frame_motor, text="Porcentaje de voltaje alimentado (%) / Applied voltage percentage (%):").grid(
        row=8, column=0, padx=10, pady=5, sticky="w"
    )
    torque_frame = ttk.Frame(frame_motor)
    torque_frame.grid(row=8, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
    torque_frame.columnconfigure(0, weight=1)
    scale_torque_pct = ttk.Scale(
        torque_frame,
        from_=0,
        to=100,
        orient="horizontal",
        variable=var_torque_pct,
        command=on_torque_pct_scale,
    )
    scale_torque_pct.grid(row=0, column=0, sticky="ew")
    entry_torque_pct = ttk.Entry(
        torque_frame,
        width=5,
        textvariable=var_torque_pct_str,
        justify="center",
    )
    entry_torque_pct.grid(row=0, column=1, padx=(8, 0))
    entry_torque_pct.bind("<KeyRelease>", on_torque_pct_entered)
    entry_torque_pct.bind("<FocusOut>", on_torque_pct_entered)
    entry_torque_pct.bind("<Return>", on_torque_pct_entered)
    ttk.Label(torque_frame, text="%").grid(row=0, column=2, padx=(4, 0))
    label_check_torque_pct = ttk.Label(frame_motor, text="")
    label_check_torque_pct.grid(row=8, column=3, padx=5)
    ayuda_torque = ttk.Label(frame_motor, text="?", style="Muted.TLabel", cursor="question_arrow")
    ayuda_torque.grid(row=8, column=5, padx=5)
    ToolTip(
        ayuda_torque,
        "Porcentaje de voltaje alimentado (0-100). Se aplica como (p/100)^2 x MNM.",
    )

    ttk.Checkbutton(
        frame_eq,
        text="Es bomba o ventilador? / Is it a pump or a fan?",
        variable=var_ventilador_bomba,
        command=actualizar_estado_factor_mmv,
    ).grid(row=10, column=0, columnspan=2, padx=10, pady=5)
    # Campo para ingresar el factor MMV manualmente (cuando NO es bomba o ventilador)
    crear_input(
        frame_eq,
        "Factor MMV (%) / MMV Factor (%)",
        11,
        help_text="Momento Medio del Ventilador (segun curva) / Average Fan Torque (according to curve)",
    )
    global entry_factor_mmv
    entry_factor_mmv = ttk.Entry(frame_eq)
    entry_factor_mmv.grid(row=11, column=1, padx=10, pady=5, sticky="ew")
    entry_factor_mmv.bind("<KeyRelease>", on_factor_mmv_entered)
    label_check_factor_mmv = ttk.Label(frame_eq, text="")
    label_check_factor_mmv.grid(row=11, column=3, padx=5)

    entry_factor_mmv.insert(0, "0.33")
    entry_factor_mmv.config(state="readonly", style="ReadonlyDark.TEntry")

    ttk.Label(frame_config_excel, text="Columna Curvas / Curves Column:").grid(
        row=1, column=3, padx=10, pady=5
    )
    global entry_col_curvas
    entry_col_curvas = ttk.Entry(frame_config_excel)
    entry_col_curvas.insert(0, mlfb_config["col_curvas"])
    entry_col_curvas.grid(row=2, column=3, padx=10, pady=5, sticky="ew")


    ttk.Label(frame_config_excel, text="Columna Stall HOT / Stall HOT Column:").grid(row=1, column=4, padx=10, pady=5)
    entry_col_hot = ttk.Entry(frame_config_excel)
    entry_col_hot.insert(0, mlfb_config["col_hot"])
    entry_col_hot.grid(row=2, column=4, padx=10, pady=5, sticky="ew")

    ttk.Label(frame_config_excel, text="Columna Stall COLD / Stall COLD Column:").grid(row=1, column=5, padx=10, pady=5)
    entry_col_cold = ttk.Entry(frame_config_excel)
    entry_col_cold.insert(0, mlfb_config["col_cold"])
    entry_col_cold.grid(row=2, column=5, padx=10, pady=5, sticky="ew")

    # Mueve el botÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n a la columna 6
    ttk.Button(
        frame_config_excel,
        text="Guardar columnas / Save columns",
        command=guardar_columnas,
        style="Accent.TButton",
    ).grid(row=2, column=6, padx=10)

    construir_adaptaciones(frame_scroll) 

    btn_consecutivos = ttk.Button(
        frame_scroll,
        text="Arranques consecutivos / Consecutive starts",
        command=activar_arranque_consecutivo,
        style="Accent.TButton",
    )

    frame_consecutivos = ttk.Labelframe(
        frame_scroll, text="Arranques consecutivos / Consecutive starts", padding=12, style="Card.TLabelframe"
    )
    frame_consecutivos.columnconfigure(1, weight=1)
    frame_consecutivos.columnconfigure(4, weight=0)

    def on_consecutivo_edit(event=None):
        actualizar_estado_consecutivos()
        programar_autocalculo()

    consecutive_check_labels.clear()

    crear_input(
        frame_consecutivos,
        "f=",
        0,
        help_text="Factor de reduccion f para el calculo del calentamiento en arranque / Reduction factor f for start-up heating calculation.",
    )
    entry_f = crear_entry(frame_consecutivos, 0)
    _crear_check_consecutivo(frame_consecutivos, 0, "entry_f")
    entry_f.bind("<KeyRelease>", on_consecutivo_edit)

    crear_input(
        frame_consecutivos,
        "GFe=",
        1,
        help_text="(diam. rotor en dm - 0.1) x p x 0.1 x (peso especifico Fe=7.8) "
        "x (long. del paquete) x (factor apilamiento=0.945) = / "
        "(rotor diam. in dm - 0.1) x p x 0.1 x (Fe specific weight=7.8) x (stack length) x (stack factor=0.945) =",
    )
    entry_gfe = crear_entry(frame_consecutivos, 1)
    _crear_check_consecutivo(frame_consecutivos, 1, "entry_gfe")
    entry_gfe.bind("<KeyRelease>", on_consecutivo_edit)

    crear_input(frame_consecutivos, "Diam. Rotor en dm / Rotor Diam. in dm:", 2)
    entry_diam_rotor = crear_entry(frame_consecutivos, 2)
    _crear_check_consecutivo(frame_consecutivos, 2, "entry_diam_rotor")
    entry_diam_rotor.bind("<KeyRelease>", on_consecutivo_edit)

    crear_input(frame_consecutivos, "Long. del paquete dm / Stack length in dm:", 3)
    entry_long_paquete = crear_entry(frame_consecutivos, 3)
    _crear_check_consecutivo(frame_consecutivos, 3, "entry_long_paquete")
    entry_long_paquete.bind("<KeyRelease>", on_consecutivo_edit)

    crear_input(frame_consecutivos, "Dibujo paquete rotor / Rotor stack drawing:", 4)
    entry_dibujo = crear_entry(frame_consecutivos, 4)
    _crear_check_consecutivo(frame_consecutivos, 4, "entry_dibujo")
    entry_dibujo.bind("<KeyRelease>", on_consecutivo_edit)

    crear_input(frame_consecutivos, "Diamxlongitud:", 5, unit_text="dm2")
    entry_diamxlongitud = crear_entry(frame_consecutivos, 5, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 5, "entry_diamxlongitud")

    crear_input(
        frame_consecutivos,
        "V'2A:",
        6,
        help_text="Incremento de temperatura del rotor (V'2A) / Rotor temperature rise (V'2A).",
        unit_text="°C",
    )
    entry_v2a = crear_entry(frame_consecutivos, 6, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 6, "entry_v2a")

    crear_input(frame_consecutivos, "V'2A Ceiling:", 7, unit_text="°C")
    entry_v2a_ceiling = crear_entry(frame_consecutivos, 7, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 7, "entry_v2a_ceiling")

    crear_input(
        frame_consecutivos,
        "taus:",
        8,
        help_text="Tiempo de enfriamiento (taus) / Cooling time (taus).",
        unit_text="seg",
    )
    entry_taus = crear_entry(frame_consecutivos, 8, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 8, "entry_taus")

    crear_input(frame_consecutivos, "taus ceiling:", 9, unit_text="seg")
    entry_taus_ceiling = crear_entry(frame_consecutivos, 9, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 9, "entry_taus_ceiling")

    crear_input(frame_consecutivos, "Taus/Ta:", 10)
    entry_taus_ta = crear_entry(frame_consecutivos, 10, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 10, "entry_taus_ta")

    ttk.Label(
        frame_consecutivos,
        text="Tipo de estator (seleccion unica) / Stator type (single selection):",
        style="FieldLabel.TLabel",
    ).grid(row=11, column=0, padx=10, pady=(8, 4), sticky="w")
    selector_estator = ttk.Frame(frame_consecutivos)
    selector_estator.grid(row=11, column=1, columnspan=3, padx=10, pady=(8, 4), sticky="w")
    ttk.Radiobutton(
        selector_estator,
        text="Estandar (std) / Standard (std)",
        variable=var_stator_tipo,
        value="std",
        command=actualizar_selector_estator,
    ).grid(row=0, column=0, padx=(0, 12), sticky="w")
    ttk.Radiobutton(
        selector_estator,
        text="Ranura Profunda / Deep Slot",
        variable=var_stator_tipo,
        value="ranura_profunda",
        command=actualizar_selector_estator,
    ).grid(row=0, column=1, padx=(0, 12), sticky="w")
    ttk.Radiobutton(
        selector_estator,
        text="Doble Cuadro / Double Cage",
        variable=var_stator_tipo,
        value="doble_cuadro",
        command=actualizar_selector_estator,
    ).grid(row=0, column=2, sticky="w")

    stator_widgets.clear()
    _, entry_ranura_profunda = _crear_fila_factor_estator(
        frame_consecutivos,
        "Factor Ranura Profunda / Deep Slot Factor:",
        12,
        "ranura_profunda",
    )
    _crear_check_consecutivo(frame_consecutivos, 12, "entry_ranura_profunda")
    entry_ranura_profunda.bind("<KeyRelease>", on_consecutivo_edit)

    _, entry_doble_cuadro = _crear_fila_factor_estator(
        frame_consecutivos,
        "Factor Doble Cuadro / Double Cage Factor:",
        13,
        "doble_cuadro",
    )
    _crear_check_consecutivo(frame_consecutivos, 13, "entry_doble_cuadro")
    entry_doble_cuadro.bind("<KeyRelease>", on_consecutivo_edit)

    _, entry_std = _crear_fila_factor_estator(
        frame_consecutivos,
        "Factor Estandar (std) / Standard Factor (std):",
        14,
        "std",
    )
    _crear_check_consecutivo(frame_consecutivos, 14, "entry_std")
    entry_std.bind("<KeyRelease>", on_consecutivo_edit)

    crear_input(
        frame_consecutivos,
        "V2aus:",
        15,
        help_text="Incremento de temperatura en estator (V2aus) / Stator temperature rise (V2aus).",
        unit_text="°C",
    )
    entry_v2aus = crear_entry(frame_consecutivos, 15, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 15, "entry_v2aus")

    crear_input(
        frame_consecutivos,
        "V2aus ceiling:",
        16,
        help_text="V2aus redondeado hacia abajo para iteraciones de arranques / V2aus rounded down for start iterations.",
        unit_text="°C",
    )
    entry_v2aus_ceiling = crear_entry(frame_consecutivos, 16, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 16, "entry_v2aus_ceiling")

    crear_input(frame_consecutivos, "Numero de arranques (frio y caliente) / Number of starts (cold and hot):", 17)
    entry_num_arranques = crear_entry(frame_consecutivos, 17)
    entry_num_arranques.insert(0, "10")
    _crear_check_consecutivo(frame_consecutivos, 17, "entry_num_arranques")
    entry_num_arranques.bind("<KeyRelease>", on_consecutivo_edit)
    crear_input(frame_consecutivos, "Arranques por Hora / Starts per Hour:", 18)
    entry_arranques_hora = crear_entry(frame_consecutivos, 18)
    entry_arranques_hora.insert(0, "4")
    _crear_check_consecutivo(frame_consecutivos, 18, "entry_arranques_hora")
    entry_arranques_hora.bind("<KeyRelease>", on_consecutivo_edit)
    crear_input(frame_consecutivos, "Tiempo ciclo arranque-paro (seg) / Start-stop cycle time (sec):", 19)
    global entry_ciclo_arr, entry_resto_arr, entry_reposo_ciclo
    entry_ciclo_arr = crear_entry(frame_consecutivos, 19, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 19, "entry_ciclo_arr")
    crear_input(frame_consecutivos, "Tiempo restante reposo sin ventilacion (seg) / Remaining no-vent rest time (sec):", 20)
    entry_resto_arr = crear_entry(frame_consecutivos, 20, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 20, "entry_resto_arr")
    entry_reposo_ciclo = crear_entry(frame_consecutivos, 21, readonly=True)
    _crear_check_consecutivo(frame_consecutivos, 21, "entry_reposo_ciclo")

    frame_grafica = ttk.Labelframe(
        frame_consecutivos,
        text="Grafica Arranques en Frio / Cold Starts Chart",
        padding=10,
        style="Card.TLabelframe",
    )
    frame_grafica.grid(row=22, column=0, columnspan=6, padx=5, pady=(8, 4), sticky="ew")
    frame_grafica.columnconfigure(0, weight=1)
    canvas_arranques = tk.Canvas(
        frame_grafica,
        height=240,
        background="#ffffff",
        highlightthickness=1,
        highlightbackground=UI_THEME["border"],
    )
    canvas_arranques.grid(row=0, column=0, sticky="ew")

    frame_tabla_arranques = ttk.Labelframe(
        frame_consecutivos,
        text="Tabla Arranques en Frio / Cold Starts Table",
        padding=10,
        style="Card.TLabelframe",
    )
    frame_tabla_arranques.grid(row=23, column=0, columnspan=6, padx=5, pady=(4, 6), sticky="ew")
    frame_tabla_arranques.columnconfigure(0, weight=1)
    tree_arranques = ttk.Treeview(
        frame_tabla_arranques,
        columns=("arr", "temp", "seg", "tipo"),
        show="headings",
        height=8,
        style="Modern.Treeview",
    )
    tree_arranques.heading("arr", text="Arr.")
    tree_arranques.heading("temp", text="C")
    tree_arranques.heading("seg", text="Seg / Sec")
    tree_arranques.heading("tipo", text="Tipo / Type")
    tree_arranques.column("arr", width=60, anchor="center")
    tree_arranques.column("temp", width=80, anchor="center")
    tree_arranques.column("seg", width=80, anchor="center")
    tree_arranques.column("tipo", width=120, anchor="w")
    tree_arranques.grid(row=0, column=0, sticky="ew")
    scrollbar_arranques = ttk.Scrollbar(frame_tabla_arranques, orient="vertical", command=tree_arranques.yview)
    tree_arranques.configure(yscroll=scrollbar_arranques.set)
    scrollbar_arranques.grid(row=0, column=1, sticky="ns")

    frame_grafica_caliente = ttk.Labelframe(
        frame_consecutivos,
        text="Grafica Arranques en Caliente / Hot Starts Chart",
        padding=10,
        style="Card.TLabelframe",
    )
    frame_grafica_caliente.grid(row=24, column=0, columnspan=6, padx=5, pady=(8, 4), sticky="ew")
    frame_grafica_caliente.columnconfigure(0, weight=1)
    canvas_arranques_caliente = tk.Canvas(
        frame_grafica_caliente,
        height=240,
        background="#ffffff",
        highlightthickness=1,
        highlightbackground=UI_THEME["border"],
    )
    canvas_arranques_caliente.grid(row=0, column=0, sticky="ew")

    frame_tabla_arranques_caliente = ttk.Labelframe(
        frame_consecutivos,
        text="Tabla Arranques en Caliente / Hot Starts Table",
        padding=10,
        style="Card.TLabelframe",
    )
    frame_tabla_arranques_caliente.grid(row=25, column=0, columnspan=6, padx=5, pady=(4, 6), sticky="ew")
    frame_tabla_arranques_caliente.columnconfigure(0, weight=1)
    tree_arranques_caliente = ttk.Treeview(
        frame_tabla_arranques_caliente,
        columns=("arr", "temp", "seg", "tipo"),
        show="headings",
        height=8,
        style="Modern.Treeview",
    )
    tree_arranques_caliente.heading("arr", text="Arr.")
    tree_arranques_caliente.heading("temp", text="C")
    tree_arranques_caliente.heading("seg", text="Seg / Sec")
    tree_arranques_caliente.heading("tipo", text="Tipo / Type")
    tree_arranques_caliente.column("arr", width=60, anchor="center")
    tree_arranques_caliente.column("temp", width=80, anchor="center")
    tree_arranques_caliente.column("seg", width=80, anchor="center")
    tree_arranques_caliente.column("tipo", width=120, anchor="w")
    tree_arranques_caliente.grid(row=0, column=0, sticky="ew")
    scrollbar_arranques_caliente = ttk.Scrollbar(
        frame_tabla_arranques_caliente, orient="vertical", command=tree_arranques_caliente.yview
    )
    tree_arranques_caliente.configure(yscroll=scrollbar_arranques_caliente.set)
    scrollbar_arranques_caliente.grid(row=0, column=1, sticky="ns")

    frame_tabla_arranques_hora = ttk.Labelframe(
        frame_consecutivos,
        text="Arranques por Hora / Starts per Hour",
        padding=10,
        style="Card.TLabelframe",
    )
    frame_tabla_arranques_hora.grid(row=26, column=0, columnspan=6, padx=5, pady=(4, 6), sticky="ew")
    frame_tabla_arranques_hora.columnconfigure(0, weight=1)
    tree_arranques_hora = ttk.Treeview(
        frame_tabla_arranques_hora,
        columns=("arr", "t", "temp"),
        show="headings",
        height=6,
        style="Modern.Treeview",
    )
    tree_arranques_hora.heading("arr", text="Arr.")
    tree_arranques_hora.heading("t", text="t (seg/sec)")
    tree_arranques_hora.heading("temp", text="C")
    tree_arranques_hora.column("arr", width=60, anchor="center")
    tree_arranques_hora.column("t", width=100, anchor="center")
    tree_arranques_hora.column("temp", width=100, anchor="center")
    tree_arranques_hora.grid(row=0, column=0, sticky="ew")
    scrollbar_arranques_hora = ttk.Scrollbar(
        frame_tabla_arranques_hora, orient="vertical", command=tree_arranques_hora.yview
    )
    tree_arranques_hora.configure(yscroll=scrollbar_arranques_hora.set)
    scrollbar_arranques_hora.grid(row=0, column=1, sticky="ns")

    frame_grafica_hora = ttk.Labelframe(
        frame_consecutivos,
        text="Grafica Arranques por Hora / Starts per Hour Chart",
        padding=10,
        style="Card.TLabelframe",
    )
    frame_grafica_hora.grid(row=27, column=0, columnspan=6, padx=5, pady=(8, 10), sticky="ew")
    frame_grafica_hora.columnconfigure(0, weight=1)
    canvas_arranques_hora = tk.Canvas(
        frame_grafica_hora,
        height=200,
        background="#ffffff",
        highlightthickness=1,
        highlightbackground=UI_THEME["border"],
    )
    canvas_arranques_hora.grid(row=0, column=0, sticky="ew")

    # BotÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n Calcular
    frame_boton = ttk.Frame(frame_scroll, style="Hero.TFrame", padding=(12, 10))
    frame_boton.pack(pady=10, padx=20, fill="x")
    frame_boton.columnconfigure(0, weight=1)
    frame_boton.columnconfigure(1, weight=1)

    ttk.Button(
        frame_boton, text="Reiniciar / Reset", command=reiniciar_calculos
    ).grid(row=0, column=0, padx=5, pady=5)

    btn_calcular = ttk.Button(
        frame_boton, text="Calcular / Calculate", command=calcular_datos, style="Accent.TButton"
    )

    ttk.Button(
        frame_boton, text="Exportar PDF / Export PDF", command=exportar_pdf_datos_motor, style="Accent.TButton"
    ).grid(row=1, column=0, pady=5)

    btn_debug = ttk.Button(
        frame_boton, text="Operaciones utilizadas / Operations used", command=mostrar_ventana_debug
    )
    btn_debug.grid(row=1, column=1, pady=5)

    actualizar_selector_estator()

    # Initialize runtime state after full widget tree is ready.
    reiniciar_calculos()





def seleccionar_archivo():
    """Prompt the user for an Excel file and store the selected path."""

    from tkinter import filedialog

    ruta = filedialog.askopenfilename(
        parent=form_toplevel, filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if ruta:
        mlfb_config["archivo"] = ruta
        guardar_configuracion(mlfb_config)
        entry_excel_path.config(state="normal")
        entry_excel_path.delete(0, tk.END)
        entry_excel_path.insert(0, ruta)
        entry_excel_path.config(state="readonly", style="ReadonlyDark.TEntry")


def guardar_columnas():
    """Persist the Excel column letters entered by the user."""

    mlfb_config["col_mlfb"] = entry_col_mlfb.get().strip().upper()
    mlfb_config["col_hp"] = entry_col_hp.get().strip().upper()
    mlfb_config["col_inercia"] = entry_col_inercia.get().strip().upper()
    mlfb_config["col_curvas"] = entry_col_curvas.get().strip().upper()
    guardar_configuracion(mlfb_config)
    messagebox.showinfo(
        "Columnas actualizadas / Columns updated",
        "Las columnas han sido guardadas. / Columns have been saved.",
        parent=form_toplevel,
    )


def crear_input(frame, text, row, help_text: str | None = None, unit_text: str | None = None):
    """Utility to create a labeled field with optional help tooltip and unit."""
    ttk.Label(frame, text=text, style="FieldLabel.TLabel").grid(
        row=row, column=0, padx=10, pady=5, sticky="w"
    )
    if help_text:
        ayuda = ttk.Label(frame, text="?", style="Muted.TLabel", cursor="question_arrow")
        ayuda.grid(row=row, column=5, padx=5)
        ToolTip(ayuda, help_text)
    if unit_text:
        ttk.Label(frame, text=unit_text, style="Muted.TLabel").grid(
            row=row, column=2, padx=5, sticky="w"
        )
from datetime import datetime

if FPDF is None:
    PDF = None
else:
    class PDF(FPDF):
        """Custom FPDF subclass with ABB-themed header and footer."""

        def header(self):
            # Fondo
            if self.page_no() <= 2:
                try:
                    fondo_path = os.path.join(os.path.dirname(__file__), "fondo_motor.png")
                    self.image(fondo_path, x=0, y=0, w=210, h=297)  # A4 tamaÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±o completo
                except Exception as e:
                    print(f"[ERROR FONDO] {e}")

            # Logo ABB
            try:
                ruta_logo = os.path.join(os.path.dirname(__file__), "logo.png")
                self.image(ruta_logo, x=10, y=8, w=30)
            except Exception as e:
                print(f"[ERROR LOGO] {e}")

            # Imagen decorativa motor
            try:
                ruta_motor = os.path.join(os.path.dirname(__file__), "motor_header.png")
                self.image(ruta_motor, x=165, y=15, w=40)
            except Exception as e:
                print(f"[ERROR IMAGEN MOTOR] {e}")

            # TÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â­tulo central
            self.set_font("Arial", "B", 12)
            self.set_text_color(0, 0, 0)
            self.cell(0, 10, "Calculo de Tiempo de Arranque / Motor Start-Up Calculation", 0, 1, "C")

            try:
                mlfb_actual = entry_mlfb.get().strip().upper()
                self.set_font("Arial", "I", 10)
                self.cell(0, 8, f"MLFB: {mlfb_actual}", 0, 1, "C")
            except Exception:
                pass

            self.ln(4)

        def footer(self):
            self.set_y(-20)
            self.set_font("Arial", "", 8)
            self.set_text_color(100)
            fecha = datetime.now().strftime("%d/%m/%Y")
            self.cell(0, 5, f"Fecha / Date: {fecha}", 0, 0, "C")

def limpiar_texto_para_pdf(texto: str) -> str:
    return texto.encode("latin-1", "ignore").decode("latin-1")

def add_ficha(pdf: Any, titulo: str, datos: list):
    titulo_limpio = limpiar_texto_para_pdf(titulo)
    pdf.set_font("Arial", "B", 11)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(0, 10, titulo_limpio, ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(0, 0, 0)

    for campo, descripcion_bi, unidad in datos:
        if hasattr(campo, "get"):
            valor = limpiar_texto_para_pdf(str(campo.get()).strip())
        else:
            valor = limpiar_texto_para_pdf(str(campo).strip())
        descripcion_limpia = limpiar_texto_para_pdf(f"{descripcion_bi}:")
        valor_limpio = limpiar_texto_para_pdf(formato_valor_pdf(valor))

        pdf.cell(130, 6, descripcion_limpia, ln=0)
        pdf.cell(0, 6, f"    {valor_limpio} {unidad}", ln=1)


    pdf.ln(5)


def formato_valor_pdf(valor):
    """Format numeric values for PDF output, trimming trailing zeros."""
    try:
        num = float(valor)
        if num.is_integer():
            return str(int(num))
        return str(Decimal(str(num))
                   .quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP)
                   ).rstrip("0").rstrip(".")
    except:
        # AquÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â­, como ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ltimo recurso, forzamos la limpieza:
        return limpiar_texto_para_pdf(str(valor))


    
def dibujar_grafica_pdf(
    pdf: Any,
    temps: list[float],
    max_temp: float,
    *,
    graph_x: float,
    graph_y: float,
    graph_w: float,
    graph_h: float,
    x_label: str,
    tick_mode: str,
    num_arranques: int | None = None,
) -> None:
    """Dibuja una grafica con bandas de estado y linea de temperatura maxima."""

    if not temps:
        return

    max_val = max(temps)
    estado = obtener_estado_temperatura(max_temp, max_val)
    line_rgb = PDF_COLOR_ESTADO.get(estado, (11, 79, 138))

    y_max = max(max_temp, max_val) * 1.1
    if y_max <= 0:
        y_max = 1

    def to_point(idx: int, temp: float) -> tuple[float, float]:
        x = graph_x + (graph_w * idx / max(1, len(temps) - 1))
        y = graph_y + graph_h - ((temp / y_max) * graph_h)
        return x, y

    # Zonas de estado
    if max_temp > 0:
        y_warn = graph_y + graph_h - ((max_temp * ESTADO_UMBRAL / y_max) * graph_h)
        y_limit = graph_y + graph_h - ((max_temp / y_max) * graph_h)
        pdf.set_fill_color(*PDF_BANDA_PELIGRO)
        pdf.rect(graph_x, graph_y, graph_w, max(0, y_limit - graph_y), style="F")
        pdf.set_fill_color(*PDF_BANDA_ALERTA)
        pdf.rect(graph_x, y_limit, graph_w, max(0, y_warn - y_limit), style="F")
        pdf.set_fill_color(*PDF_BANDA_OK)
        pdf.rect(graph_x, y_warn, graph_w, max(0, (graph_y + graph_h) - y_warn), style="F")

    # Borde
    pdf.set_draw_color(160, 160, 160)
    pdf.rect(graph_x, graph_y, graph_w, graph_h)

    # Grid y etiquetas Y
    pdf.set_font("Arial", "", 8)
    y_ticks = 5
    for i in range(y_ticks + 1):
        y_val = y_max * i / y_ticks
        y = graph_y + graph_h - ((y_val / y_max) * graph_h)
        pdf.set_draw_color(230, 230, 230)
        pdf.line(graph_x, y, graph_x + graph_w, y)
        pdf.set_text_color(80, 80, 80)
        pdf.text(graph_x - 12, y + 2, str(formato_valor_pdf(round(y_val, 1))))

    # Grid y etiquetas X
    if tick_mode == "arranques" and num_arranques:
        step = max(1, math.ceil(num_arranques / 6))
        for arr in range(1, num_arranques + 1, step):
            idx = 2 * (arr - 1)
            x = graph_x + (graph_w * idx / max(1, len(temps) - 1))
            pdf.set_draw_color(240, 240, 240)
            pdf.line(x, graph_y, x, graph_y + graph_h)
            pdf.text(x - 2, graph_y + graph_h + 5, str(arr))
    else:
        step = max(1, math.ceil(len(temps) / 6))
        for idx in range(0, len(temps), step):
            x = graph_x + (graph_w * idx / max(1, len(temps) - 1))
            pdf.set_draw_color(240, 240, 240)
            pdf.line(x, graph_y, x, graph_y + graph_h)
            pdf.text(x - 2, graph_y + graph_h + 5, str(idx + 1))

    pdf.set_text_color(60, 60, 60)
    pdf.text(graph_x + graph_w / 2 - 18, graph_y + graph_h + 12, x_label)
    pdf.text(graph_x - 10, graph_y - 2, "degC")

    # Linea de temp maxima
    y_line = graph_y + graph_h - ((max_temp / y_max) * graph_h)
    pdf.set_draw_color(200, 0, 0)
    pdf.line(graph_x, y_line, graph_x + graph_w, y_line)
    pdf.set_text_color(200, 0, 0)
    pdf.text(graph_x + graph_w - 20, y_line - 2, f"{formato_valor_pdf(max_temp)} degC")

    # Serie
    pdf.set_draw_color(*line_rgb)
    pdf.set_line_width(0.8)
    pts = [to_point(i, t) for i, t in enumerate(temps)]
    for i in range(len(pts) - 1):
        pdf.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1])
    pdf.set_fill_color(*line_rgb)
    for x, y in pts:
        pdf.ellipse(x - 0.6, y - 0.6, 1.2, 1.2, style="F")

    # Leyenda debajo de la grÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡fica
    pdf.set_text_color(*line_rgb)
    pdf.text(graph_x + graph_w / 2 - 20, graph_y + graph_h + 20, f"Estado: {estado}")
    pdf.set_line_width(0.2)
    pdf.set_text_color(0, 0, 0)


def dibujar_grafica_pdf_hora(
    pdf: Any,
    tiempos: list[float],
    temps: list[float],
    *,
    graph_x: float,
    graph_y: float,
    graph_w: float,
    graph_h: float,
) -> None:
    if not temps or not tiempos:
        return
    y_max = max(temps) * 1.1 if max(temps) > 0 else 1
    y_min = min(temps) * 0.9 if min(temps) > 0 else 0
    x_max = max(tiempos)

    def to_point(t_val: float, temp: float) -> tuple[float, float]:
        x = graph_x + (graph_w * t_val / max(1, x_max))
        y = graph_y + graph_h - ((temp - y_min) / max(1e-6, y_max - y_min)) * graph_h
        return x, y

    pdf.set_draw_color(160, 160, 160)
    pdf.rect(graph_x, graph_y, graph_w, graph_h)

    # grid Y
    pdf.set_font("Arial", "", 8)
    y_ticks = 5
    for i in range(y_ticks + 1):
        y_val = y_min + (y_max - y_min) * i / y_ticks
        y = graph_y + graph_h - ((y_val - y_min) / max(1e-6, y_max - y_min)) * graph_h
        pdf.set_draw_color(230, 230, 230)
        pdf.line(graph_x, y, graph_x + graph_w, y)
        pdf.set_text_color(80, 80, 80)
        pdf.text(graph_x - 12, y + 2, str(formato_valor_pdf(round(y_val, 1))))

    # grid X
    step = max(1, int(x_max // 6))
    for xv in range(0, int(x_max) + 1, step):
        x = graph_x + (graph_w * xv / max(1, x_max))
        pdf.set_draw_color(230, 230, 230)
        pdf.line(x, graph_y, x, graph_y + graph_h)
        pdf.set_text_color(80, 80, 80)
        pdf.text(x - 2, graph_y + graph_h + 5, str(xv))

    pts = [to_point(t, temp) for t, temp in zip(tiempos, temps)]
    pdf.set_draw_color(11, 79, 138)
    pdf.set_line_width(0.8)
    for i in range(len(pts) - 1):
        pdf.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1])
    pdf.set_fill_color(11, 79, 138)
    for x, y in pts:
        pdf.ellipse(x - 0.6, y - 0.6, 1.2, 1.2, style="F")
    pdf.set_text_color(60, 60, 60)
    pdf.text(graph_x + graph_w / 2 - 20, graph_y + graph_h + 12, "Tiempo (seg)")
    pdf.text(graph_x - 10, graph_y - 2, "degC")


def exportar_pdf_datos_motor():
    """Generate a detailed PDF report with all calculated values."""
    from tkinter import filedialog

    if FPDF is None:
        messagebox.showerror(
            "Dependencia faltante / Missing dependency",
            "La exportacion a PDF requiere la libreria 'fpdf'.\n"
            "PDF export requires the 'fpdf' library.\n"
            "Instale/Install: pip install fpdf2",
            parent=form_toplevel,
        )
        return

    nombre_mlfb = entry_mlfb.get().strip().upper().replace("/", "_") or "resultado_arranque"
    fecha_archivo = datetime.now().strftime("%d-%m-%Y")
    base = f"{nombre_mlfb} - {fecha_archivo}"

    output_path = filedialog.asksaveasfilename(
        parent=form_toplevel,
        title="Guardar PDF / Save PDF",
        defaultextension=".pdf",
        initialfile=f"{base}.pdf",
        filetypes=[("Archivos PDF / PDF Files", "*.pdf")],
    )
    if not output_path:
        return

    # Small progress dialog to provide feedback while generating the report.
    ventana_procesando = tk.Toplevel(form_toplevel)
    ventana_procesando.title("Procesando / Processing")
    etiqueta_procesando = ttk.Label(
        ventana_procesando, text="Procesando, por favor espere... / Processing, please wait..."
    )
    etiqueta_procesando.pack(padx=20, pady=20)
    ventana_procesando.update_idletasks()

    # Center the processing dialog over the form window
    try:
        x = form_toplevel.winfo_rootx() + (form_toplevel.winfo_width() // 2) - (
            ventana_procesando.winfo_width() // 2
        )
        y = form_toplevel.winfo_rooty() + (form_toplevel.winfo_height() // 2) - (
            ventana_procesando.winfo_height() // 2
        )
        ventana_procesando.geometry(f"+{x}+{y}")
    except Exception:
        pass

    ventana_procesando.update()

    # Build report in deterministic section order to ease support and auditing.
    pdf = PDF()
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()

    # ----------------- DATOS DEL MOTOR -----------------
    datos_motor = [
        (entry_hp, "Motor Power / Potencia nominal del motor (HP)", "HP"),
        (entry_rpm_motor, "Motor Speed / Velocidad nominal del motor en revoluciones por minuto", "min^-1"),
        (entry_torque_pct, "Applied Voltage / Porcentaje de voltaje alimentado", "%"),
        (entry_mnm, "Nominal Torque / Momento nominal del motor (MNM)", "kgm"),
        (entry_jm, "Rotor Inertia / Momento de inercia del rotor del motor (JM)", "kgm2"),
        (entry_gd2m, "Motor GD2 / Momento cuadratico del rotor (GD2M = 4*pi*JM)", "kgm2"),
        (entry_mmm, "Average Motor Torque / Momento medio del motor (MMM = MNM segun curva)", "kgm"),
    ]

    add_ficha(pdf, "MOTOR DATA / DATOS DEL MOTOR", datos_motor)

    # ----------------- DATOS DEL EQUIPO -----------------
    datos_equipo = [
        (entry_rpm_eq, "Driven Equipment Speed / Velocidad del equipo accionado (RPM)", "min^-1"),
        (entry_wr2, "Load Inertia WR2 / Momento de inercia del equipo (WR2)", "kgm2"),
        (entry_mnv, "Load Nominal Torque / Momento nominal del equipo (MNV)", "kgm"),
        (entry_mmv, "Average Load Torque / Momento medio del ventilador (MMV)", "kgm"),
        (entry_jv, "Load Inertia JV / Momento de inercia equivalente del ventilador (JV)", "kgm2"),
        (entry_mmv_nmot, "MMV referred to Motor / MMV referido a las RPM del motor", "kgm"),
        (entry_gd2v, "Load GD2 / Momento cuadratico del ventilador (GD2V = 4*pi*JV)", "kgm2"),
        (entry_gd2v_nmot, "GD2V referred to Motor / GD2V referido a las RPM del motor", "kgm2"),
        (entry_ngd2, "Total GD2 / Sumatoria de momentos (nGD2)", "kgm2"),
        # Solo se reporta el tiempo de arranque redondeado hacia arriba
        # pero se muestra como "Tiempo de arranque" en el PDF
        (entry_ta_ceiling, "Start-up Time (ta) / Tiempo de arranque (ta)", "seg"),
    ]

    add_ficha(pdf, "DRIVEN EQUIPMENT / EQUIPO ACCIONADO", datos_equipo)

    # ----------------- EVALUACION DE ARRANQUE (primera pagina) -----------------
    # Summary is intentionally compact and duplicated in both status and detail.
    ta_val = safe_float(entry_ta.get().strip()) if entry_ta.get().strip() else None
    stall_hot = safe_float(entry_stall_hot.get().strip()) if entry_stall_hot.get().strip() else None
    stall_cold = safe_float(entry_stall_cold.get().strip()) if entry_stall_cold.get().strip() else None
    mmm_val = safe_float(entry_mmm.get().strip()) if entry_mmm.get().strip() else None
    mmv_nmot_val = safe_float(entry_mmv_nmot.get().strip()) if entry_mmv_nmot.get().strip() else None

    if ta_val is None or stall_hot is None or stall_cold is None:
        texto_eval = "DATOS INSUFICIENTES PARA EVALUAR ARRANQUE / INSUFFICIENT DATA TO EVALUATE START-UP"
        detalle_eval = (
            "Complete ta, Stall HOT y Stall COLD para emitir un diagnostico valido / "
            "Complete ta, Stall HOT and Stall COLD for a valid diagnosis."
        )
        color_eval = (120, 120, 120)
    elif mmm_val is not None and mmv_nmot_val is not None and mmm_val <= mmv_nmot_val:
        texto_eval = "NO SOPORTA ARRANQUE (PAR MOTOR <= PAR CARGA) / DOES NOT SUPPORT START-UP (MOTOR TORQUE <= LOAD TORQUE)"
        detalle_eval = (
            "El par del motor no supera al de la carga, por lo que no acelera / "
            "Motor torque does not exceed load torque, so it cannot accelerate."
        )
        color_eval = (200, 0, 0)
    elif ta_val <= 0:
        texto_eval = "NO SOPORTA ARRANQUE (TA <= 0) / DOES NOT SUPPORT START-UP (TA <= 0)"
        detalle_eval = (
            "Valor no fisico o datos inconsistentes; revise entradas / "
            "Non-physical value or inconsistent data; review inputs."
        )
        color_eval = (200, 0, 0)
    else:
        if ta_val < stall_hot and ta_val < stall_cold:
            texto_eval = "SOPORTA ARRANQUE EN CALIENTE Y EN FRIO / SUPPORTS HOT AND COLD START-UP"
            detalle_eval = (
                "Ta esta por debajo de ambos limites Stall (HOT y COLD) / "
                "Ta is below both Stall limits (HOT and COLD)."
            )
            color_eval = (0, 120, 0)
        elif ta_val < stall_cold:
            texto_eval = "SOPORTA SOLO ARRANQUE EN FRIO / SUPPORTS ONLY COLD START-UP"
            detalle_eval = (
                "Cumple en frio pero excede en caliente / "
                "Passes in cold condition but exceeds in hot condition."
            )
            color_eval = (255, 140, 0)
        elif ta_val < stall_hot:
            texto_eval = "SOPORTA SOLO ARRANQUE EN CALIENTE / SUPPORTS ONLY HOT START-UP"
            detalle_eval = (
                "Cumple en caliente pero excede en frio / "
                "Passes in hot condition but exceeds in cold condition."
            )
            color_eval = (0, 140, 255)
        else:
            texto_eval = "NO SOPORTA ARRANQUE / DOES NOT SUPPORT START-UP"
            detalle_eval = (
                "Excede los limites Stall HOT y COLD / "
                "Exceeds both Stall HOT and COLD limits."
            )
            color_eval = (200, 0, 0)

    altura_bloque = 34  # titulo + caja de estado + mini explicacion
    if pdf.get_y() + altura_bloque > (pdf.h - pdf.b_margin):
        pdf.add_page()

    pdf.set_font("Arial", "B", 12)
    pdf.set_text_color(0)
    pdf.cell(0, 10, "EVALUACION DE ARRANQUE / START-UP EVALUATION", ln=True, align="C")

    pdf.set_fill_color(245, 245, 245)
    pdf.set_draw_color(*color_eval)
    pdf.set_text_color(*color_eval)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, limpiar_texto_para_pdf(texto_eval), border=1, ln=True, fill=True, align="C")
    pdf.set_text_color(80, 80, 80)
    pdf.set_font("Arial", "", 9)
    pdf.multi_cell(0, 5, limpiar_texto_para_pdf(detalle_eval), align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(4)

    if "var_consecutivos_activo" in globals() and var_consecutivos_activo.get():
        pdf.add_page()
        factor_estator, tipo_estator = obtener_factor_estator()
        datos_consecutivos = [
            (entry_f, "f=", ""),
            (entry_gfe, "GFe=", "kg"),
            (entry_diam_rotor, "Diam. Rotor en dm / Rotor Diam. in dm", "dm"),
            (entry_long_paquete, "Long. del paquete dm / Stack length in dm", "dm"),
            (entry_dibujo, "Dibujo paquete rotor / Rotor stack drawing", ""),
            (entry_diamxlongitud, "Diamxlongitud", "dm2"),
            (entry_v2a, "V'2A", "°C"),
            (entry_v2a_ceiling, "V'2A Ceiling", "°C"),
            (entry_taus, "taus", "seg"),
            (entry_taus_ceiling, "taus ceiling", "seg"),
            (entry_taus_ta, "Taus/Ta", ""),
            (tipo_estator, "Tipo de estator activo / Active stator type", ""),
            (formato_valor_pdf(factor_estator), "Factor estator activo / Active stator factor", ""),
            (entry_ranura_profunda, "Ranura Profunda / Deep Slot", ""),
            (entry_doble_cuadro, "Doble cuadro / Double Cage", ""),
            (entry_std, "std", ""),
            (entry_v2aus, "V2aus", "°C"),
            (entry_v2aus_ceiling, "V2aus ceiling", "°C"),
        ]
        add_ficha(pdf, "ARRANQUES CONSECUTIVOS / CONSECUTIVE STARTS", datos_consecutivos)

        pdf.add_page()
        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 8, "ARRANQUES EN FRIO - TABLA Y GRAFICA / COLD STARTS - TABLE AND CHART", ln=True)
        pdf.set_font("Arial", "B", 9)
        table_w = 25 + 25 + 20 + 35
        table_x = (pdf.w - table_w) / 2
        pdf.set_x(table_x)
        pdf.cell(25, 6, "Arr.", 1)
        pdf.cell(25, 6, "degC", 1)
        pdf.cell(20, 6, "Seg/Sec", 1)
        pdf.cell(35, 6, "Tipo/Type", 1)
        pdf.ln(6)

        num_arranques = safe_int(entry_num_arranques.get(), default=10)
        num_arranques = max(1, num_arranques)
        max_temp = TEMP_MAX_FRIO

        v2a_ceiling_val = safe_float(entry_v2a_ceiling.get())
        v2aus_ceiling_val = safe_float(entry_v2aus_ceiling.get())
        ta_val = safe_float(entry_ta.get())
        taus_ceiling_val = safe_float(entry_taus_ceiling.get())
        factor_estator, tipo_estator = obtener_factor_estator()

        filas = []
        temps = []
        if None not in (v2a_ceiling_val, ta_val, taus_ceiling_val, factor_estator):
            filas, temps = generar_series_arranques(
                num_arranques,
                v2a_ceiling_val,
                v2aus_ceiling_val,
                ta_val,
                taus_ceiling_val,
                factor_estator,
                max_temp,
            )

        pdf.set_font("Arial", "", 9)
        estados_tabla = estados_progresivos(filas, max_temp)
        for (arr, temp, seg, tipo), estado in zip(filas, estados_tabla):
            fill = PDF_BANDA_OK if estado == "OK" else PDF_BANDA_ALERTA if estado == "Cerca" else PDF_BANDA_PELIGRO
            pdf.set_fill_color(*fill)
            pdf.set_x(table_x)
            pdf.cell(25, 6, str(arr), 1, 0, fill=True)
            pdf.cell(25, 6, str(formato_valor_pdf(temp)), 1, 0, fill=True)
            pdf.cell(20, 6, str(formato_valor_pdf(seg)), 1, 0, fill=True)
            pdf.cell(35, 6, tipo, 1, 1, fill=True)

        # Grafica
        if temps:
            graph_x = 20
            graph_y = pdf.get_y() + 8
            graph_w = 170
            graph_h = 75
            temps_plot = temps
            estados_plot = estados_progresivos([(str(i), t, 0.0, "") for i, t in enumerate(temps)], max_temp)
            for idx, est in enumerate(estados_plot):
                if est == "Excede":
                    temps_plot = temps[: idx + 1]
                    break
            graph_x = (pdf.w - graph_w) / 2
            dibujar_grafica_pdf(
                pdf,
                temps_plot,
                max_temp,
                graph_x=graph_x,
                graph_y=graph_y,
                graph_w=graph_w,
                graph_h=graph_h,
                x_label="Arranques en frio / Cold starts",
                tick_mode="arranques",
                num_arranques=num_arranques,
            )

        # Pagina: Arranques en caliente
        pdf.add_page()
        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 8, "ARRANQUES EN CALIENTE / HOT STARTS", ln=True)

        num_arranques_hot = safe_int(entry_num_arranques.get(), default=10)
        num_arranques_hot = max(1, num_arranques_hot)
        max_temp_hot = TEMP_MAX_CALIENTE
        temp_inicial = TEMP_INICIAL_CALIENTE
        tx_val = TX_CALIENTE

        filas_hot = []
        temps_hot = []
        if None not in (v2a_ceiling_val, v2aus_ceiling_val, ta_val, taus_ceiling_val):
            filas_hot, temps_hot = generar_series_arranques_caliente(
                num_arranques_hot,
                v2a_ceiling_val,
                v2aus_ceiling_val,
                ta_val,
                taus_ceiling_val,
                temp_inicial,
                tx_val,
                max_temp_hot,
            )

        pdf.set_font("Arial", "B", 9)
        table_w = 25 + 25 + 20 + 35
        table_x = (pdf.w - table_w) / 2
        pdf.set_x(table_x)
        pdf.cell(25, 6, "Arr.", 1)
        pdf.cell(25, 6, "degC", 1)
        pdf.cell(20, 6, "Seg/Sec", 1)
        pdf.cell(35, 6, "Tipo/Type", 1)
        pdf.ln(6)
        pdf.set_font("Arial", "", 9)
        estados_hot_tabla = estados_progresivos(filas_hot, max_temp_hot)
        for (arr, temp, seg, tipo), estado in zip(filas_hot, estados_hot_tabla):
            fill = PDF_BANDA_OK if estado == "OK" else PDF_BANDA_ALERTA if estado == "Cerca" else PDF_BANDA_PELIGRO
            pdf.set_fill_color(*fill)
            pdf.set_x(table_x)
            pdf.cell(25, 6, str(arr), 1, 0, fill=True)
            pdf.cell(25, 6, str(formato_valor_pdf(temp)), 1, 0, fill=True)
            pdf.cell(20, 6, str(formato_valor_pdf(seg)), 1, 0, fill=True)
            pdf.cell(35, 6, tipo, 1, 1, fill=True)

        if temps_hot:
            graph_x = 20
            graph_y = pdf.get_y() + 12
            graph_w = 170
            graph_h = 85
            temps_hot_plot = temps_hot
            estados_hot = estados_progresivos([(str(i), t, 0.0, "") for i, t in enumerate(temps_hot)], max_temp_hot)
            for idx, est in enumerate(estados_hot):
                if est == "Excede":
                    temps_hot_plot = temps_hot[: idx + 1]
                    break
            graph_x = (pdf.w - graph_w) / 2
            dibujar_grafica_pdf(
                pdf,
                temps_hot_plot,
                max_temp_hot,
                graph_x=graph_x,
                graph_y=graph_y,
                graph_w=graph_w,
                graph_h=graph_h,
                x_label="tiempo / time",
                tick_mode="tiempo",
            )

        # Arranques por hora
        filas_hora, n_hora, ciclo, resto, reposo_ciclo = calcular_arranques_por_hora()
        pdf.add_page()
        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 8, f"ARRANQUES POR HORA / STARTS PER HOUR ({n_hora} por hora/per hour)", ln=True, align="C")
        pdf.set_font("Arial", "", 9)
        pdf.cell(0, 6, f"Tiempo de 1 ciclo arranque-paro / Start-stop cycle time: {formato_valor_pdf(ciclo)} seg/sec", ln=True, align="C")
        pdf.cell(0, 6, f"Tiempo restante de reposo sin ventilacion / Remaining no-vent rest time: {formato_valor_pdf(resto)} seg/sec", ln=True, align="C")
        pdf.cell(0, 6, f"Tiempo de reposo sin ventilacion por ciclo / No-vent rest time per cycle: {formato_valor_pdf(reposo_ciclo)} seg/sec", ln=True, align="C")
        pdf.ln(3)

        table_w = 25 + 35 + 35
        table_x = (pdf.w - table_w) / 2
        pdf.set_x(table_x)
        pdf.set_font("Arial", "B", 9)
        pdf.cell(25, 6, "Arr.", 1)
        pdf.cell(35, 6, "t (seg)", 1)
        pdf.cell(35, 6, "degC", 1)
        pdf.ln(6)
        pdf.set_font("Arial", "", 9)
        for arr, t_abs, temp in filas_hora:
            pdf.set_x(table_x)
            pdf.cell(25, 6, str(arr), 1)
            pdf.cell(35, 6, formato_valor_pdf(t_abs), 1)
            pdf.cell(35, 6, formato_valor_pdf(temp), 1, ln=1)

        if filas_hora:
            graph_w = 170
            graph_h = 70
            graph_x = (pdf.w - graph_w) / 2
            graph_y = pdf.get_y() + 8
            tiempos_plot = [t for _, t, _ in filas_hora]
            temps_plot = [temp for _, _, temp in filas_hora]
            dibujar_grafica_pdf_hora(
                pdf,
                tiempos_plot,
                temps_plot,
                graph_x=graph_x,
                graph_y=graph_y,
                graph_w=graph_w,
                graph_h=graph_h,
            )

    # ----------------- GUARDAR -----------------
    # Save once and then open both folder and artifact for operator convenience.
    try:
        pdf.output(output_path)
        etiqueta_procesando.config(text="Proceso completado / Process completed")
        ventana_procesando.update()
        messagebox.showinfo(
            "PDF generado / PDF generated",
            f"El archivo se ha guardado en / File saved at:\n{output_path}",
            parent=form_toplevel,
        )
        abrir_ruta(os.path.dirname(output_path))
        abrir_ruta(output_path)
    except Exception as e:
        messagebox.showerror(
            "Error al guardar PDF / Error saving PDF",
            f"No se pudo guardar el PDF / Could not save PDF:\n{e}",
            parent=form_toplevel,
        )
    finally:
        ventana_procesando.destroy()


def crear_entry(
    frame,
    row,
    readonly: bool = False,
    *,
    state: str | None = None,
    style: str | None = None,
) -> ttk.Entry:
    """Crea un ``ttk.Entry`` con estilo y estado opcional."""

    if state is None:
        state = "normal" if not readonly else "readonly"
    if style is None:
        style = "TEntry" if not readonly else "ReadonlyDark.TEntry"

    entry = ttk.Entry(frame, state=state, style=style)
    entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
    return entry

def mostrar_ventana_debug():
    """Display logged operations and table formulas in a structured window."""
    # Force a fresh silent calculation so traces match current inputs.
    calcular_datos(silencioso=True, actualizar_botones=False, debug_always=True)
    ventana_debug = tk.Toplevel(form_toplevel)
    ventana_debug.title("Operaciones Utilizadas")
    ventana_debug.geometry("1100x700")

    notebook = ttk.Notebook(ventana_debug)
    notebook.pack(expand=True, fill="both")

    # PestaÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±a 1: detalle de operaciones (texto)
    # Tab 1: raw operation trace as generated by debug_print/logger.
    tab_texto = ttk.Frame(notebook)
    notebook.add(tab_texto, text="Detalle / Details")
    texto_debug = scrolledtext.ScrolledText(tab_texto, wrap=tk.WORD, font=("Consolas", 10))
    texto_debug.pack(expand=True, fill="both")
    try:
        with open("debug_log.txt", "r", encoding="utf-8") as f:
            texto_debug.insert(tk.END, f.read())
    except Exception as e:
        texto_debug.insert(tk.END, f"Error al leer debug_log.txt / Error reading debug_log.txt:\n{e}")

    # PestaÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±a 2: tablas con operaciÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n usada
    # Tab 2: computed tables enriched with operation formulas by row.
    tab_tablas = ttk.Frame(notebook)
    notebook.add(tab_tablas, text="Tablas / Tables")
    tab_tablas.columnconfigure(0, weight=1)
    tab_tablas.rowconfigure(0, weight=1)
    tab_tablas.rowconfigure(1, weight=1)
    tab_tablas.rowconfigure(2, weight=1)

    def _crear_tabla(parent: ttk.Frame, titulo: str):
        frame = ttk.Labelframe(parent, text=titulo, padding=6)
        frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(
            frame,
            columns=("arr", "temp", "seg", "tipo", "estado", "operacion"),
            show="headings",
            height=8,
        )
        tree.heading("arr", text="Arr.")
        tree.heading("temp", text="Temp")
        tree.heading("seg", text="Seg/Sec")
        tree.heading("tipo", text="Tipo/Type")
        tree.heading("estado", text="Estado/Status")
        tree.heading("operacion", text="Operacion usada / Operation used")
        tree.column("arr", width=60, anchor="center")
        tree.column("temp", width=90, anchor="center")
        tree.column("seg", width=90, anchor="center")
        tree.column("tipo", width=120, anchor="w")
        tree.column("estado", width=90, anchor="center")
        tree.column("operacion", width=550, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        return frame, tree

    frame_f, tree_f = _crear_tabla(tab_tablas, "Tabla Arranques en Frio / Cold Starts Table")
    frame_f.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)
    frame_c, tree_c = _crear_tabla(tab_tablas, "Tabla Arranques en Caliente / Hot Starts Table")
    frame_c.grid(row=1, column=0, sticky="nsew", padx=8, pady=6)
    frame_h, tree_h = _crear_tabla(tab_tablas, "Tabla Arranques por Hora / Starts per Hour Table")
    frame_h.grid(row=2, column=0, sticky="nsew", padx=8, pady=6)

    def _estado_temp(temp: float, max_temp: float) -> str:
        return obtener_estado_temperatura(max_temp, temp)

    def _cargar_desde_tree(origen: ttk.Treeview, destino: ttk.Treeview, max_temp: float, modo: str):
        # Rebuild rows so exported operations remain understandable post-run.
        for iid in origen.get_children():
            vals = origen.item(iid, "values")
            if not vals:
                continue
            arr = str(vals[0]) if len(vals) > 0 else ""
            temp_txt = str(vals[1]) if len(vals) > 1 else ""
            seg_txt = str(vals[2]) if len(vals) > 2 else ""
            tipo = str(vals[3]) if len(vals) > 3 else ""
            temp_val = safe_float(temp_txt) or 0.0
            estado = _estado_temp(temp_val, max_temp)

            if modo == "frio":
                if tipo == "calentamiento":
                    oper = "T_cal = T_prev + V'2A_ceiling; seg = ta"
                else:
                    oper = "T_enf = (i=1 ? V2aus_ceiling : T_cal * factor_estator_activo); seg = taus_ceiling"
            elif modo == "caliente":
                if tipo == "calentamiento":
                    oper = "T_cal = T_prev + V'2A_ceiling; seg = ta/tx"
                else:
                    oper = "T_enf = T_cal - V2aus_ceiling; seg = taus_ceiling"
            else:
                oper = "t = t_prev + ta_ceiling + taus_ceiling + reposo_ciclo; T = T_prev + V'2A_ceiling"

            destino.insert("", "end", values=(arr, temp_txt, seg_txt, tipo or "-", estado, oper))

    if "tree_arranques" in globals():
        _cargar_desde_tree(tree_arranques, tree_f, TEMP_MAX_FRIO, "frio")
    if "tree_arranques_caliente" in globals():
        _cargar_desde_tree(tree_arranques_caliente, tree_c, TEMP_MAX_CALIENTE, "caliente")
    if "tree_arranques_hora" in globals():
        # convertir a formato comun para que vea tabla con operacion
        for iid in tree_arranques_hora.get_children():
            vals = tree_arranques_hora.item(iid, "values")
            if not vals:
                continue
            arr = str(vals[0]) if len(vals) > 0 else ""
            tseg = str(vals[1]) if len(vals) > 1 else ""
            temp_txt = str(vals[2]) if len(vals) > 2 else ""
            temp_val = safe_float(temp_txt) or 0.0
            estado = _estado_temp(temp_val, TEMP_MAX_CALIENTE)
            oper = "Arr_i: t_heat=t_prev+ta; t_cool=t_heat+taus; t_rep=t_cool+reposo_ciclo; T_rep=T_cool-(reposo_ciclo/60)"
            tree_h.insert("", "end", values=(arr, temp_txt, tseg, "arranque/hora", estado, oper))


# ==== Ventanas adicionales ====
def construir_adaptaciones(parent):
    """Create and layout the read-only fields for calculated adaptations."""
    global entry_mmv_nmot, entry_gd2v, entry_gd2v_nmot, entry_ngd2, entry_ta
    global label_check_mmv_nmot, label_check_gd2v, label_check_gd2v_nmot
    global label_check_ngd2, label_check_ta, label_check_ta_ceiling
    frame = ttk.Labelframe(parent, text="Adaptaciones / Adaptations", padding=10)
    frame.pack(padx=20, pady=5, expand=True, fill="both")
    frame.columnconfigure(1, weight=1)
    frame.columnconfigure(4, weight=0)
    crear_input(
        frame,
        "MMV nMot. / MMV at Motor Speed:",
        0,
        help_text="Momento Medio del Ventilador (segun curva) / Average Fan Torque (according to curve)",
        unit_text="kgm",
    )
    entry_mmv_nmot = crear_entry(frame, 0, readonly=True)
    label_check_mmv_nmot = ttk.Label(frame, text="")
    label_check_mmv_nmot.grid(row=0, column=3, padx=5)
    crear_input(
        frame,
        "GD2V:",
        1,
        help_text="Momento cuarto de inercia del ventilador (GD2V) / Fan fourth inertia moment (GD2V).",
        unit_text="kgm2",
    )
    entry_gd2v = crear_entry(frame, 1, readonly=True)
    label_check_gd2v = ttk.Label(frame, text="")
    label_check_gd2v.grid(row=1, column=3, padx=5)
    crear_input(
        frame,
        "GD2V nMot.:",
        2,
        help_text="GD2 del ventilador referido a las RPM del motor / Fan GD2 referred to motor RPM.",
        unit_text="kgm2",
    )
    entry_gd2v_nmot = crear_entry(frame, 2, readonly=True)
    label_check_gd2v_nmot = ttk.Label(frame, text="")
    label_check_gd2v_nmot.grid(row=2, column=3, padx=5)
    crear_input(
        frame,
        "nGD2:",
        3,
        help_text="Sumatoria de momentos de inercia referidos al motor / Sum of inertia moments referred to motor.",
        unit_text="kgm2",
    )
    entry_ngd2 = crear_entry(frame, 3, readonly=True)
    label_check_ngd2 = ttk.Label(frame, text="")
    label_check_ngd2.grid(row=3, column=3, padx=5)
    crear_input(
        frame,
        "ta:",
        4,
        help_text="Tiempo de arranque del motor (ta) / Motor start-up time (ta).",
        unit_text="seg",
    )
    entry_ta = crear_entry(frame, 4, readonly=True)
    label_check_ta = ttk.Label(frame, text="")
    label_check_ta.grid(row=4, column=3, padx=5)

    crear_input(
        frame,
        "ta CEILING:",
        5,
        help_text="Tiempo de arranque redondeado para evaluaciones conservadoras / Rounded start-up time for conservative evaluations.",
        unit_text="seg",
    )
    global entry_ta_ceiling
    entry_ta_ceiling = crear_entry(frame, 5, readonly=True)
    label_check_ta_ceiling = ttk.Label(frame, text="")
    label_check_ta_ceiling.grid(row=5, column=3, padx=5)

    return frame







