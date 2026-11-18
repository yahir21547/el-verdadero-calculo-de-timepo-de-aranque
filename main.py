"""GUI entry point for selecting motor calculation tools."""

from __future__ import annotations

import logging
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from PIL import Image, ImageTk

from style_utils import aplicar_colorimetria

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Main window text resources and layout constants
# These constants centralize copy and sizing to keep UI updates simple.
# ---------------------------------------------------------------------------
TITLE = "Plataforma de Calculos de Motores / Motor Calculations Platform"
WELCOME_TEXT = (
    "Bienvenido a la Plataforma de Calculos de Motores / "
    "Welcome to the Motor Calculations Platform"
)
QUESTION_TEXT = (
    "Que tipo de calculo quieres realizar? / "
    "What type of calculation do you want to perform?"
)
OPTION_CALCULO = "Calculo de Tiempo de Arranque / Start-Up Time Calculation"
DEFAULT_OPTION = "Seleccionar... / Select..."
HEADER_IMG_SIZE = (100, 100)
WINDOW_GEOMETRY = "500x300"

# Root window is created before importing child forms that depend on Tk state.
root = tk.Tk()

from formulario_arranque import abrir_formulario_arranque


def _asset_path(filename: str) -> Path:
    """Build an absolute path to an asset inside this module directory."""
    return Path(__file__).resolve().parent / filename


def _load_header_image(path: Path, label: str) -> ImageTk.PhotoImage | None:
    """Load and resize a header image for Tkinter labels."""
    try:
        image = Image.open(path)
        image.thumbnail(HEADER_IMG_SIZE)
        return ImageTk.PhotoImage(image)
    except Exception as exc:
        logger.warning("[%s] No se pudo cargar %s: %s", label, path, exc)
        return None


def _configure_window(window: tk.Tk) -> None:
    """Configure root window geometry and style."""
    window.title(TITLE)
    window.geometry(WINDOW_GEOMETRY)
    window.resizable(True, True)

    aplicar_colorimetria(window)

    style = ttk.Style(window)
    style.configure("TButton", font=("Arial", 11))
    style.configure("TLabel", font=("Arial", 11))


def _build_layout(window: tk.Tk) -> tuple[ttk.Frame, tk.StringVar, ttk.Combobox]:
    """Create and place all main widgets."""
    window.columnconfigure(0, weight=1)
    window.rowconfigure((0, 1, 2, 3, 4), weight=1)

    frame = ttk.Frame(window)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(0, weight=1)
    for row in range(5):
        frame.rowconfigure(row, weight=1)

    header = ttk.Frame(frame)
    header.grid(row=0, column=0, pady=(10, 0))
    if getattr(window, "logo_img", None):
        ttk.Label(header, image=window.logo_img).pack(side="left", padx=10)
    if getattr(window, "motor_img", None):
        ttk.Label(header, image=window.motor_img).pack(side="right", padx=10)

    ttk.Label(frame, text=WELCOME_TEXT, font=("Arial", 14, "bold")).grid(
        row=1,
        column=0,
        pady=(20, 10),
        sticky="n",
    )
    ttk.Label(frame, text=QUESTION_TEXT).grid(row=2, column=0, pady=10, sticky="n")

    # Keep the selector extensible in case new tools are enabled later.
    options = [DEFAULT_OPTION, OPTION_CALCULO]
    selection = tk.StringVar(value=options[0])
    combo = ttk.Combobox(frame, textvariable=selection, values=options, state="readonly")
    combo.current(0)
    combo.grid(row=3, column=0, pady=5, sticky="n")

    return frame, selection, combo


def continuar() -> None:
    """Launch the tool selected in the combo box."""
    selected_option = seleccion.get()
    if selected_option == OPTION_CALCULO:
        # Preserve current screen coordinates so child form opens in-context.
        root.update_idletasks()
        pos_x = root.winfo_rootx()
        pos_y = root.winfo_rooty()
        root.withdraw()
        abrir_formulario_arranque(root, pos_x, pos_y)
    else:
        messagebox.showwarning(
            "Advertencia / Warning",
            "Por favor, selecciona un calculo valido / Please select a valid calculation.",
            parent=root,
        )


# Boot sequence: style, load assets, build UI, and enter Tk event loop.
_configure_window(root)

root.logo_img = _load_header_image(_asset_path("logo.png"), "LOGO")
root.motor_img = _load_header_image(_asset_path("motor_header.png"), "MOTOR")

main_frame, seleccion, combobox = _build_layout(root)

ttk.Button(main_frame, text="Continuar / Continue", command=continuar).grid(
    row=4,
    column=0,
    pady=20,
    sticky="n",
)

root.mainloop()
