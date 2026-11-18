"""Tkinter UI for comparing and decoding MLFB codes."""

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from PIL import Image, ImageTk
import os

from style_utils import ABB_COLORS, aplicar_colorimetria


# Ruta al icono que se muestra en la ventana del comparador
LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.ico")

# Diccionario que describe el significado de cada carácter por posición
significado_por_posicion = {
    1: {"1": "Motor Series: Standard or Definite Purpose"},
    2: {"L": "Standard GP/SD", "P": "Definite Purpose"},
    3: {"E": "Standard GP/SD", "C": "Definite Purpose"},
    4: {"4": "SD200", "6": "Next Gen NEMA"},
    5: {"3": "SD200", "5": "DP200 HPS"},
    6: {"2": "Eficiencia Estándar"},
    7: {"1": "Tipo Básico", "2": "Tipo Mejorado"},
    8: {"4": "Frame 444-449", "5": "Frame 5000s"},
    9: {
        "B": "Shaft 444-445T",
        "C": "Shaft 447-449T",
        "D": "Shaft L449T",
        "A": "Shaft 509-5011",
        "F": "Shaft 444-445TS",
        "G": "Shaft 447-449TS",
    },
    10: {"A": "2 polos", "B": "4 polos", "C": "6 polos", "D": "8 polos"},
}


def abrir_comparador_mlfb(root_principal):
    """Open a window to compare two MLFB codes character by character."""
    ventana = tk.Toplevel(root_principal)
    ventana.title("Comparador de MLFB")
    ventana.geometry("600x550")
    ventana.resizable(True, True)

    aplicar_colorimetria(ventana)

    def cerrar():
        """Close the comparator window and return to the main window."""
        ventana.destroy()
        root_principal.deiconify()

    ventana.protocol("WM_DELETE_WINDOW", cerrar)

    def get_significado(pos, valor):
        """Return the human description for a character in the MLFB string."""
        return significado_por_posicion.get(pos, {}).get(valor, "Sin descripción")

    def comparar_mlfb():
        """Display the differences between the two entered MLFB codes."""
        mlfb1 = "".join(filter(str.isalnum, entry_mlfb1.get().strip().upper()))
        mlfb2 = "".join(filter(str.isalnum, entry_mlfb2.get().strip().upper()))

        if not mlfb1 or not mlfb2:
            messagebox.showerror(
                "Error", "Ambos campos MLFB deben estar llenos.", parent=ventana
            )
            return

        longitud_max = max(len(mlfb1), len(mlfb2))
        diferencias.delete(0, tk.END)

        for i in range(longitud_max):
            char1 = mlfb1[i] if i < len(mlfb1) else "(vacío)"
            char2 = mlfb2[i] if i < len(mlfb2) else "(vacío)"
            if char1 != char2:
                pos = i + 1
                significado1 = get_significado(pos, char1)
                significado2 = get_significado(pos, char2)
                diferencias.insert(tk.END, f"Posición {pos}:")
                diferencias.insert(tk.END, f"   - MLFB 1: {char1} → {significado1}")
                diferencias.insert(tk.END, f"   - MLFB 2: {char2} → {significado2}")
                diferencias.insert(tk.END, "")

    # Cargar y mostrar el logo de la aplicación si es posible
    try:
        img = Image.open(LOGO_PATH)
        img = img.resize((120, 120), Image.LANCZOS)
        logo = ImageTk.PhotoImage(img)
        label_logo = ttk.Label(ventana, image=logo)
        label_logo.image = logo  # Referencia para evitar garbage collection
        label_logo.pack(pady=10)
    except Exception as e:
        # No es crítico si la imagen falla: solo avisamos al usuario
        messagebox.showwarning(
            "Advertencia", f"No se pudo cargar el logo: {e}", parent=ventana
        )

    ttk.Label(
        ventana, text="Comparador de MLFB con Significado", font=("Arial", 16, "bold")
    ).pack()

    frame_entries = ttk.Frame(ventana)
    frame_entries.pack(pady=10)

    ttk.Label(frame_entries, text="MLFB 1:").grid(row=0, column=0, sticky="e", padx=5)
    entry_mlfb1 = ttk.Entry(frame_entries, width=30)
    entry_mlfb1.grid(row=0, column=1, padx=5)

    ttk.Label(frame_entries, text="MLFB 2:").grid(row=1, column=0, sticky="e", padx=5)
    entry_mlfb2 = ttk.Entry(frame_entries, width=30)
    entry_mlfb2.grid(row=1, column=1, padx=5)

    ttk.Button(ventana, text="Comparar MLFB", command=comparar_mlfb).pack(pady=10)

    # Listbox where the character-by-character comparison results are shown
    diferencias = tk.Listbox(
        ventana,
        selectmode=tk.SINGLE,
        height=15,
        font=("Arial", 10),
        bg=ABB_COLORS["bg"],
        fg=ABB_COLORS["fg"],
        selectbackground=ABB_COLORS["accent"],
        selectforeground=ABB_COLORS["fg"],
    )
    diferencias.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)