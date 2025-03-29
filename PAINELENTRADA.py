import tkinter as tk
import subprocess
import customtkinter as ctk
from tkinter import messagebox

# Set the appearance mode to light
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")  # You can choose other themes, but blue is fine with light mode

def run_script(script_name):
    """Executes a given Python script and handles potential errors."""
    try:
        subprocess.Popen(["python", script_name])
    except FileNotFoundError:
        messagebox.showerror("Error", f"Script '{script_name}' not found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while running '{script_name}': {e}")
    root.destroy()

def run_gerador_contrato():
    run_script("GERADORCONTRATO.py")

def prazo_determinado():
    run_script("CONTRATODETERMINADO.py")

def termo_divida():
    run_script("TERMO_DIVIDA.py")

def investidor():
    run_script("INVESTIDOR.py")

def update_layout():
    """Updates the layout of the buttons based on window width."""
    width = root.winfo_width()
    if width >= 700:  # Adjusted threshold
        # Vertical layout
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=0)
        label.grid(row=0, column=0, columnspan=1, padx=20, pady=20, sticky="nsew")
        btn_determinado.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        btn_indeterminado.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        btn_termo_divida.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        btn_investidor.grid(row=4, column=0, padx=20, pady=(10, 20), sticky="ew")
    else:
        # Two-column layout
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        label.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")
        btn_determinado.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        btn_indeterminado.grid(row=1, column=1, padx=20, pady=(0, 10), sticky="ew")
        btn_termo_divida.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        btn_investidor.grid(row=2, column=1, padx=20, pady=10, sticky="ew")

root = ctk.CTk()
try:
    root.iconbitmap("icons/icone.ico")
except tk.TclError:
    print("Warning: Icon file 'icons/icone.ico' not found.")
root.title("IZI CAR - Gerador de Contratos Profissional")
root.geometry("600x450")  # Increased initial size

# Define colors for the menu bar based on the light theme
menu_bg_color = "#f0f0f0"  # A light gray color (standard Tkinter light background)
menu_fg_color = "black"

# Menu Bar
menubar = tk.Menu(root, background=menu_bg_color, foreground=menu_fg_color)
filemenu = tk.Menu(menubar, tearoff=0, background=menu_bg_color, foreground=menu_fg_color)
filemenu.add_command(label="Sair", command=root.quit)
menubar.add_cascade(label="Arquivo", menu=filemenu)

helpmenu = tk.Menu(menubar, tearoff=0, background=menu_bg_color, foreground=menu_fg_color)
helpmenu.add_command(label="Sobre", command=lambda: messagebox.showinfo("Sobre", "Gerador de Contratos IZI CAR - Versão 1.0"))
menubar.add_cascade(label="Ajuda", menu=helpmenu)

root.config(menu=menubar)

# Main Frame for content organization
main_frame = ctk.CTkFrame(root, corner_radius=10)
main_frame.pack(padx=20, pady=20, fill="both", expand=True)

# Configure grid layout for the main frame
main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_columnconfigure(1, weight=1)
main_frame.grid_rowconfigure(0, weight=1)
main_frame.grid_rowconfigure(1, weight=1)
main_frame.grid_rowconfigure(2, weight=1)
main_frame.grid_rowconfigure(3, weight=1)

label = ctk.CTkLabel(main_frame, text="Selecione a Ação Desejada", font=ctk.CTkFont(size=20, weight="bold"))

btn_determinado = ctk.CTkButton(main_frame, text="Gerar Contrato Prazo Determinado", command=prazo_determinado)
btn_indeterminado = ctk.CTkButton(main_frame, text="Gerar Contrato Prazo Indeterminado", command=run_gerador_contrato)
btn_termo_divida = ctk.CTkButton(main_frame, text="Gerar Termo de Dívida", command=termo_divida)
btn_investidor = ctk.CTkButton(main_frame, text="Gerar Contrato de Investidor", command=investidor)

# Initial layout
update_layout()

# Bind the resize event to update the layout dynamically
root.bind("<Configure>", lambda event: update_layout())

# Status Bar
statusbar = tk.Label(root, text="Status: Normal", bd=1, relief=tk.SUNKEN, anchor=tk.W, fg="green", font=("TkDefaultFont", 10, "bold"))
statusbar.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()