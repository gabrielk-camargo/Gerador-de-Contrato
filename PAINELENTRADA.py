import tkinter as tk
import subprocess

def run_gerador_contrato():
    # Executa o script "GERADORCONTRATO.py" (Prazo Indeterminado)
    subprocess.Popen(["python", "GERADORCONTRATO.py"])
    root.destroy()

def prazo_determinado():
    # Executa o script "CONTRATODETERMINADO.py" (Prazo Determinado)
    subprocess.Popen(["python", "CONTRATODETERMINADO.py"])
    root.destroy()

root = tk.Tk()
root.iconbitmap("icone.ico")
root.title("GERADOR CONTRATOS IZI CAR")
root.geometry("500x400")

label = tk.Label(root, text="GERADOR CONTRATOS IZI CAR")
label.pack(pady=20)

btn_determinado = tk.Button(root, text="PRAZO DETERMINADO", command=prazo_determinado)
btn_determinado.pack(pady=10)

btn_indeterminado = tk.Button(root, text="PRAZO INDETERMINADO", command=run_gerador_contrato)
btn_indeterminado.pack(pady=10)

root.mainloop()
