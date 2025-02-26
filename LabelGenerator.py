import tkinter as tk
from tkinter import messagebox, ttk
from fpdf import FPDF
import math
import os

# Constantes pour la taille d'une feuille A4 en mm
A4_WIDTH = 210
A4_HEIGHT = 297
MARGE_FEUILLE = 20  # Marge pour éviter que l'impression ne coupe la première ligne ou colonne

class EtiquetteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Générateur d'étiquettes")
        self.root.geometry("500x350")
        self.root.configure(bg="#f0f0f0")

        # Appliquer le thème 'clam' pour une interface plus moderne
        style = ttk.Style()
        style.theme_use("vista")
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.configure("TEntry", font=("Segoe UI", 10), padding=5)
        style.configure("TLabel", font=("Segoe UI", 10), background="#f0f0f0", padding=5)
        style.configure("TCombobox", font=("Segoe UI", 10), padding=5)

        # Interface modernisée
        frame = ttk.Frame(root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))
        
        ttk.Label(frame, text="Largeur de l'étiquette (mm):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.largeur_entry = ttk.Entry(frame)
        self.largeur_entry.grid(row=0, column=1, pady=5, padx=10)
        
        ttk.Label(frame, text="Hauteur de l'étiquette (mm):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.hauteur_entry = ttk.Entry(frame)
        self.hauteur_entry.grid(row=1, column=1, pady=5, padx=10)
        
        ttk.Label(frame, text="Marge entre étiquettes (mm):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.marge_entry = ttk.Entry(frame)
        self.marge_entry.grid(row=2, column=1, pady=5, padx=10)
        
        ttk.Label(frame, text="Nombre d'étiquettes:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.nb_etiquettes_entry = ttk.Entry(frame)
        self.nb_etiquettes_entry.grid(row=3, column=1, pady=5, padx=10)

        ttk.Label(frame, text="Taille de la police:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.font_size_entry = ttk.Entry(frame)
        self.font_size_entry.grid(row=4, column=1, pady=5, padx=10)
        self.font_size_entry.insert(0, "10")
        
        ttk.Label(frame, text="Police d'écriture:").grid(row=5, column=0, sticky=tk.W, pady=5)
        available_fonts = ["Arial", "Courier", "Times"]  # Liste simplifiée pour le PDF
        self.font_var = tk.StringVar()
        self.font_var.set("Arial")
        self.font_menu = ttk.Combobox(frame, textvariable=self.font_var, values=available_fonts, state="readonly")
        self.font_menu.grid(row=5, column=1, pady=5, padx=10)
        
        ttk.Button(frame, text="Personnaliser & Générer PDF", command=self.personnaliser_etiquettes).grid(row=6, columnspan=2, pady=15)
        
        
    def personnaliser_etiquettes(self):
        try:
            largeur = float(self.largeur_entry.get())
            hauteur = float(self.hauteur_entry.get())
            marge = float(self.marge_entry.get())
            nb_etiquettes = int(self.nb_etiquettes_entry.get())
            font_size = int(self.font_size_entry.get())
            font_name = self.font_var.get()
        except ValueError:
            messagebox.showerror("Erreur", "Veuillez entrer des valeurs valides.")
            return
        
        if largeur <= 0 or hauteur <= 0 or nb_etiquettes <= 0 or marge < 0 or font_size <= 0:
            messagebox.showerror("Erreur", "Veuillez entrer des valeurs valides.")
            return
        
        self.custom_window = tk.Toplevel(self.root)
        self.custom_window.title("Personnalisation des étiquettes")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.7)
        window_height = int(screen_height * 0.7)
        self.custom_window.geometry(f"{window_width}x{window_height}")
        
        frame = ttk.Frame(self.custom_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.text_entries = []
        nb_colonnes = min(5, nb_etiquettes)  # Limite le nombre de colonnes affichées
        for i in range(nb_etiquettes):
            entry = ttk.Entry(scrollable_frame, font=("Segoe UI", 12), width=20)
            entry.grid(row=i // nb_colonnes, column=i % nb_colonnes, padx=5, pady=5, ipadx=5, ipady=5)
            self.text_entries.append(entry)
        
        ttk.Button(scrollable_frame, text="Valider", command=lambda: self.valider_textes(largeur, hauteur, marge, nb_etiquettes, font_name, font_size)).grid(row=(nb_etiquettes // nb_colonnes) + 1, columnspan=nb_colonnes, pady=10)

    def valider_textes(self, largeur, hauteur, marge, nb_etiquettes, font_name, font_size):
        self.textes = [entry.get() for entry in self.text_entries]
        self.custom_window.destroy()
        self.generer_pdf(largeur, hauteur, marge, nb_etiquettes, font_name, font_size)

    def generer_pdf(self, largeur, hauteur, marge, nb_etiquettes, font_name, font_size):
        nb_colonnes = math.floor((A4_WIDTH - 2 * MARGE_FEUILLE - marge) / (largeur + marge))
        nb_lignes = math.floor((A4_HEIGHT - 2 * MARGE_FEUILLE - marge) / (hauteur + marge))
        nb_par_page = nb_colonnes * nb_lignes
        
        if nb_par_page == 0:
            messagebox.showerror("Erreur", "Les dimensions sont trop grandes pour une feuille A4.")
            return
        
        nb_pages = math.ceil(nb_etiquettes / nb_par_page)
        
        pdf = FPDF('P', 'mm', 'A4')
        pdf.set_margins(MARGE_FEUILLE, MARGE_FEUILLE, MARGE_FEUILLE)
        pdf.set_auto_page_break(auto=True, margin=MARGE_FEUILLE)
        
        etiquette_index = 0
        for _ in range(nb_pages):
            pdf.add_page()
            pdf.set_font(font_name, size=font_size)
            total_largeur_utilisee = nb_colonnes * (largeur + marge) - marge
            total_hauteur_utilisee = nb_lignes * (hauteur + marge) - marge
            offset_x = (A4_WIDTH - total_largeur_utilisee) / 2
            offset_y = (A4_HEIGHT - total_hauteur_utilisee) / 2
            
        for i in range(nb_lignes):
            for j in range(nb_colonnes):
                if etiquette_index >= nb_etiquettes:
                    break
                x = j * (largeur + marge) + offset_x
                y = i * (hauteur + marge) + offset_y
                pdf.rect(x, y, largeur, hauteur)

                # Récupération du texte à afficher
                texte = self.textes[etiquette_index] if etiquette_index < len(self.textes) else f'Étiquette {etiquette_index + 1}'

                # Mesure de la largeur et de la hauteur du texte
                text_width = pdf.get_string_width(texte)
                line_height = font_size * 0.8  # Ajustement pour la hauteur de ligne

                # Calcul des coordonnées pour centrer le texte
                text_x = x + (largeur - text_width) / 2
                text_y = y + (hauteur - line_height) / 2

                # Positionnement du texte centré
                pdf.set_xy(text_x, text_y)
                pdf.cell(text_width, line_height, texte, border=0, align='C')

                etiquette_index += 1

        
        output_file = "etiquettes.pdf"
        counter = 1
        while os.path.exists(output_file):
            output_file = f"etiquettes_{counter}.pdf"
            counter += 1
        
        pdf.output(output_file)
        messagebox.showinfo("Succès", f"PDF généré avec succès sous le nom '{output_file}'.")


if __name__ == "__main__":
    root = tk.Tk()
    app = EtiquetteApp(root)
    root.mainloop()