import tkinter as tk
from tkinter import font
from tkinter import ttk
import openpyxl as pyxl
import glob
import csv
import os
from tkinter import filedialog


class Application(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        bg_color = "#55868C"
        button_color = "#7F636E"
        text_color = "#000000"

        # Création de la fenêtre principale
        self.title("Relevé de compte automatique")
        self.geometry(self.windowResolution)  # Taille fixe de la fenêtre
        self.configure(bg=bg_color)  # Couleur de fond de la fenêtre
        self.resizable(False, False)  # Désactiver le redimensionnement de la fenêtre
        self.iconbitmap("logo.ico")

        # Récupération de la taille de l'écran
        largeur_ecran = self.winfo_screenwidth()
        hauteur_ecran = self.winfo_screenheight()

        x = (largeur_ecran - int(self.largeur_fenetre)) // 2
        y = (hauteur_ecran - int(self.largeur_fenetre)) // 2

        # Configuration de la position et de la taille de la fenêtre
        self.geometry(
            f"{int(self.largeur_fenetre)}x{int(self.largeur_fenetre)}+{x}+{y}"
        )

        # Police
        header_font = font.Font(family="Baskerville", size=20, weight="bold")
        label_font = font.Font(family="Baskerville", size=14, weight="bold")
        button_font = font.Font(family="Baskerville", size=14, weight="bold")

        # Espacement
        y_spacing = 40  # Espacement vertical entre les éléments
        x_spacing = -10

        # Entête
        header_label = tk.Label(
            self,
            text="Relevé de compte automatique",
            font=header_font,
            bg=bg_color,
            fg=text_color,
        )
        header_label.place(x=100, y=20)

        # Canvas pour afficher les cercles
        canvas = tk.Canvas(
            self,
            bg="#55868C",
            width=self.largeur_fenetre,
            height=self.largeur_fenetre,
        )
        canvas.place(x=0, y=0)

        # Rectangle autour de l'entête
        header_rectangle = tk.Canvas(
            self, bg=bg_color, highlightbackground="black", highlightthickness=2
        )
        header_rectangle.place(x=95, y=15, width=430, height=48)
        header_label.lift(aboveThis=header_rectangle)

        # Bouton Donner feuille de compta
        feuille_btn = tk.Button(
            self,
            text="Donner feuille de compta",
            # command=lambda: donnerFile("feuille_btn"),
            font=button_font,
            bg=button_color,
            fg=text_color,
        )
        feuille_btn.place(x=50 + x_spacing, y=80 + y_spacing, width=250)

        # Bouton Donner relevé de compte
        releve_btn = tk.Button(
            self,
            text="Donner relevé de compte",
            # command=lambda: donnerFile("releve_btn"),
            font=button_font,
            bg=button_color,
            fg=text_color,
        )
        releve_btn.place(x=50 + x_spacing, y=130 + y_spacing, width=250)
        # Entrée pour le mois et l'année
        selection_label = tk.Label(
            self,
            text="Donner le mois et l'année:",
            font=label_font,
            bg=bg_color,
            fg=text_color,
        )
        selection_label.place(x=50 + x_spacing, y=180 + y_spacing)

        # Menu déroulant pour le mois
        month_var = tk.StringVar(self)
        month_var.set(
            "Mois"
        )  # Ajoutez cette ligne pour définir "Mois" comme valeur par défaut
        month_dropdown = tk.OptionMenu(
            self,
            month_var,
            "Mois",  # Ajoutez "Mois" comme première option dans la liste
            "Janvier",
            "Février",
            "Mars",
            "Avril",
            "Mai",
            "Juin",
            "Juillet",
            "Août",
            "Septembre",
            "Octobre",
            "Novembre",
            "Décembre",
        )

        month_dropdown.config(
            font=button_font, bg=button_color, fg=text_color, highlightthickness=0
        )

        month_dropdown.place(x=50 + x_spacing, y=220 + y_spacing, width=250)

        # Menu déroulant pour l'année
        year_var = tk.StringVar(self)
        year_var.set("Année")
        year_dropdown = tk.OptionMenu(self, year_var, "2023", "2024", "2025")
        year_dropdown.config(
            font=button_font, bg=button_color, fg=text_color, highlightthickness=0
        )
        year_dropdown.place(x=320 + x_spacing, y=220 + y_spacing, width=150)

        # Bouton Soumettre le mois et l'annéereleve_btn
        submit_btn = tk.Button(
            self,
            text="Soumettre",
            # command=SoumettreMoisAnnee,
            font=button_font,
            bg=button_color,
            fg=text_color,
        )
        submit_btn.place(x=490 + x_spacing, y=220 + y_spacing, width=100)

        # Affichage du mois et de l'année sélectionnés
        selected_label = tk.Label(
            self,
            text="Mois et année sélectionnés:",
            font=label_font,
            bg=bg_color,
            fg=text_color,
        )
        selected_label.place(x=50 + x_spacing, y=280 + y_spacing)

        selected_month_label = tk.Label(
            self, textvariable=month_var, font=label_font, bg=bg_color, fg=text_color
        )
        selected_month_label.place(x=50, y=320 + y_spacing)

        selected_year_label = tk.Label(
            self, textvariable=year_var, font=label_font, bg=bg_color, fg=text_color
        )
        selected_year_label.place(x=180, y=320 + y_spacing)

        # Bouton Lancer le programme
        launch_btn = tk.Button(
            self,
            text="Lancer le programme",
            # command=ExecProgramme,
            font=button_font,
            bg=button_color,
            fg=text_color,
        )
        launch_btn.place(x=20, y=450, width=250)

        # Bouton Fermer la page
        close_btn = tk.Button(
            self,
            text="Fermer",
            command=self.quit,
            font=button_font,
            bg=button_color,
            fg=text_color,
        )
        close_btn.place(x=430, y=450, width=150)

        self.largeur_fenetre = "600"
        self.hauteur_fenetre = "500"

        self.windowResolution = self.largeur_fenetre + "x" + self.hauteur_fenetre
        self.fileFeuilleDeCompta = ""
        self.fileReleveDeCompte = ""

        self.selected_month = ""
        self.selected_year = ""
        self.stratingLine = 0

    ## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
    ## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
    ## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
    ## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
    ## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
    ## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##

    # Couleurs


if __name__ == "__main__":
    app = Application()
    app.mainloop()
