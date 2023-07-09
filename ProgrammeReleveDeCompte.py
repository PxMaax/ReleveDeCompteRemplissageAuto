import tkinter as tk
from tkinter import font
from tkinter import ttk
import openpyxl as pyxl
import glob
import csv
import os
from tkinter import filedialog


## Déclaration de variable

largeur_fenetre = "600"
hauteur_fenetre = "500"
windowResolution = largeur_fenetre + "x" + hauteur_fenetre

fileFeuilleDeCompta = ""
fileReleveDeCompte = ""


def donnerFile(nomBouton):
    File = filedialog.askopenfilename(initialdir="Desktop/", title="Feuille de compta")
    print(File)
    if ".xlsx" in File:
        if File != "":
                afficher_cercle(nomBouton, "true")
                assign_file(nomBouton,File)
                return File
    else:
        afficher_cercle(nomBouton, "false")
        afficher_error_message(
            window,
            "ERREUR: Extension fichier non reconnue. \n Tu t'es surement trompé de fichier. \n Si le problème persite, appelle Maxime",
        )
        return ""

def assign_file(nomBouton,file):
    if nomBouton == "feuille_btn":
        fileFeuilleDeCompta = file
    elif nomBouton == "releve_btn":
        fileReleveDeCompte = file
    return "true"
        

def afficher_cercle(nomFunction, valeur):
        ## ajout cercle premier bouton
    if nomFunction == "feuille_btn": 
        xcercle = int(largeur_fenetre) - 25
        ycercle = feuille_btn.winfo_y() + 18
        ## ajout cercle deuxieme bouton
    elif nomFunction == "releve_btn":
        xcercle = int(largeur_fenetre) - 25
        ycercle = releve_btn.winfo_y() + 20

    rayon = 20

    if valeur == "false":
        canvas.create_oval(xcercle - rayon, ycercle - rayon, xcercle + rayon, ycercle + rayon, fill="red")
    elif valeur == "true":
        canvas.create_oval(xcercle - rayon, ycercle - rayon, xcercle + rayon, ycercle + rayon, fill="green")


def afficher_error_message(parent, message):
    # Fonction pour gérer l'événement du clic sur le bouton "OK"
    def fermer_fenetre():
        error_fenetre.destroy()

    # Création de la fenêtre
    error_fenetre = tk.Toplevel(parent)
    error_fenetre.title("Message d'erreur")
    error_fenetre.iconbitmap("logo.ico")

    # Création d'un label pour afficher le message
    label_message = tk.Label(error_fenetre, text=message, font=label_font)
    label_message.pack(padx=20, pady=20)

    # Création d'un bouton "OK"
    bouton_ok = tk.Button(
        error_fenetre, text="OK", command=fermer_fenetre, width=10, height=2
    )
    bouton_ok.pack(pady=20)

    # Calcul de la taille de la fenêtre en fonction de la longueur du message
    largeur_error_fenetre = max(350, label_message.winfo_reqwidth() + 40)
    hauteur_error_fenetre = 180

    # Récupération de la taille de la fenêtre parente
    parent.update_idletasks()  # Actualisation des tâches du parent
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()

    # Calcul de la position de la fenêtre pour la centrer
    xerrorwindow = parent.winfo_rootx() + (parent_width - largeur_error_fenetre) // 2
    yerrorwindow = parent.winfo_rooty() + (parent_height - hauteur_error_fenetre) // 2

    # Configuration de la position et de la taille de la fenêtre
    error_fenetre.geometry(
        f"{largeur_error_fenetre}x{hauteur_error_fenetre}+{xerrorwindow}+{yerrorwindow}"
    )

    # Lancement de la boucle principale de la fenêtre
    error_fenetre.mainloop()


## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##
## FONCTION EN RAPPORT AVEC L'INTERFACE GRAPHIQUE ##

# Couleurs
bg_color = "#55868C"
button_color = "#7F636E"
text_color = "#000000"



def SoumettreMoisAnnee():
    # Logique de la fonction SoumettreMoisAnnee
    selected_month = month_var.get()
    selected_year = year_var.get()
    print("Mois sélectionné :", selected_month)
    print("Année sélectionnée :", selected_year)


def LancerProgramme():
    # Logique de la fonction LancerProgramme
    print("Lancer le programme")


# Création de la fenêtre principale
window = tk.Tk()
window.title("Relevé de compte automatique")
window.geometry(windowResolution)  # Taille fixe de la fenêtre
window.configure(bg=bg_color)  # Couleur de fond de la fenêtre
window.resizable(False, False)  # Désactiver le redimensionnement de la fenêtre
window.iconbitmap("logo.ico")

# Récupération de la taille de l'écran
largeur_ecran = window.winfo_screenwidth()
hauteur_ecran = window.winfo_screenheight()

x = (largeur_ecran - int(largeur_fenetre)) // 2
y = (hauteur_ecran - int(hauteur_fenetre)) // 2

# Configuration de la position et de la taille de la fenêtre
window.geometry(f"{int(largeur_fenetre)}x{int(hauteur_fenetre)}+{x}+{y}")

# Police
header_font = font.Font(family="Baskerville", size=20, weight="bold")
label_font = font.Font(family="Baskerville", size=14, weight="bold")
button_font = font.Font(family="Baskerville", size=14, weight="bold")

# Espacement
y_spacing = 40  # Espacement vertical entre les éléments
x_spacing = -10


# Entête
header_label = tk.Label(
    window,
    text="Relevé de compte automatique",
    font=header_font,
    bg=bg_color,
    fg=text_color,
)
header_label.place(x=100, y=20)

# Canvas pour afficher les cercles
canvas = tk.Canvas(window, bg="#55868C", width=largeur_fenetre, height=hauteur_fenetre)
canvas.place(x=0, y=0)

# Rectangle autour de l'entête
header_rectangle = tk.Canvas(
    window, bg=bg_color, highlightbackground="black", highlightthickness=2
)
header_rectangle.place(x=95, y=15, width=430, height=48)
header_label.lift(aboveThis=header_rectangle)

# Bouton Donner feuille de compta
feuille_btn = tk.Button(
    window,
    text="Donner feuille de compta",
    command=lambda :donnerFile("feuille_btn"),
    font=button_font,
    bg=button_color,
    fg=text_color,
)
feuille_btn.place(x=50 + x_spacing, y=80 + y_spacing, width=250)

# Bouton Donner relevé de compte
releve_btn = tk.Button(
    window,
    text="Donner relevé de compte",
    command= lambda : donnerFile("releve_btn"),
    font=button_font,
    bg=button_color,
    fg=text_color,
)
releve_btn.place(x=50 + x_spacing, y=130 + y_spacing, width=250)
# Entrée pour le mois et l'année
selection_label = tk.Label(
    window,
    text="Donner le mois et l'année:",
    font=label_font,
    bg=bg_color,
    fg=text_color,
)
selection_label.place(x=50 + x_spacing, y=180 + y_spacing)

# Menu déroulant pour le mois
month_var = tk.StringVar(window)
month_var.set("Mois")  # Ajoutez cette ligne pour définir "Mois" comme valeur par défaut
month_dropdown = tk.OptionMenu(
    window,
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
year_var = tk.StringVar(window)
year_var.set("Année")
year_dropdown = tk.OptionMenu(window, year_var, "2023", "2024", "2025")
year_dropdown.config(
    font=button_font, bg=button_color, fg=text_color, highlightthickness=0
)
year_dropdown.place(x=320 + x_spacing, y=220 + y_spacing, width=150)

# Bouton Soumettre le mois et l'annéereleve_btn
submit_btn = tk.Button(
    window,
    text="Soumettre",
    command=SoumettreMoisAnnee,
    font=button_font,
    bg=button_color,
    fg=text_color,
)
submit_btn.place(x=490 + x_spacing, y=220 + y_spacing, width=100)

# Affichage du mois et de l'année sélectionnés
selected_label = tk.Label(
    window,
    text="Mois et année sélectionnés:",
    font=label_font,
    bg=bg_color,
    fg=text_color,
)
selected_label.place(x=50 + x_spacing, y=280 + y_spacing)

selected_month_label = tk.Label(
    window, textvariable=month_var, font=label_font, bg=bg_color, fg=text_color
)
selected_month_label.place(x=50, y=320 + y_spacing)

selected_year_label = tk.Label(
    window, textvariable=year_var, font=label_font, bg=bg_color, fg=text_color
)
selected_year_label.place(x=180, y=320 + y_spacing)

# Bouton Lancer le programme
launch_btn = tk.Button(
    window,
    text="Lancer le programme",
    command=LancerProgramme,
    font=button_font,
    bg=button_color,
    fg=text_color,
)
launch_btn.place(x=20, y=450, width=250)

# Bouton Fermer la page
close_btn = tk.Button(
    window,
    text="Fermer",
    command=window.quit,
    font=button_font,
    bg=button_color,
    fg=text_color,
)
close_btn.place(x=430, y=450, width=150)


# Lancement de la boucle principale
window.mainloop()
