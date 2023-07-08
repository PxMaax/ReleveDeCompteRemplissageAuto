import tkinter as tk

## add test

def DonnerFeuilleDeCompta():
    # Logique de la fonction DonnerFeuilleDeCompta
    print("Donner feuille de compta")

def DonnerReleveDeCompte():
    # Logique de la fonction DonnerReleveDeCompte
    print("Donner relevé de compte")

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
window.geometry("600x500")  # Taille fixe de la fenêtre
window.configure(bg="#B2EBB7")  # Couleur de fond de la fenêtre

# Couleurs
bg_color = "#B2EBB7"
button_color = "#EDAFB8"
text_color = "#000000"

# Police
header_font = ("Baskerville", 20, "bold")
label_font = ("Baskerville", 14)
button_font = ("Baskerville", 14, "bold")

# Entête
header_label = tk.Label(window, text="Relevé de compte automatique", font=header_font, bg=bg_color, fg=text_color)
header_label.place(x=50, y=20)

# Bouton Donner feuille de compta
feuille_btn = tk.Button(window, text="Donner feuille de compta", command=DonnerFeuilleDeCompta, font=button_font, bg=button_color, fg=text_color)
feuille_btn.place(x=50, y=80, width=250)

# Carré vide pour la validation de Donner feuille de compta
feuille_validation = tk.Label(window, text="V", font=("Baskerville", 24, "bold"), bg=bg_color, fg="green")
feuille_validation.place(x=320, y=82)

# Bouton Donner relevé de compte
releve_btn = tk.Button(window, text="Donner relevé de compte", command=DonnerReleveDeCompte, font=button_font, bg=button_color, fg=text_color)
releve_btn.place(x=50, y=130, width=250)

# Carré vide pour la validation de Donner relevé de compte
releve_validation = tk.Label(window, text="V", font=("Baskerville", 24, "bold"), bg=bg_color, fg="green")
releve_validation.place(x=320, y=132)

# Entrée pour le mois et l'année
selection_label = tk.Label(window, text="Donner le mois et l'année:", font=label_font, bg=bg_color, fg=text_color)
selection_label.place(x=50, y=180)

# Menu déroulant pour le mois
month_var = tk.StringVar(window)
month_dropdown = tk.OptionMenu(window, month_var, "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre")
month_dropdown.config(font=label_font)
month_dropdown.place(x=50, y=220, width=250)

# Menu déroulant pour l'année
year_var = tk.StringVar(window)
year_dropdown = tk.OptionMenu(window, year_var, "2023", "2024", "2025")
year_dropdown.config(font=label_font)
year_dropdown.place(x=320, y=220, width=150)

# Bouton Soumettre le mois et l'année
submit_btn = tk.Button(window, text="Soumettre", command=SoumettreMoisAnnee, font=button_font, bg=button_color, fg=text_color)
submit_btn.place(x=490, y=220, width=100)

# Affichage du mois et de l'année sélectionnés
selected_label = tk.Label(window, text="Mois et année sélectionnés:", font=label_font, bg=bg_color, fg=text_color)
selected_label.place(x=50, y=280)

selected_month_label = tk.Label(window, textvariable=month_var, font=label_font, bg=bg_color, fg=text_color)
selected_month_label.place(x=50, y=320)

selected_year_label = tk.Label(window, textvariable=year_var, font=label_font, bg=bg_color, fg=text_color)
selected_year_label.place(x=180, y=320)

# Bouton Lancer le programme
launch_btn = tk.Button(window, text="Lancer le programme", command=LancerProgramme, font=button_font, bg=button_color, fg=text_color)
launch_btn.place(x=50, y=380, width=250)

# Bouton Fermer la page
close_btn = tk.Button(window, text="Fermer", command=window.quit, font=button_font, bg=button_color, fg=text_color)
close_btn.place(x=320, y=380, width=150)

# Lancement de la boucle principale
window.mainloop()
