import openpyxl as pyxl
import glob
import tkinter as tk
import tkinter.font
from tkinter import font
import csv
import os
from tkinter import filedialog

bg_color = "#55868C"
button_color = "#7F636E"
text_color = "#000000"

largeur_fenetre = "600"
hauteur_fenetre = "500"
fenetreResolution = largeur_fenetre + "x" + hauteur_fenetre

policyHeader="Baskerville"
sizeHeader=20
weightHeader= "bold"

policyBody="Baskerville"
sizeBody=14
weightBody= "bold"

# Espacement
y_spacing = 40  # Espacement vertical entre les éléments
x_spacing = -10 # Espacement horizontal entre les élements

class ReleveAuto:
    tablfile = ""  ## valeur de stockage du lien du paquet de copie
    paquetfile = ""
    mois = ""
    annee = 0
    lignedebut = 0
    version = "Version 1.2"
    nbJours = 0
    
    fileFeuilleDeCompta = ""
    fileReleveDeCompte = ""

    def AffilierTableur(
        self, file
    ):  ## deff d'affiliation de correctionfile Appellée apres pression sur le bouton
        self.tablfile = file

    def fillJour(self, lignecourrante, worksheet, file):
        wbjour = pyxl.load_workbook(file)  ## wb du jour
        wsjour = wbjour.active  ## sheet de la fiche de caisse du jour
        valRJ = 0  ## valeur remboursement Jeu
        valRL = 0  ## valeur remboursement Loto
        for ligne in range(
            30
        ):  ## parcours des lignes dans la feuille de caisse puis check des valeurs des cases dans la colone B pour aller chercher les valeurs correspondantes et les mettre dans le tbleur final
            case = "B" + str(ligne + 1)  ## memoire pour savoir quelle ligne on se situ

            if wsjour[case].value == "Espèces":  ##test si correspond à espece
                casevaleur = "D" + str(
                    ligne + 1
                )  ## case dans laquellle on recupere le nombre de la fiche de caisse
                worksheet["B" + lignecourrante].value = wsjour[
                    casevaleur
                ].value  ## mise dans la valeur du tableau de compta

            if wsjour[case].value == "Chèque":
                casevaleur = "D" + str(ligne + 1)
                worksheet["C" + lignecourrante].value = wsjour[casevaleur].value

            if wsjour[case].value == "CB":
                casevaleur = "D" + str(ligne + 1)
                worksheet["D" + lignecourrante].value = wsjour[casevaleur].value

            if wsjour[case].value == "Gain tirage et sport":
                casevaleur = "D" + str(ligne + 1)
                valRL = wsjour[casevaleur].value + valRL
                print(" test gaintirage loto DANS TIRAGE :")
                print(wsjour[casevaleur].value)

            if wsjour[case].value == "Gain grattage":
                casevaleur = "D" + str(ligne + 1)
                valRJ = wsjour[casevaleur].value + valRJ

            if wsjour[case].value == "Especes POINT VERT":
                casevaleur = "D" + str(ligne + 1)
                casenb = "C" + str(ligne + 1)
                worksheet["I" + lignecourrante].value = wsjour[casevaleur].value
                worksheet["J" + lignecourrante].value = wsjour[casenb].value

            if wsjour[case].value == "Avoir":
                casevaleur = "D" + str(ligne + 1)
                worksheet["M" + lignecourrante].value = wsjour[casevaleur].value

            if wsjour[case].value == "REMB. JEU":
                casevaleur = "D" + str(ligne + 1)
                valRJ = valRJ + wsjour[casevaleur].value

            if wsjour[case].value == "REMB. LOTO":
                casevaleur = "D" + str(ligne + 1)
                print(" TEST RB LOTO AVANT ADD:")
                print(wsjour[casevaleur].value)
                valRL = valRL + wsjour[casevaleur].value
                print(" DANS TEST REMB LOTO " + str(valRL))

            if wsjour[case].value == "Mise en compte":
                casevaleur = "D" + str(ligne + 1)
                worksheet["O" + lignecourrante].value = wsjour[casevaleur].value

            if wsjour[case].value == "Paiement facture":
                casevaleur = "C" + str(ligne + 1)
                worksheet["P" + lignecourrante].value = wsjour[casevaleur].value

        print("FIN ADD:" + str(valRL))
        worksheet[
            "H" + lignecourrante
        ].value = valRL  ## remboursemenent loto car 2 cas et impossible de faire .value + .value
        worksheet["G" + lignecourrante].value = valRJ

    def AffilierDoss(
        self, file
    ):  ## def d'affiliation de paquetfile Apellée apres pression dur le bouton
        self.paquetfile = file

    def AffilierMA(self, mois, annee):  ## same pour nbquestion
        self.mois = mois
        self.annee = int(annee)

    def Execution(
        self,
    ):
        TablAnnee = self.tablfile  ## variable

        file_list = glob.glob(
            self.paquetfile + "/*.*"
        )  ## pour récuprer le ficher de toutes les feuilles de caisses
        print(file_list)
        wb = pyxl.load_workbook(TablAnnee)  ## recupération de la feuille de compta
        file_list.sort()

        for sheet in wb:  ## recherche de la bonne feuille de la bonne année
            if sheet.title == str(
                self.annee
            ):  ##si le nom de la feuille corrsepond à l'année
                ws = sheet  ##stockage de la sheet
        nombre = 1  ## pour stocker et creer la case de départ
        colone = "A"  ## pour faire les cases

        while nombre < 467:  ## recherche du mois dans le tbleur de compta
            casedate = colone + str(nombre)
            if ws[casedate].value == self.mois:
                self.lignedebut = (
                    nombre + 3
                )  ## + 3 car la case de juillet est 3 case au dessus de la premiere case à remplir
            nombre = nombre + 1

        if (
            self.mois == "JANVIER"
            or self.mois == "MARS"
            or self.mois == "MAI"
            or self.mois == "JUILLET"
            or self.mois == "AOUT"
            or self.mois == "OCTOBRE"
            or self.mois == "DECEMBRE"
        ):
            self.nbjours = 31
        elif self.mois == "FEVRIER":
            if self.annee % 4 == 0:
                self.nbjours = 29
            else:
                self.nbjours = 28
        else:
            self.nbjours = 30

        if self.mois == "JANVIER":
            for jour in range(
                self.nbjours
            ):  ## parcours des feuilles de caisses par jour
                if jour == 1:  ## si on est le premier janvier pas de caisse a faire
                    print()
                else:
                    self.fillJour(str(self.lignedebut + jour), ws, file_list[jour - 1])

        elif self.mois == "MAI":
            for jour in range(
                self.nbjours
            ):  ## parcours des feuilles de caisses par jour
                if jour == 0:
                    print()
                else:
                    self.fillJour(str(self.lignedebut + jour), ws, file_list[jour - 1])

        elif self.mois == "DECEMBRE":
            for jour in range(
                self.nbjours
            ):  ## parcours des feuilles de caisses par jour
                if jour == 24:
                    print()
                elif jour < 24:
                    self.fillJour(str(self.lignedebut + jour), ws, file_list[jour])
                elif jour > 24:
                    self.fillJour(str(self.lignedebut + jour), ws, file_list[jour])
        else:
            for jour in range(
                self.nbjours
            ):  ## parcours des feuilles de caisses par jour
                self.fillJour(str(self.lignedebut + jour), ws, file_list[jour])
        print("chaussure :) ")  ## j'ai retrouvé mes chaussures!!
        wb.save(TablAnnee)  ##save de la feuille de compta


class Application(tk.Tk):
    
    def __init__(self):
        tk.Tk.__init__(self)
        self.resizable(False, False)
        self.iconbitmap("Logo.ico")
        self.title("Relevé de compte automatique")
        self.configure(bg=bg_color)
        self.ReleveAuto = ReleveAuto()
        self.creer_widgets()
        
        # Récupération de la taille de l'écran
        largeur_ecran = self.winfo_screenwidth()
        hauteur_ecran = self.winfo_screenheight()
        x = (largeur_ecran - int(largeur_fenetre)) // 2
        y = (hauteur_ecran - int(hauteur_fenetre)) // 2
        y_spacing = 40  # Espacement vertical entre les éléments
        x_spacing = -10

        # Configuration de la position et de la taille de la fenêtre
        self.geometry(f"{int(largeur_fenetre)}x{int(hauteur_fenetre)}+{x}+{y}")
 

    def creer_widgets(self):
        canvas = tk.Canvas(self, bg=bg_color, width=largeur_fenetre, height=hauteur_fenetre)
        canvas.place(x=0, y=0)
        header_label = tk.Label(
            canvas,
            text="Relevé de compte automatique",
            font=font.Font(family=policyHeader, size=sizeHeader, weight=weightHeader),
            bg=bg_color,
            fg=text_color,
            )
        header_label.place(x=100, y=20)
        
        # Rectangle autour de l'entête
        header_rectangle = tk.Canvas(
            canvas, bg=bg_color, highlightbackground="black", highlightthickness=2
        )
        header_rectangle.place(x=95, y=15, width=430, height=48)
        header_label.lift(aboveThis=header_rectangle)
        
        # Bouton Donner feuille de compta
        feuille_btn = tk.Button(
            canvas,
            text="Donner feuille de compta",
            command=lambda : self.donnerFile(canvas,"feuille_btn"),
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=button_color,
            fg=text_color,
        )
        feuille_btn.place(x=50 + x_spacing, y=80 + y_spacing, width=250)
        
        releve_btn = tk.Button(
            canvas,
            text="Donner relevé de compte",
            command=lambda: self.donnerFile(canvas,"releve_btn"),
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=button_color,
            fg=text_color,
        )
        releve_btn.place(x=50 + x_spacing, y=130 + y_spacing, width=250)
        
        selection_label = tk.Label(
            canvas,
            text="Donner le mois et l'année:",
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=bg_color,
            fg=text_color,
        )
        selection_label.place(x=50 + x_spacing, y=180 + y_spacing)
        
        month_var = tk.StringVar(canvas)
        month_var.set("Mois")  # Ajoutez cette ligne pour définir "Mois" comme valeur par défaut
        month_dropdown = tk.OptionMenu(
            canvas,
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
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody), bg=button_color, fg=text_color, highlightthickness=0
        )

        month_dropdown.place(x=50 + x_spacing, y=220 + y_spacing, width=250)

        # Menu déroulant pour l'année
        year_var = tk.StringVar(canvas)
        year_var.set("Année")
        year_dropdown = tk.OptionMenu(canvas, year_var, "2023", "2024", "2025")
        year_dropdown.config(
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),bg=button_color, fg=text_color, highlightthickness=0
        )
        year_dropdown.place(x=320 + x_spacing, y=220 + y_spacing, width=150)


        # Affichage du mois et de l'année sélectionnés
        selected_label = tk.Label(
            canvas,
            text="Mois et année sélectionnés:",
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=bg_color,
            fg=text_color,
        )
        selected_label.place(x=50 + x_spacing, y=280 + y_spacing)
        
        selected_month_label = tk.Label(
            canvas, textvariable=month_var, font=font.Font(family=policyBody, size=sizeBody, weight=weightBody), bg=bg_color, fg=text_color
        )
        selected_month_label.place(x=50, y=320 + y_spacing)

        selected_year_label = tk.Label(
            canvas, textvariable=year_var, font=font.Font(family=policyBody, size=sizeBody, weight=weightBody), bg=bg_color, fg=text_color
        )
        selected_year_label.place(x=180, y=320 + y_spacing)
        
        # Bouton Lancer le programme
        launch_btn = tk.Button(
            canvas,
            text="Lancer le programme",
            command=lambda:self.ExecProgramme(),
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=button_color,
            fg=text_color,
        )
        launch_btn.place(x=20, y=450, width=250)
        
        # Bouton Fermer la page
        close_btn = tk.Button(
            canvas,
            text="Fermer",
            command=self.destroy,
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=button_color,
            fg=text_color,
        )
        close_btn.place(x=430, y=450, width=150)
        
        
    def afficher_error_message(parent, message):
    # Fonction pour gérer l'événement du clic sur le bouton "OK"
        def fermer_fenetre():
            error_fenetre.destroy()

        # Création de la fenêtre
        error_fenetre = tk.Toplevel(parent)
        error_fenetre.title("Message d'erreur")
        error_fenetre.iconbitmap("logo.ico")

        # Création d'un label pour afficher le message
        label_message = tk.Label(error_fenetre, text=message, font=font.Font(family=policyBody, size=sizeBody, weight=weightBody))
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

        # Bouton Soumettre le mois et l'annéereleve_btn
        # submit_btn = tk.Button(
        #     canvas,
        #     text="Soumettre",
        #     command=lambda : self.SoumettreMoisAnnee(),
        #     font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
        #     bg=button_color,
        #     fg=text_color,
        # )
        # submit_btn.place(x=490 + x_spacing, y=220 + y_spacing, width=100)




        # self.canva1 = tk.Canvas(
        #     self, width=hauteur_fenetre, height=largeur_fenetre
        # )  ## canva principal qui fait toute la fenetre
        # self.canva1.create_rectangle(
        #     20, 20, 380, 50, outline="black"
        # )  ## rectangle du haut, pour mettre titre
        # self.canva1.create_text(
        #     200, 35, text="Compta Automatique Maman"
        # )  ## Titre dans le rectangle
        # self.canva1.place(x=0, y=0)  ## placer le canva
        # self.canva1.create_oval(360, 60, 390, 90, fill="red")  ## voyant check tableur
        # self.canva1.create_oval(
        #     360, 105, 390, 135, fill="red"
        # )  ## voyant check Paquet de copie

        # self.canva2 = tk.Canvas(self, width=200, height=75)
        # self.canva2.create_text(
        #     100, 25, text="MOIS ANNEE", font=tkinter.font.Font(size=12)
        # )  ## Titre dans le rectangle
        # self.canva2.place(x=200, y=225)

        # OptionList = [
        #     "JANVIER",
        #     "FEVRIER",
        #     "MARS",
        #     "AVRIL",
        #     "MAI",
        #     "JUIN",
        #     "JUILLET",
        #     "AOUT",
        #     "SEPTEMBRE",
        #     "OCTOBRE",
        #     "NOVEMBRE",
        #     "DECEMBRE",
        # ]
        # self.variable = tk.StringVar(self)
        # self.variable.set(OptionList[0])
        # self.opt = tk.OptionMenu(self, self.variable, *OptionList)
        # self.opt.config(width=10, font=("Helvetica", 12))
        # self.opt.place(x=20, y=180)

        # OptionList2 = [2018, 2019, 2020, 2021, 2022, 2023]
        # self.variable2 = tk.StringVar(self)
        # self.variable2.set(OptionList2[4])
        # self.opt2 = tk.OptionMenu(self, self.variable2, *OptionList2)
        # self.opt2.config(width=10, font=("Helvetica", 12))
        # self.opt2.place(x=200, y=180)

        # self.bouton_askCompta = tk.Button(
        #     text="Donner le tableur",
        #     command=lambda: Compta.AffilierTableur(self.AskComptaFile()),
        # ).place(x=20, y=70)
        # self.bouton_askDoss = tk.Button(
        #     text="Donner le paquet de feuille de caisse",
        #     command=lambda: Compta.AffilierDoss(self.AskPackFile()),
        # ).place(x=20, y=110)
        # self.canva1.create_text(125, 160, text="Donner le mois et l'année")
        # self.bouton_MA = tk.Button(
        #     text="Soumettre le mois et l'année", command=lambda: self.AskMA()
        # ).place(x=20, y=220)
        # self.boutonComptaAuto = tk.Button(
        #     text="Lancer compta automatique", command=lambda: self.TestCompta()
        # ).place(x=20, y=310)
        # self.quitButon = tk.Button(
        #     self, text="Quitter", command=lambda: self.destroy()
        # ).place(x=330, y=360)

        # self.canva1.create_text(30, 390, text=self.Compta.version)

    def donnerFile(self,Cnvas,nomBouton):
        File = filedialog.askopenfilename(initialdir="Desktop/", title="Feuille de compta")
        print(File)
        
        if nomBouton == "feuille_btn":
            xcercle = int(largeur_fenetre) - 25
            ycercle = self.winfo_y() + 18
            ## ajout cercle deuxieme bouton
        elif nomBouton == "releve_btn":
            xcercle = int(largeur_fenetre) - 25
            ycercle = self.winfo_y() + 20
            
        rayon = 20    
        
        if ".xlsx" in File:
            if File != "":
                ##Create signal good format
                    Cnvas.create_oval(
                xcercle - rayon,
                ycercle - rayon,
                xcercle + rayon,
                ycercle + rayon,
                fill="green",
                )
                    if nomBouton == "feuille_btn":
                        
                        
                        print("feuille de compta donnée")
                        self.ReleveAuto.fileFeuilleDeCompta = File
                        
                    elif nomBouton == "releve_btn":
                        print("relevé de compte donné")
                        self.ReleveAuto.fileReleveDeCompte = File
                        
        else:
            Cnvas.create_oval(
                xcercle - rayon,
                ycercle - rayon,
                xcercle + rayon,
                ycercle + rayon,
                fill="red",
            )
            self.afficher_error_message(
                "ERREUR: Extension fichier non reconnue. \n Tu t'es surement trompée de fichier. \n Si le problème persiste, appelle Maxime.",
            )
        return ""


if __name__ == "__main__":
    app = Application()
    app.mainloop()

print("chaussure :) ")  ##j'ai retrouvé mes chaussures !
