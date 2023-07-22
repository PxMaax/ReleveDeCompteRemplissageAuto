import openpyxl as pyxl
import glob
import tkinter as tk
import tkinter.font
from tkinter import font
import csv
import os
from tkinter import filedialog
import re
from datetime import date
from erreur_class import ErreurExcel

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

# Obtenir la date d'aujourd'hui
aujourd_hui = date.today()

# Convertir la date en chaîne de caractères
date_str = aujourd_hui.strftime("%d/%m/%Y")  # Format: jour/mois/année

# Espacement
y_spacing = 40  # Espacement vertical entre les éléments
x_spacing = -10 # Espacement horizontal entre les élements

class ReleveAuto:
    
    mois = ""
    annee = 0
    lignedebut = 0
    version = "Version 1"
    nbJours = 0
    
    ligne_depart_releve_de_compte = 0
    
    fileFeuilleDeCompta = ""
    fileReleveDeCompte = ""
    
    ligne_depart_releve = 11
    colone_depart_releve = "B"
    
    tableau_de_compta_wb = None
    releve_de_compte_wb  = None
    
    tableau_de_compta_ws = None
    releve_de_compte_ws = None
    
    excel_des_problemes = pyxl.Workbook()
    sheet_des_problemes = excel_des_problemes.active
    sheet_des_problemes['A1'] = "Case"
    sheet_des_problemes['B1'] = "Erreur"
    
    tableau_erreur = []
    
    case_iteration = ""
    
    
    def fill_fichier_error (self):
        for index, (valeur_1, valeur_2) in enumerate(self.tableau_erreur, start=2):
            self.sheet_des_problemes.cell(row=index, column=1, value=valeur_1)
            self.sheet_des_problemes.cell(row=index, column=2, value=valeur_2)
        
    def cocher_releve_de_compte(self,ligne):
        self.releve_de_compte_ws["E" + ligne] = "X"
        return 
    
    def verif_CB(self,valeur_SC,valeur_AvC ):
        
        """
            verif_CB : Check les valeurs trouvées dans le relevé de compte match avec les fiches de caisse

            :param valeur_SC : valeur carte Sans contact
            :param valeur_AvC : valeur carte Avc Contact
            :return : Vrai si match
        """ 
        if (self.tableau_de_compta_ws["D"+self.lignedebut] != (valeur_SC + valeur_AvC) ):
            raise ErreurExcel(self.case_iteration, "la date n'a pas été trouvée")
        
        return True
    
    def getValeurCredit (self, current_cell):
        
        """
            getValeurCredit : Function trouve la valeur de crédit

            :param current_cell : case de la valeur qu'on cherche
            :return: Credit lié à la case en cours
        """ 
    
        return self.releve_de_compte_ws["D" + current_cell.row ].value
        
    def extract_match_date_from_string(self,value_case_releve):
        # Utiliser une expression régulière pour rechercher la date au format "DD/MM"
        date_pattern = r'\b\d{2}/\d{2}\b'
        match = re.search(date_pattern, value_case_releve)
        date_string = match.group()
        if date_string:
            return date_string
        else:
            raise ErreurExcel(self.case_iteration, "la date n'a pas été trouvée")
 
    def trouver_case_carte_meme_date(self, ligne_current_cell, target_date):
        
        """
            trouver_case_carte_meme_date : Function qui permet de trouver l'autre carte bleue pour compléter la journée
            
            La fonction check dans les +15/-15 cases autour si elle contient le code des cartes sans contact
            
            :param ligne_current_cell : la ligne de la case courrante
            :param target_date : la date à trouver dans les cases
            :return: valeur de la CB SC
        """ 
        
        # Coordonnées de la case actuelle
        flag = False
        # print( "current row")
        # print(current_row)
        # print("target value")
        # print(target_value)
        # # Parcourir les 15 cases au-dessus et en dessous de la case actuelle
        for row_offset in range(-15, 16):

                # Coordonnées de la case à vérifier
                row_to_check = ligne_current_cell + row_offset
                # print('row to check ')
                # print(row_to_check)
                # Récupérer la valeur et la date dans la case à vérifier
                if row_to_check >0:
                    cell_value = self.releve_de_compte_ws.cell(row=row_to_check, column=2).value
                    if (cell_value is not None and cell_value != "" ):
                        print("cell value : " + cell_value)
                    # Vérifier si la valeur et la date correspondent aux cibles
                    if (cell_value is not None and cell_value != "" ) and ("223293501" in cell_value) and (target_date in cell_value):
                        cell_date = self.releve_de_compte_ws.cell(row=row_to_check, column=2)
                        print("cell date")
                        print(cell_date.value)
                        if flag is True :
                            raise ErreurExcel(self.case_iteration,"Plusieurs cartes trouvées pour la date : " + target_date)
                        else: flag = True
                        
                    if flag is False:
                        raise ErreurExcel(self.case_iteration,"Aucune carte sans contact n'a été trouvée à cette date : " + target_date)
            
        return cell_date

    ## lire la case, analyser son contenu, mettre les bonnes valeurs au bon endroit, ou reporter valeur dans un excel
    def lecture_ligne_releve(self,ligne_releve, case_releve):
        try:
            value_case_releve = case_releve.value
            ##Si la case est une carte bleue
            if ("REMISE CARTE") in value_case_releve:
                print("find remise carte in line : "+ str(ligne_releve))
                ## Si la carte est une AvC
                if ("0697462") in value_case_releve:
                        print("find card Avc : "+ str(ligne_releve))
                    ## Rechercher la CB SC de la ^m journée
                    ## Ajouter les deux cartes bleues ensemble
                    ## check avec le tableur compta s'il y a match 
                    ## Case cochée si oui, Erreur si non
                    
                    ## match group : format "DD/MM"
                    
                        complete_date = self.extract_match_date_from_string(value_case_releve)

                        if complete_date != "no match":
                            case_carte_sc = self.trouver_case_carte_meme_date(ligne_releve,complete_date)                         
                            if self.verif_CB(self.getValeurCredit(case_carte_sc),self.getValeurCredit(self.case_iteration)) == True :
                                self.cocher_releve_de_compte(str(ligne_releve+ self.lignedebut ))
                            else : raise   
                            ErreurExcel(self.case_iteration,"Les valeurs de carte bleue ne correspondent pas pour cette date : " + complete_date )
                ## DABoi
            elif ("2903075") in value_case_releve:
                self.tableau_de_compta_ws["AI" + str(ligne_releve)].value = -(self.getValeurCredit(self.case_iteration).value)
                
        except ErreurExcel as e:  
                self.tableau_erreur.append((e.current_cellCoord, e.details_error))        
                
                return
            
    def AffilierDoss(
        self, file
    ):  ## def d'affiliation de paquetfile Apellée apres pression dur le bouton
        self.paquetfile = file

    def AffilierMA(self, mois, annee):  ## same pour nbquestion
        self.mois = mois
        self.annee = int(annee)

    def Execution(
        self
    ):
        
        try :
            self.tableau_de_compta_wb = pyxl.load_workbook(self.fileFeuilleDeCompta)
            self.releve_de_compte_wb  = pyxl.load_workbook(self.fileReleveDeCompte)

            for sheet in self.tableau_de_compta_wb:  ## recherche de la bonne feuille de la bonne année
                if sheet.title == str(
                    self.annee
                ):  ##si le nom de la feuille corrsepond à l'année
                    self.tableau_de_compta_ws = sheet  ##stockage de la sheet
            
            ## pour stocker et creer la case de départ
            ligne = 1  
            ## pour faire les cases
            colone = "A"  
            flag = True
            
            ## recherche du mois dans le tbleur de compta
            while ligne < 467 & flag == True:  
                casedate = colone + str(ligne)
                if self.tableau_de_compta_ws[casedate].value == self.mois:
                    self.lignedebut = (
                        ligne + 3,
                    )
                    flag = False
                    colone = "V"
                ligne = ligne + 1
                
                
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

            ## lire la colone de relevé de compte
            ##case de départ du relevé de compte

            
            self.releve_de_compte_ws = self.releve_de_compte_wb.active
            ligne_releve = 11 
            i = 0
            self.case_iteration =  str(self.colone_depart_releve) + str(self.ligne_depart_releve)
            while (self.releve_de_compte_ws[self.case_iteration].value is not None ):
                print(i)
                self.lecture_ligne_releve(ligne_releve,self.releve_de_compte_ws[self.case_iteration])
                ligne_releve = ligne_releve + 1
                self.case_iteration = "B" + str(ligne_releve)
                i = i+1
            
            print("chaussure :) ")  ## j'ai retrouvé mes chaussures!!
            
            fichier_enregistrement = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
        )
            
            
            self.fill_fichier_error()
            self.tableau_de_compta_wb.save(self.fileFeuilleDeCompta)  ##save de la feuille de compta
            self.excel_des_problemes.save(fichier_enregistrement)
            
        except ErreurExcel as e:  
                self.tableau_erreur.append((e.current_cellCoord, e.details_error))      
            
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
 
    ## la fonction qui permet de générer tous les éléments du front
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
        self.feuille_btn = tk.Button(
            canvas,
            text="Donner feuille de compta",
            command=lambda : self.donnerFile(canvas,"feuille_btn"),
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=button_color,
            fg=text_color,
        )
        self.feuille_btn.place(x=50 + x_spacing, y=80 + y_spacing, width=250)
        
        self.releve_btn = tk.Button(
            canvas,
            text="Donner relevé de compte",
            command=lambda: self.donnerFile(canvas,"releve_btn"),
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=button_color,
            fg=text_color,
        )
        self.releve_btn.place(x=50 + x_spacing, y=130 + y_spacing, width=250)
        
        selection_label = tk.Label(
            canvas,
            text="Saisir le mois et l'année:",
            font=font.Font(family=policyBody, size=sizeBody, weight=weightBody),
            bg=bg_color,
            fg=text_color,
        )
        selection_label.place(x=50 + x_spacing, y=180 + y_spacing)
        
        self.month_var = tk.StringVar(canvas)
        self.month_var.set("Mois")  # Ajoutez cette ligne pour définir "Mois" comme valeur par défaut
        month_dropdown = tk.OptionMenu(
            canvas,
            self.month_var,
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
        self.year_var = tk.StringVar(canvas)
        self.year_var.set("Année")
        year_dropdown = tk.OptionMenu(canvas, self.year_var, "2023", "2024", "2025")
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
            canvas, textvariable=self.month_var, font=font.Font(family=policyBody, size=sizeBody, weight=weightBody), bg=bg_color, fg=text_color
        )
        selected_month_label.place(x=50, y=320 + y_spacing)

        selected_year_label = tk.Label(
            canvas, textvariable=self.year_var, font=font.Font(family=policyBody, size=sizeBody, weight=weightBody), bg=bg_color, fg=text_color
        )
        selected_year_label.place(x=180, y=320 + y_spacing)
        
        # Bouton Lancer le programme
        launch_btn = tk.Button(
            canvas,
            text="Lancer le programme",
            command=lambda:self.TestVariable(),
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
        
    ## foncton permetant d'afficher un message d'erreur    
    def afficher_message(parent, message):
    # Fonction pour gérer l'événement du clic sur le bouton "OK"
        def fermer_fenetre():
            message_fenetre.destroy()

        # Création de la fenêtre
        message_fenetre = tk.Toplevel(parent)
        message_fenetre.title("Message d'erreur")
        message_fenetre.iconbitmap("logo.ico")

        # Création d'un label pour afficher le message
        label_message = tk.Label(message_fenetre, text=message, font=font.Font(family=policyBody, size=sizeBody, weight=weightBody))
        label_message.pack(padx=20, pady=20)

        # Création d'un bouton "OK"
        bouton_ok = tk.Button(
            message_fenetre, text="OK", command=fermer_fenetre, width=10, height=2
        )
        bouton_ok.pack(pady=20)

        # Calcul de la taille de la fenêtre en fonction de la longueur du message
        largeur_message_fenetre = max(350, label_message.winfo_reqwidth() + 40)
        hauteur_message_fenetre = 180

        # Récupération de la taille de la fenêtre parente
        parent.update_idletasks()  # Actualisation des tâches du parent
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        # Calcul de la position de la fenêtre pour la centrer
        xerrorwindow = parent.winfo_rootx() + (parent_width - largeur_message_fenetre) // 2
        yerrorwindow = parent.winfo_rooty() + (parent_height - hauteur_message_fenetre) // 2

        # Configuration de la position et de la taille de la fenêtre
        message_fenetre.geometry(
            f"{largeur_message_fenetre}x{hauteur_message_fenetre}+{xerrorwindow}+{yerrorwindow}"
        )

        # Lancement de la boucle principale de la fenêtre
        message_fenetre.mainloop()

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

    ## correspond à la commande effectuée lors du push des 2 premiers boutons
    def donnerFile(self,Cnvas,nomBouton):
        File = filedialog.askopenfilename(initialdir="Desktop/", title="Feuille de compta")
        print(File)
        
        if nomBouton == "feuille_btn":
            xcercle = int(largeur_fenetre) - 25
            ycercle = self.feuille_btn.winfo_y() + 18
            ## ajout cercle deuxieme bouton
        elif nomBouton == "releve_btn":
            xcercle = int(largeur_fenetre) - 25
            ycercle = self.releve_btn.winfo_y() + 20
            
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
            self.afficher_message(
                "ERREUR: Extension fichier non reconnue. \n Tu t'es surement trompée de fichier. \n Si le problème persiste, appelle Maxime.",
            )
        return ""
    
    
    

    def TestVariable(self):
        
        self.ReleveAuto.mois = self.month_var.get()
        self.ReleveAuto.annee = self.year_var.get()
        
        if self.ReleveAuto.fileFeuilleDeCompta == "":
            self.afficher_message("Tu n'as pas donné de tableur de compta")

        elif self.ReleveAuto.fileReleveDeCompte == "":
            self.afficher_message(
                "Tu n'as donné de le relevé de compte"
            )

        elif self.month_var.get() == "Mois":
            self.afficher_message("Tu n'as saisi le mois ")
        

        elif self.year_var.get() == "Année":
            self.afficher_message("Tu n'as saisi l'année ")

        else:        
            
            try: 
                self.ReleveAuto.mois = self.month_var.get()
                self.ReleveAuto.annee = self.year_var.get()
                self.ReleveAuto.Execution()
                self.afficher_message("Compta effectuée")
                            
            except ErreurExcel as e:  
                self.afficher_message(self,e.details_error)   
                    
        
    
if __name__ == "__main__":
    app = Application()
    app.mainloop()

print("chaussure :) ")  ##j'ai retrouvé mes chaussures !
