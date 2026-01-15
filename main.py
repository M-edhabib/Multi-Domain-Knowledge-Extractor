import tkinter as tk
import os
from tkinter import filedialog
import fitz
import re
import pandas as pd 
import shutil
import pyttsx3
def creation_dictionnaire_de_valeur_pdf(debut,fin,texte):
    pattern = rf"{debut}(.*?){fin}"
    resultat= re.search(pattern, texte,re.DOTALL)
    a= "n.a." 
    if resultat== None : 
        dictionnaire_resultat={f"{debut}":"Aucune correspondance"}
    elif a not in resultat.group(0).strip() and resultat != None:
        selection=resultat.group(0).strip()
        lignes = selection.split("\n")
        parties_numerique= [ligne for ligne in lignes if any(c.isdigit() for c in ligne) ]
        parties_texte= [ligne for ligne in lignes if ligne.replace(" ","").isalpha() or (not any(c.isdigit() for c in ligne) )] 
        dictionnaire_resultat=dict(zip(parties_texte,parties_numerique))
    else :
        dictionnaire_resultat={f"{debut}":"Aucune correspondance"}
    return dictionnaire_resultat
def extract_text_from_pdf(chemin_pdf,x):
    #ouvrir le fichier en mode lecture binaire
    doc= fitz.open(chemin_pdf) 
    #obtenir la page numéro x
    page= doc.load_page(x)
    text = page.get_text()
    text= text.replace('\n0\n','\n')
    text= text.replace('|','qui correspond à ')
    text= text.replace('Votre évolution: ','')
    text= text.replace('󰠟','').replace('󱖔','').replace('󰍍','').replace('󰀄','').replace('󱖘','').replace('󰴥','').replace('󰋜','')
    text =text.replace('\n / ', ' sur ')
    text= text.replace(f'\n{x}\n','\nfin\n')
    text= text.replace('\n& prov.\n',' & prov.\n')
    text =text.replace('\n€\n','€\n')
    text =text.replace('CA',"Chiffre d'Affaires")
    text = text.replace('\nmarché \n','marché\n')
    text = text.replace('\ns\n','s\n')
    text = text.replace('\nsecteur en France\n',' secteur en France\n')
    text = text.replace('\ncommune considérée\n',' commune considérée\n')
    text = text.replace('\nentre les deux derniers recensements\n',' entre les deux derniers recensements\n')
    text = text.replace('\nconsidérée\n',' considérée\n')
    text = text.replace('\nmême département.\n',' même département\n')
    text = text.replace('\n2022 vs 2021\n',' en 2022 par rapport à 2021\n')
    text = text.replace('\nconcurrents',' concurrents')
    text = text.replace('\nconcurrent',' concurrent')
    text =text.replace('\n\n','\n')
    text= text.replace("\nd'affaires\n"," d'affaires\n")
    text= text.replace("\n(EBE)\n"," (EBE)\n")
    return text
def taille_fichier(chemin_pdf):
    #Ouvrir du fichier PDF
    doc=fitz.open(chemin_pdf)
    #Taille du fichier PDF
    taille= len(doc)
    return taille
fenetre=tk.Tk()
fenetre.title("Interprétation DPS")
commentaire = tk.Label(fenetre , text = "Cliquez sur traiter les fichiers pour sélectionner le dossier des DPS.")
commentaire.pack()
bouton= tk.Button(fenetre, text="Traiter les fichier")
bouton.pack()

def Interpretation_fin_DPS_MP3():
    fichier_erreurs=[]
    compteur=1
    #ouvrir une boite de dialogue pour choisir le répertoire source
    repertoire_source = filedialog.askdirectory()
    #obtenir la liste des fichiers dans la répertoire source
    fichiers = os.listdir(repertoire_source)
    fichiers_pdf=[f for f in fichiers if f.endswith(".pdf")] 
    for fichier_pdf in fichiers_pdf:
        try:
            label_fichier.config(text ="Traitement du fichier :"+ fichier_pdf)
            fenetre.update()
            chemin_pdf = os.path.join(repertoire_source,fichier_pdf)
            #obtenir le nom du fichier sans l'extension
            nom_fichier_sans_extension = os.path.splitext(os.path.basename(chemin_pdf))[0]
            #créer le nom du dossier
            nom_dossier= os.path.join(os.path.dirname(chemin_pdf),nom_fichier_sans_extension) 
            #créer le dossier s'il n'existe pas
            os.makedirs(nom_dossier, exist_ok=True)
            chemin_pdf_dans_dossier = os.path.join(nom_dossier, fichier_pdf)
            if os.path.exists(chemin_pdf_dans_dossier):
                if os.path.getmtime(chemin_pdf) > os.path.getmtime(chemin_pdf_dans_dossier):
                    shutil.copy(chemin_pdf, nom_dossier)
            else:
                # Si le fichier PDF n'existe pas dans le dossier, le copie dans le nouveau dossier
                shutil.copy(chemin_pdf, nom_dossier)
                
            #copier le fichier DPS dans le nouveau dossier
            shutil.copy(chemin_pdf,nom_dossier)
            texte=""
            for i in range(taille_fichier(chemin_pdf)):
                texte+= extract_text_from_pdf(chemin_pdf,i)
            recettes = creation_dictionnaire_de_valeur_pdf("Recettes","Excédent",texte)
            if recettes["Recettes"] != 'Aucune correspondance':
                recettes_pd=pd.DataFrame.from_dict(recettes, orient='index', columns= ['Valeur'])
            elif recettes["Recettes"] =='Aucune correspondance':
                recettes=creation_dictionnaire_de_valeur_pdf("Votre chiffre d'affaires","Votre marge brute",texte)
                if recettes["Votre chiffre d'affaires"] != 'Aucune correspondance':
                    recettes_pd=pd.DataFrame.from_dict(recettes, orient='index', columns= ['Valeur'])
                else:
                    recettes=creation_dictionnaire_de_valeur_pdf("Votre chiffre d'affaires","Votre marge brute",texte) 
                    recettes_pd=pd.DataFrame.from_dict(recettes, orient='index', columns= ['Valeur'])
            excedent = creation_dictionnaire_de_valeur_pdf("Excédent","Résultat",texte)
            if excedent["Excédent"] != 'Aucune correspondance':
                excedent_pd=pd.DataFrame.from_dict(excedent, orient='index', columns= ['Valeur'])
            elif excedent["Excédent"] =='Aucune correspondance':
                excedent=creation_dictionnaire_de_valeur_pdf("Votre marge brute","Votre rentabilité",texte)
                if excedent["Votre marge brute"] != 'Aucune correspondance':
                    excedent_pd=pd.DataFrame.from_dict(excedent, orient='index', columns= ['Valeur'])
                else:
                    excedent=creation_dictionnaire_de_valeur_pdf("Votre marge brute","Votre rentabilité",texte) 
                    excedent_pd=pd.DataFrame.from_dict(excedent, orient='index', columns= ['Valeur'])
            resultat= creation_dictionnaire_de_valeur_pdf("Résultat","Achats",texte)
            if resultat["Résultat"] != 'Aucune correspondance':
                resultat_pd=pd.DataFrame.from_dict(resultat, orient='index', columns= ['Valeur'])
            elif resultat["Résultat"] =='Aucune correspondance':
                resultat= creation_dictionnaire_de_valeur_pdf("Votre rentabilité (EBE)","qui correspond à",texte)
                if resultat["Votre rentabilité (EBE)"] != 'Aucune correspondance':
                    resultat_pd=pd.DataFrame.from_dict(resultat, orient='index', columns= ['Valeur'])
                else:
                    resultat= creation_dictionnaire_de_valeur_pdf("Votre rentabilité","qui correspond à",texte)
                    resultat_pd=pd.DataFrame.from_dict(resultat, orient='index', columns= ['Valeur'])        
            achats_direct=creation_dictionnaire_de_valeur_pdf("Achats","Salaires",texte)
            achats_direct_pd=pd.DataFrame.from_dict(achats_direct, orient='index', columns= ['Valeur'])
            Salaire=creation_dictionnaire_de_valeur_pdf("Salaires","Charges Externes",texte)
            Salaire_pd=pd.DataFrame.from_dict(Salaire, orient='index', columns= ['Valeur'])
            charge_externes=creation_dictionnaire_de_valeur_pdf("Charges Externes","Impôts",texte)
            charge_externes_pd=pd.DataFrame.from_dict(charge_externes, orient='index', columns= ['Valeur'])
            Impot=creation_dictionnaire_de_valeur_pdf("Impôts","Dotations",texte)
            Impot_pd=pd.DataFrame.from_dict(Impot, orient='index', columns= ['Valeur'])
            Dotations_aux_Amort_et_Prov=creation_dictionnaire_de_valeur_pdf("Dotations","fin",texte)
            Dotations_aux_Amort_et_Prov_pd=pd.DataFrame.from_dict(Dotations_aux_Amort_et_Prov, orient='index', columns= ['Valeur'])
            evo_recettes = creation_dictionnaire_de_valeur_pdf("Évolution de vos recettes","Évolution de votre excédent",texte)
            if evo_recettes["Évolution de vos recettes"] != 'Aucune correspondance':
                evo_recettes_pd=pd.DataFrame.from_dict(evo_recettes, orient='index', columns= ['Valeur'])
            elif evo_recettes["Évolution de vos recettes"] =='Aucune correspondance':
                evo_recettes=creation_dictionnaire_de_valeur_pdf("Evolution de votre chiffre d'affaires","Evolution de votre taux de marge brute",texte)
                if evo_recettes["Evolution de votre chiffre d'affaires"] != 'Aucune correspondance':
                    evo_recettes_pd=pd.DataFrame.from_dict(evo_recettes, orient='index', columns= ['Valeur'])
                else:
                    evo_recettes=creation_dictionnaire_de_valeur_pdf("Evolution de votre chiffre d'affaires","Evolution de votre taux de marge brute",texte)
                    evo_recettes_pd=pd.DataFrame.from_dict(evo_recettes, orient='index', columns= ['Valeur'])
            evo_excedent = creation_dictionnaire_de_valeur_pdf("Évolution de votre excédent","Évolution de votre résultat",texte)
            if evo_excedent["Évolution de votre excédent"] != 'Aucune correspondance':
                evo_excedent_pd=pd.DataFrame.from_dict(evo_excedent, orient='index', columns= ['Valeur'])
            elif evo_excedent["Évolution de votre excédent"] == 'Aucune correspondance':
                evo_excedent=creation_dictionnaire_de_valeur_pdf("Evolution de votre taux de marge brute","Evolution de votre taux de rentabilité",texte)
                if evo_excedent["Evolution de votre taux de marge brute"] != 'Aucune correspondance':
                    evo_excedent_pd=pd.DataFrame.from_dict(evo_excedent, orient='index', columns= ['Valeur'])
                else:
                    evo_excedent=creation_dictionnaire_de_valeur_pdf("Evolution de votre taux de marge brute","Evolution de votre taux de rentabilité",texte)
                    evo_excedent_pd=pd.DataFrame.from_dict(evo_excedent, orient='index', columns= ['Valeur'])
            evo_resultat= creation_dictionnaire_de_valeur_pdf("Évolution de votre résultat","Evolution du poids de vos achats directs",texte)
            if evo_resultat["Évolution de votre résultat"] != 'Aucune correspondance':
                evo_resultat_pd=pd.DataFrame.from_dict(evo_resultat, orient='index', columns= ['Valeur'])
            elif evo_resultat["Évolution de votre résultat"] == 'Aucune correspondance':
                evo_resultat=creation_dictionnaire_de_valeur_pdf("Evolution de votre taux de rentabilité","Evolution du poids de vos achats directs",texte)
                if evo_resultat["Evolution de votre taux de rentabilité"] != 'Aucune correspondance':
                    evo_resultat_pd=pd.DataFrame.from_dict(evo_resultat, orient='index', columns= ['Valeur'])
                else:
                    evo_resultat=creation_dictionnaire_de_valeur_pdf("Evolution de votre taux de rentabilité","Evolution du poids de vos achats directs",texte)
                    evo_resultat_pd=pd.DataFrame.from_dict(evo_resultat, orient='index', columns= ['Valeur'])
            evo_achats_direct=creation_dictionnaire_de_valeur_pdf("Evolution du poids de vos achats directs","Evolution du poids de vos salaires & charges",texte)
            evo_achats_direct_pd=pd.DataFrame.from_dict(evo_achats_direct, orient='index', columns= ['Valeur'])
            evo_Salaire=creation_dictionnaire_de_valeur_pdf("Evolution du poids de vos salaires & charges","Evolution du poids de vos charges externes",texte)
            evo_Salaire_pd=pd.DataFrame.from_dict(evo_Salaire, orient='index', columns= ['Valeur'])
            evo_charge_externes=creation_dictionnaire_de_valeur_pdf("Evolution du poids de vos charges externes","fin",texte)
            evo_charge_externes_pd=pd.DataFrame.from_dict(evo_charge_externes, orient='index', columns= ['Valeur'])
            croissance_CA = creation_dictionnaire_de_valeur_pdf("Conjoncture du secteur","Attractivité de l'emplacement",texte)
            croissance_CA_pd=pd.DataFrame.from_dict(croissance_CA, orient='index', columns= ['Valeur'])
            Attractivite_emplacement = creation_dictionnaire_de_valeur_pdf("Attractivité de l'emplacement","Population locale",texte)
            Attractivite_emplacement_pd=pd.DataFrame.from_dict(Attractivite_emplacement, orient='index', columns= ['Valeur'])
            population_loc= creation_dictionnaire_de_valeur_pdf("Population locale","Situation concurrentielle",texte)
            population_loc_pd=pd.DataFrame.from_dict(population_loc, orient='index', columns= ['Valeur'])
            Situation_concurrentielle=creation_dictionnaire_de_valeur_pdf("Situation concurrentielle","Immobilier",texte)
            Situation_concurrentielle_pd=pd.DataFrame.from_dict(Situation_concurrentielle, orient='index', columns= ['Valeur'])
            Immobilier=creation_dictionnaire_de_valeur_pdf("Immobilier","fin",texte)
            Immobilier_pd=pd.DataFrame.from_dict(Immobilier, orient='index', columns= ['Valeur'])
            chemin_excel= os.path.join(nom_dossier, nom_fichier_sans_extension+'.xlsx')
            with pd.ExcelWriter(chemin_excel , engine = 'xlsxwriter') as writer:
                recettes_pd.to_excel(writer , sheet_name ='recettes_pd')
                evo_recettes_pd.to_excel(writer , sheet_name ='evo_recettes_pd')
                excedent_pd.to_excel(writer , sheet_name ='excedent_pd')
                evo_excedent_pd.to_excel(writer , sheet_name ='evo_excedent_pd')
                resultat_pd.to_excel(writer , sheet_name ='resultat_pd')
                evo_resultat_pd.to_excel(writer , sheet_name ='evo_resultat_pd')
                achats_direct_pd.to_excel(writer , sheet_name ='achats_direct_pd')
                Salaire_pd.to_excel(writer , sheet_name ='Salaire_pd')
                charge_externes_pd.to_excel(writer , sheet_name ='charge_externes_pd')
                Impot_pd.to_excel(writer , sheet_name ='Impot_pd')
                Dotations_aux_Amort_et_Prov_pd.to_excel(writer , sheet_name ='Dotations_aux_Amort_et_Prov_pd')
                evo_achats_direct_pd.to_excel(writer , sheet_name ='evo_achats_direct_pd')
                evo_Salaire_pd.to_excel(writer , sheet_name ='evo_Salaire_pd')
                evo_charge_externes_pd.to_excel(writer , sheet_name ='evo_charge_externes_pd')
                croissance_CA_pd.to_excel(writer , sheet_name ='croissance_CA_pd')
                Attractivite_emplacement_pd.to_excel(writer , sheet_name ='Attractivite_emplacement_pd')
                population_loc_pd.to_excel(writer , sheet_name ='population_loc_pd')
                Situation_concurrentielle_pd.to_excel(writer , sheet_name ='Situation_concurrentielle_pd')
                Immobilier_pd.to_excel(writer , sheet_name ='Immobilier_pd')
            xls=pd.ExcelFile(chemin_excel)
            feuilles = xls.sheet_names
            interpretation=f"""Analyse de votre situation financière :
            """
            for feuille in feuilles:
                df=pd.read_excel(xls,feuille)
                df.rename(columns={'Unnamed: 0':'Mesures', }, inplace = True)
                df['Valeur']=df['Valeur'].str.replace('pt','point').str.replace('\u202f','')
                if feuille =='recettes_pd' :
                    l=[]
                    for mes in df['Mesures']:
                        l.append(mes)
                    if (l[0] == "Recettes" or l[0] =="Votre chiffre d'affaires") and (df.loc[0,'Valeur']!='Aucune correspondance'):
                        for mesure in l:
                            #conditions:
                            if (mesure == "Recettes" ) and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'):
                                #interprétation   
                                interpretation+= f""" Votre recette :Représente les revenus que vous avez générés .
                                - Vos revenus sont de {df.loc[l.index(mesure),'Valeur']}"""
                                R=mesure 
                            elif (mesure == "Votre chiffre d'affaires" ) and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'):
                                #interprétation   
                                interpretation+= f""" {df.loc[l.index(mesure),'Mesures']} :Représente les revenus que vous avez générés .
                                - Vos revenus sont de {df.loc[l.index(mesure),'Valeur']}"""
                                R=mesure 
                                #conditions:
                            if (mesure == "Médiane du marché") and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'):
                                # Conditions d'interprétation du chiffre d'affaire (recettes) par rapport à la médiane du marché
                                if (float(df.loc[l.index(R),'Valeur'].replace(',','.').replace('\xa0€','')))> (float(df.loc[l.index(mesure),'Valeur'].replace(',','.').replace('\xa0€',""))):
                                    interpretation_med_CA = 'supérieurs à la médiane'
                                elif (float(df.loc[l.index(R),'Valeur'].replace(',','.').replace('\xa0€',""))== float(df.loc[l.index(mesure),'Valeur'].replace(',','.').replace('\xa0€',""))):
                                    interpretation_med_CA = 'égales à la médiane'
                                else :
                                    interpretation_med_CA = 'inférieurs à la médiane'
                                #interprétation 
                                interpretation+= f""", ce qui est {interpretation_med_CA} {df.loc[l.index(mesure),'Valeur']}. Cette médiane divise votre panel en deux : la première moitié a une valeur inférieure et la seconde moitié a une valeur supérieure.
                                    """
                            #conditions:
                            if(mesure == "Votre positionnement") and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'): 
                                #interprétation
                                interpretation+= f"""Votre positionnement par rapport à votre panel est {df.loc[l.index(mesure),'Valeur']} entreprise . Le panel est constitué avec les données d’entreprises du même secteur d’activité (NAF) et du secteur géographique correspondant à celui de votre entreprise.
                                    """
                if feuille =='evo_recettes_pd':
                    #conditions:
                    if df.loc[0,'Mesures']=='Évolution de vos recettes' or df.loc[0,'Mesures']=="Evolution de votre chiffre d'affaires": 
                        #Conditions d'interprétation évolution du chiffre d'affaire (recettes) par rapport à l'année précédente :
                        if (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>0 and (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100) <0.1:
                            evolution_CA = 'amélioration modeste'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>=0.1:
                            evolution_CA = 'amélioration significative'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)<0 and (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100) > -0.1 :
                            evolution_CA = 'légère diminution'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)==0:
                            evolution_CA = 'Stabilité'
                        else :
                            evolution_CA= 'diminution significative'
                        #interprétation
                        interpretation+= f"""Votre évolution par rapport à l'année dernière est de {df.loc[0,'Valeur']}. Il s'agit d'une {evolution_CA}.
                        """
                if feuille =='excedent_pd':
                    l=[]
                    for mes in df['Mesures']:
                        l.append(mes)
                    if (l[0] == "Excédent" or l[0] =="Votre marge brute") and (df.loc[0,'Valeur']!='Aucune correspondance'):
                        for mesure in l:
                            #conditions:
                            if (mesure == "Excédent"):                                     
                                E=mesure
                                #interprétation
                                interpretation+=f"""Votre ratio d'{df.loc[l.index(mesure),'Mesures']} :évalue votre {df.loc[l.index(mesure),'Mesures']} par rapport à votre {R}. Votre {df.loc[l.index(E),'Mesures']} est de {df.loc[l.index(E),'Valeur']} . 
                                        """
                            elif( mesure =="Votre marge brute"):
                                E=mesure
                                #interprétation
                                interpretation+=f"""Le ratio de {df.loc[l.index(mesure),'Mesures']} :évalue {df.loc[l.index(mesure),'Mesures']} par rapport à {R} . {df.loc[l.index(E),'Mesures']} est de {df.loc[l.index(E),'Valeur']} . 
                                        """
                            #conditions:
                            if (mesure == "Médiane du marché") and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'):
                                #Conditions interprétation de la marge brute (excédent) par rapport à la médiane du marché
                                if  (float(df.loc[l.index(E),'Valeur'].replace(',','.').replace('%',''))> float(df.loc[l.index(mesure),'Valeur'].replace(',','.').replace('%',""))):
                                    interpretation_med_Exc = 'supérieur à la médiane'
                                elif (float(df.loc[l.index(E),'Valeur'].replace(',','.').replace('%',''))== float(df.loc[l.index(mesure),'Valeur'].replace(',','.').replace('%',""))):
                                    interpretation_med_Exc = 'égale à la médiane '
                                else :
                                    interpretation_med_Exc = 'inférieur à la médiane'
                                #interprétation
                                interpretation+= f"""La médiane du secteur selon votre panel est égale à {df.loc[l.index(mesure),'Valeur']} , vous êtes {interpretation_med_Exc}. 
                                """
                            #conditions:
                            if(mesure =='Votre positionnement') and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'):
                                interpretation+= f""" Votre positionnement par rapport à votre panel est  {df.loc[l.index(mesure),'Valeur']}.
                                """   
                if feuille =='evo_excedent_pd':
                    #conditions:
                    if df.loc[0,'Mesures']=='Évolution de votre excédent' or df.loc[0,'Mesures']=="Evolution de votre taux de marge brute":
                        #interprétation évolution de la marge brute (excédent) par rapport à l'année précédente :
                        if (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))>0 and (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point',''))) <10:
                            evolution_MB = 'amélioration modeste'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))>=10:
                            evolution_MB = 'amélioration significative'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))<0 and (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point',''))) > -10 :
                            evolution_MB = 'légère diminution'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))==0:
                            evolution_MB='Stabiltié'
                        else :
                            evolution_MB = 'diminution significative'
                        #interprétation            
                        interpretation+=f"""- Votre évolution par rapport à l'année dernière est de {df.loc[0,'Valeur']}. Il s'agit d'une {evolution_MB}. Cette évolution mesure l'écart entre vos ratios de l'année courante et ceux de l'année précédente et elle est mesuré par point.
                            """
                if feuille =='resultat_pd':
                    l=[]
                    for mes in df['Mesures']:
                        l.append(mes)
                    if (l[0] == "Résultat" or l[0] =="Votre rentabilité (EBE)") and (df.loc[0,'Valeur']!='Aucune correspondance'):
                        for mesure in l:
                            #conditions:
                            if (mesure == "Résultat"): 
                                Re=mesure 
                                #interprétation 
                                interpretation+= f"""Votre ratio du {df.loc[l.index(mesure),'Mesures']}: évalue votre {df.loc[l.index(mesure),'Mesures']} par rapport à votre {R}. Votre {df.loc[l.index(Re),'Mesures']} est égale à {df.loc[l.index(Re),'Valeur']}. """
                                
                            elif( mesure =="Votre rentabilité (EBE)"):
                                Re=mesure
                                #interprétation 
                                interpretation+= f""" le ratio de {df.loc[l.index(mesure),'Mesures']}: évalue {df.loc[l.index(mesure),'Mesures']} par rapport à {R}. {df.loc[l.index(Re),'Mesures']} est égale à {df.loc[l.index(Re),'Valeur']}. """
                            #conditions:
                            if (mesure == "Médiane du marché") and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'):
                                # Conditions d'interprétation de la rentabilité (résultat) par rapport à la médiane du marché
                                if  (float(df.loc[l.index(Re),'Valeur'].replace(',','.').replace('%',''))> float(df.loc[l.index(mesure),'Valeur'].replace(',','.').replace('%',""))):
                                    interpretation_med_Re = 'supérieur à la médiane'
                                elif (float(df.loc[l.index(Re),'Valeur'].replace(',','.').replace('%',''))== float(df.loc[l.index(mesure),'Valeur'].replace(',','.').replace('%',""))):
                                    interpretation_med_Re = 'égale à la médiane'
                                else :
                                    interpretation_med_Re = 'inférieur à la médiane'
                                #interprétation 
                                interpretation+= f""" la médiane du secteur est {df.loc[l.index(mesure),'Valeur']}, votre êtes {interpretation_med_Re}.
                                    """
                            #conditions:
                                if(mesure =='Votre positionnement') and (df.loc[l.index(mesure),'Valeur']!='Aucune correspondance'): 
                                    #interprétation
                                    interpretation+= f"""Votre positionnement par raport à votre panel est  {df.loc[l.index(mesure),'Valeur']}
                                    print(mesure)
                                    """                 
                if feuille =='evo_resultat_pd':
                    #conditions:
                    if df.loc[0,'Mesures']=='Évolution de votre résultat' or df.loc[0,'Mesures']=="Evolution de votre taux de rentabilité":
                        #interprétation évolution de la rentabilité (résultat) par rapport à l'année précédente :
                        if (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))>0 and (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point',''))) <10:
                            evolution_Re = 'amélioration modeste'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))>=10:
                            evolution_Re = 'amélioration significative'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))<0 and (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point',''))) > -10 :
                            evolution_Re = 'légère diminution'
                        elif (float(df.loc[0,'Valeur'].replace(',','.').split(" ")[0].replace('point','')))==0:
                            evolution_Re='Stabilité'
                        else :
                            evolution_Re = 'diminution significative'
                        #interprétation
                        interpretation+= f"""Votre évolution par rapport à l'année dernière est de {df.loc[0,'Valeur']}. Il s'agit d'une {evolution_Re}.
                            """                
                        
                if feuille =='achats_direct_pd':
                    if df.loc[0,'Mesures'] =='Achats directs' and df.loc[0,'Valeur']!='Aucune correspondance':
                        #interprétation
                        interpretation+=f"""Vos {df.loc[0,'Mesures']} représentent {df.loc[0,'Valeur'].replace(' %','%').split(" ")[0]} de vos revenus .
                            """
                if feuille =='charge_externes_pd':
                    if df.loc[0,'Mesures'] =='Charges Externes' and df.loc[0,'Valeur']!='Aucune correspondance':
                        #interprétation
                        interpretation+= f""" Vos {df.loc[0,'Mesures']} représentent {df.loc[0,'Valeur'].replace(' %','%').split(" ")[0]} de vos revenus .
                            """
                if feuille =='croissance_CA_pd':
                    #conditions:
                    if df.loc[1,'Mesures']== "Evolution de la croissance du chiffre d'affaires du secteur en France" and df.loc[1,'Valeur']!='Aucune correspondance': 
                        #interprétation conjoncture secteur:
                        if (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>0 and(float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)<0.1:
                            Croissance_CA = 'amélioration modeste'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>=0.1:
                            Croissance_CA = 'amélioration significative'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)<0 and (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>-0.1 :
                            Croissance_CA = 'légère diminution'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)==0:
                            Croissance_CA='Stabilité'
                        else :
                            Croissance_CA = 'diminution significative'
                    #interprétation
                        interpretation+= f"""Panorama du marché :
                        Les revenus générés dans votre secteur sur la France ont enregistré une {Croissance_CA} égale à {df.loc[1,'Valeur']}.
                        """
                if feuille =='Attractivite_emplacement_pd':
                    #conditions:
                    if df.loc[1,'Mesures']== "Evolution du nombre d’établissements dans la commune considérée" and df.loc[1,'Valeur']!='Aucune correspondance':
                        #interprétation attractivité de l'emplacement:
                        if (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>0 and(float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)<0.1:
                            Attractivté_emp = 'amélioration modeste'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>=0.1:
                            Attractivté_emp = 'amélioration significative'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)<0 and (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>-0.1 :
                            Attractivté_emp = 'légère diminution'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)==0:
                            Attractivté_emp = 'Stabilité'
                        else :
                            Attractivté_emp = 'diminution significative'
                    #interprétation
                    interpretation+= f"""Le nombre d'établissements dans votre commune a enregistré une {Attractivté_emp} égale à {df.loc[1,'Valeur']}.
                        """
                if feuille =='Situation_concurrentielle_pd':
                    #conditions:
                    if df.loc[1,'Mesures']== "Evolution du nombre de concurrents dans la commune considérée" and df.loc[1,'Valeur']!='Aucune correspondance':
                        #interprétation situation concurrentielle:
                        if (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0]))>0 and(float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0]))<5:
                            Situation_concurrentielle = 'amélioration modeste'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0]))>=5:
                            Situation_concurrentielle = 'amélioration significative'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0]))<0 and (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0]))>-5 :
                            Situation_concurrentielle = 'légère diminution'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0]))==0:
                            Situation_concurrentielle = 'stabilité'
                        else :
                            Situation_concurrentielle = 'diminution significative'
                    #interprétation
                    interpretation+= f"""Le nombre de concurrents dans votre commune a enregistré une {Situation_concurrentielle} égale à {df.loc[1,'Valeur']}.
                        """
                if feuille =='Immobilier_pd':
                    #conditions:
                    if df.loc[1,'Mesures']== "Evolution des prix des locaux commerciaux dans le même département" and df.loc[1,'Valeur']!='Aucune correspondance':
                        #interprétation Immobilier:
                        if (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>0 and(float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)<0.1:
                            Immobilier = 'augmentation modeste'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>=0.1:
                            Immobilier = 'augmentation significative'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)<0 and (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)>-0.1 :
                            Immobilier = 'légère diminution'
                        elif (float(df.loc[1,'Valeur'].replace(',','.').split(" ")[0].replace('%',''))/100)==0:
                            Immobilier = 'stabilité'
                        else :
                            Immobilier = 'diminution significative'
                    #interprétation
                    interpretation+= f"""Le prix des locaux commerciaux dans votre département a enregistré une {Immobilier} égale à {df.loc[1,'Valeur']}.""" 
            chemin_script = os.path.join(nom_dossier,'SCRIPT.txt')
            # Ouvrir le fichier en mode écriture ('w')
            with open(chemin_script, 'w') as fichier:
            # Écrire le script dans le fichier
                fichier.write(interpretation)
            engine =pyttsx3.init()
            voices = engine.getProperty('voices')
            engine.setProperty('voice', voices[0].id)
            engine.setProperty('rate',220)
            chemin_mp3=os.path.join(nom_dossier,'Interpretation_DPS.mp3')
            engine.save_to_file(interpretation, chemin_mp3)
            engine.runAndWait()
            label_compteur.config(text= "Nombre de fichier traités :"+ str(compteur) +"/"+str(len(fichiers_pdf)))
            compteur+=1
        except Exception as e:
            print(f"Une erreur s'est produite avec le fichier {fichier_pdf} : {str(e)} ")
    chemin_erreur = os.path.join(nom_dossier,'Erreurs.txt')
    with open (chemin_erreur ,'w') as f:
        for fich in fichier_erreurs :
            f.write(f"{fich}\n")
    fenetre.update()   
    fin_tache.config(text ="Tâches Terminées ")
    bouton_fermer.pack()
    
#Associer l'action du bouton à la fonction on_button_click        
bouton.config(command=Interpretation_fin_DPS_MP3)
#Lancer la boucle principale de Tkinter
label_fichier =tk.Label(fenetre, text ="")
label_fichier.pack()
fin_tache =tk.Label(fenetre, text ="")
fin_tache.pack()
label_compteur =tk.Label(fenetre, text="")
label_compteur.pack()
bouton_fermer=tk.Button(fenetre,text="Fermer", command = fenetre.destroy)
fenetre.mainloop()
