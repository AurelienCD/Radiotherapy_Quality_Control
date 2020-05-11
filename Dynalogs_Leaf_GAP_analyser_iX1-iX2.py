# -*- coding: utf-8 -*-

# Analyse des GAP entre les lames aux R1-R2 et ceci pour les bancs de lame A et B
# Auteur : Aurélien Corroyer-Dulmont
# Version : 09 mars 2020

# Update 09/03/2020 : all the interesting results are automatically upload to the excel file Alex made. To do that pandas, openpyxl and xlsxwriter are used
# Update 09/03/2020 : ask for new measurement   
# Update 12/03/2020 : will look at the folder in "Z:/qualité" and performed automatically the analysis to all the dynalogs files present in the folder, it will also copy all the results in the excel file and activate VBA macro to archive the results



from __future__ import print_function
import importlib
import os, string
import unittest
import logging
from math import *
import time
import datetime
import codecs
import array
import tkinter
from tkinter.filedialog import *
import statistics
import pandas as pad 
from openpyxl import load_workbook
import win32com.client
from path import Path
import shutil

date = datetime.datetime.now()

UserName = input("\nMerci de taper les initiales de l'utilisateur:\n\n")


def Dynalogs_Leaf_GAP_analyser(filepath):

    ## Récupération de la date de traitement et du n° du RapidArc avec le nom du fichier dynalogs ##
    LenghtFilePath = len(filepath)
    AcquisitionDate = filepath[LenghtFilePath-38:LenghtFilePath-30]
    AcquisitionDateExcel = str(AcquisitionDate[6:]) + "/" + str(AcquisitionDate[4:6]) + "/" + str(AcquisitionDate[:4])
    Banc = filepath[LenghtFilePath-39]
    NumeroRapid =  filepath[LenghtFilePath-14:LenghtFilePath-13]
    MachineName = "RapidArc_iX_" + str(NumeroRapid)

    savepath = "Z:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Results_iX1_iX2/Dynalogs_LEAF_GAP_analyser_results_" + str(MachineName) + "_" + str(AcquisitionDate[6:]) + str(AcquisitionDate[4:6]) + str(AcquisitionDate[:4]) + ".txt"
    

    # Calcul du nombre de ligne qui est variable d'un patient à l'autre
    file = open(filepath, 'r')

    LineNotEmpty = [1]
    LineCount = 0

    while (not LineNotEmpty) != True:
       LineNotEmpty = file.readlines(1)
       LineCount += 1

    LineCount -= 1 # car la boucle fait une ligne de plus vu qu'elle va jusqu'à ce que la ligne soit vide
    LineCount -= 6 # car on enlève les 6 premières lignes qui sont des informations sur l'ARC et non des mesures dynalogs
    file.close()   # on ferme le fichier car on souhaite l'ouvrir ensuite dans le code en partant du début
    # fin calcul du nombre de ligne


    file = open(filepath, 'r')

    i = 0
    for i in range(6):
        file.readlines(1)   
        
    # déclaration des list qui contiendront les valeurs des positions de lames attendues et réelles et celle contenant la différence (en mm) entre chaque lame
    ExpPosLeaf = []
    RealPosLeaf = []
    LeafGAP = []

    i = 0
    j = 0
    k = 0
    for i in range(LineCount): #boucle sur l'ensemble du nombre de ligne du fichier dynalog
        LineRaw = file.readlines(1) #première ligne d'intérêt avec les valeurs de positions de lames, mais copie dans une list tout le premier paragraphe, donc pas utilisable
        LineListStr = ",".join(LineRaw) # passage en mode string pour ensuite repasser en mode list mais en séparant les chiffres par la virgule de façon à avoir une liste avec une donnée par ...je ne connais pas le terme
        LineListGood = LineListStr.split(",") #permet d'avoir les valeurs du premier paragraphe dans une list
        
        for j in range(60):
            ExpPosLeaf.append(LineListGood[14+k])
            RealPosLeaf.append(LineListGood[15+k])
            k += 4
            j += 1
        k = 0

    # Boucle pour récupérer la différence entre les positions attendues et réelles (en 100ème de mm)
    m = 0
    IndexRepLame = LineCount*60 #car 60 lames
    for m in range(IndexRepLame):
        Difference = int(ExpPosLeaf[m]) - int(RealPosLeaf[m])
        LeafGAP.append(abs(Difference))

    # détermination de la différence max entre position attendues et réelles et ceci pour le banc de lame A et/ou B sachant que les positions seront positives ou négatives en fonction du banc de lames
    MinimumDifference = min(LeafGAP)
    MaximumDifference = max(LeafGAP)

    MaxDifference = MaximumDifference/100

    ######### Calcul de la conformité et mention dans le fichier résultats###########
    if MaxDifference < 1:
        ResultLeafGAP = "Conforme"
    else:
        if MaxDifference > 1:
            ResultLeafGAP = "Hors tolérance"
        else:
            ResultLeafGAP = "Limite" #car si pas < ou > à 1 alors = à 1


    LeafGAPIndex = LeafGAP.index(MaximumDifference) # donne l'indice où la différence est maximale, afin de retourner au n° de la lame

    #####création d'une table où les multiples des positions de lames sont présentes (dans le but de remonter au n° de lame défectueuse)
    TableLame = []
    MaxLeafGAP = []
    MeanLeafGAP = []
    SDLeafGAP = []
    LeafGAPTemp = []
    i = 0
    j = 0
    k = 0
    for k in range(60):
        for i in range(LineCount):
            TableLame.append(k+j)
            x = LeafGAP[k+j]
            LeafGAPTemp.append(x)
            j += 60
        MaxLeafGAP.append(max(LeafGAPTemp))
        MeanLeafGAP.append(statistics.mean(LeafGAPTemp))
        SDLeafGAP.append(statistics.stdev(LeafGAPTemp))
        LeafGAPTemp = []
        k += 1
        j = 0

    i=0
    for i in range(60):
        MeanLeafGAP[i] = round(MeanLeafGAP[i]/100, 3) # arrondit à 2 décimales et passage en mm
        SDLeafGAP[i] = round(SDLeafGAP[i]/100, 3)
        
    ## calcul des moyennes des écarts moyens et de l'écart-type de ces moyennes
    MeanLeafGAPAllLeaf = statistics.mean(MeanLeafGAP)
    SDLeafGAPAllLeaf = statistics.stdev(MeanLeafGAP)
    MeanLeafGAPAllLeaf = round(MeanLeafGAPAllLeaf, 3)
    SDLeafGAPAllLeaf = round(SDLeafGAPAllLeaf, 4)

    TableLameIndex = MaxLeafGAP.index(MaximumDifference) #on obtient donc l'indice dans la table qui est directement un multiple de 60 cad si <60 alors ce sera la lame n°1; si compris entre 1 et 2 alors lame n°2 etc etc...
    LameNumber = TableLameIndex+1 #car TableLameIndex renvoie un indice et non le n° de lame

    
    ######### Infos dans l'invite de commande Python###########
    if Banc == "A":
        print("POUR LE BANC A :\n")
        print(u"L'écart maximal entre la position attendue et la position réelle est de : " + str(MaxDifference) + " mm et ceci pour la lame n°" + str(LameNumber) +"\n")
        print(u"L'écart moyen entre la position attendue et la position réelle est de : " + str(MeanLeafGAPAllLeaf) + " mm \n")
        print(u"L'écart-type entre ces moyennes est de : " + str(SDLeafGAPAllLeaf) + " mm \n")
        i = 0
        for i in range(len(MaxLeafGAP)): # Boucle pour obtenir les écarts max des 60 lames
            print("Ecart maximal / moyen / SD pour la lame n°" + str(i+1) + ": " + str(MaxLeafGAP[i]/100) + " / " + str(MeanLeafGAP[i]) + " / " + str(SDLeafGAP[i]) + " mm")

            
    else:
        print("POUR LE BANC B :\n")
        print(u"L'écart maximal entre la position attendue et la position réelle est de : " + str(MaxDifference) + " mm et ceci pour la lame n°" + str(LameNumber) +"\n")
        print(u"L'écart moyen entre la position attendue et la position réelle est de : " + str(MeanLeafGAPAllLeaf) + " mm \n")
        print(u"L'écart-type entre ces moyennes est de : " + str(SDLeafGAPAllLeaf) + " mm \n")
        i = 0
        for i in range(len(MaxLeafGAP)): # Boucle pour obtenir les écarts max des 60 lames
            print("Ecart maximal / moyen / SD pour la lame n°" + str(i+1) + ": " + str(MaxLeafGAP[i]/100) + " / " + str(MeanLeafGAP[i]) + " / " + str(SDLeafGAP[i]) + " mm")


    print("\n")             
    print(u"Le résultat du test est : " + str(ResultLeafGAP.upper()) +"\n\n")
    print(u"L'ensemble des résultats sont dans le dossier : \n" + str(savepath) +"\n\n")
    
    ######### Création et remplissage fichier text pour stocker les résultats###########
    filesave = open(savepath, 'a')
    filesave = codecs.open(savepath, 'a', encoding='Latin-1')     # Encodage du fichier pour écriture incluant les "é" ###
    filesave.write(u"Résultats de l'analyse des Dynalogs")
    filesave.write("\n\n")
    filesave.write("Machine : " + str(MachineName))
    filesave.write("\n\n")
    filesave.write("Date d'acquisition : " + str(AcquisitionDate[6:]) + "/" + str(AcquisitionDate[4:6]) + "/" + str(AcquisitionDate[:4]))
    filesave.write("\n\n")
    filesave.write("Date d'analyse : " + str(date.day) + "/" + str(date.month) + "/" + str(date.year))
    filesave.write("\n\n")
    filesave.write("\n\n")
    if Banc == "A":
        filesave.write(u"Résultats de l'analyse des Dynalogs pour le banc A")
    else:
        filesave.write(u"Résultats de l'analyse des Dynalogs pour le banc B")
    filesave.write("\n\n")
    filesave.write("L'écart maximal entre la position attendue et la position réelle est de : " + str(MaxDifference) + " mm et ceci pour la lame n°" + str(LameNumber))
    filesave.write("\n\n")
    filesave.write(u"L'écart moyen entre la position attendue et la position réelle est de : " + str(MeanLeafGAPAllLeaf) + " mm")
    filesave.write("\n\n")
    filesave.write(u"L'écart-type entre ces moyennes est de : " + str(SDLeafGAPAllLeaf) + " mm")
    filesave.write("\n\n")
    filesave.write("Le résultat du test est : " + str(ResultLeafGAP.upper()))
    filesave.write("\n\n")

    i = 0
    for i in range(len(MaxLeafGAP)): # Boucle pour obtenir les écarts max des 60 lames
      filesave.write("Ecart maximal / moyen / SD pour la lame n°" + str(i+1) + ": " + str(MaxLeafGAP[i]/100) + " / " + str(MeanLeafGAP[i]) + " / " + str(SDLeafGAP[i]) + " mm")
      filesave.write("\n")  

    filesave.write("\n\n")
    filesave.close()
    
    ListOfResults = []
    ListOfResults = [MachineName, str(AcquisitionDateExcel), str(MaxDifference), str(MeanLeafGAPAllLeaf), str(SDLeafGAPAllLeaf)]
    
    return (ListOfResults)


#########                               9th Mars update                                   ###########
######### Create Pandas Excel functions to upload the results in the excel data base file ###########
def ExportToExcel(UserName, ListOfResultsA, ListOfResultsB):
    MachineName = str(ListOfResultsA[0])
    ListOfResultsToExcel = [ListOfResultsA[1]]
    ListOfResultsToExcel.append(UserName)
    ListOfResultsA = ListOfResultsA[2:]
    ListOfResultsB = ListOfResultsB[2:]
    ListOfResultsToExcel = [ListOfResultsToExcel+ListOfResultsA+ListOfResultsB]
    df = pad.DataFrame(ListOfResultsToExcel)
    #df = df.transpose()
    book = load_workbook('Z:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Répétabilité DLG PFROTAT.xlsm', read_only=False, keep_vba=True)   #####CHANGER PATH pour Z:/Aurelien_Dynalogs/Results/LEAF_GAP_PFROTAT/Répétabilité DLG PFROTAT.xlsm##########
    writer = pad.ExcelWriter('Z:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Répétabilité DLG PFROTAT.xlsm', engine='openpyxl') 
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df.to_excel(writer, str(MachineName),startrow=9, startcol=0, header=False, index=False)
    writer.save()

    xl = win32com.client.Dispatch('Excel.Application')
    xl.Workbooks.Open(Filename = 'Z:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Répétabilité DLG PFROTAT.xlsm', ReadOnly=1)  
    xl.Worksheets(str(MachineName)).Activate()
    ws = xl.ActiveSheet
    if MachineName == "RapidArc_iX_1":
        xl.Application.Run("Répétabilité_DLG_PFROTAT_IX1")
    else:
        xl.Application.Run("Répétabilité_DLG_PFROTAT_IX2")
    #xl.Application.Run("Actualiser_graphique_IX1")
    xl.Application.Quit()
    del xl


#########                               12th Mars update                                   ###########
#########                            loop to analyse all the files in the folder           ###########
### créer une fonction "fileListReturn" plutôt qu'un ensemble de ligne de code
fileList = []
for f in Path('Z:/Aurelien_Dynalogs/0000_Fichiers_Dynalogs_A_Analyser/PFRTOTAT_TOP').walkfiles(): 
    fileList.append(f)

newFileList = []
lastFileList = []
for i in range(len(fileList)):
    newFileList.append(fileList[i].replace('Path(',''))
    lastFileList.append(newFileList[i].replace('\\','/'))
    i += 1

dynalogFileList = lastFileList

question = input("\n\nIl y a " + str(len(dynalogFileList)) + " fichiers dynalogs à analyser,\n\nContinuer ? (O/N)\n\n")

if question == "O":
    for i in range(int(len(dynalogFileList)/2)):
        ListOfResultsA = Dynalogs_Leaf_GAP_analyser(str(dynalogFileList[i]))
        ListOfResultsB = Dynalogs_Leaf_GAP_analyser(str(dynalogFileList[i+int(len(dynalogFileList)/2)]))
        ExportToExcel(UserName, ListOfResultsA, ListOfResultsB)
        i += 1

    print("\n\nANALYSE TERMINEE\n\n")
    os.startfile("Z:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Répétabilité DLG PFROTAT.xlsm")
    os.system("pause")
    #### Déplacement des fichiers dynalogs analysés dans un répertoire d'archive ###
    for file in dynalogFileList:
        shutil.move(file, 'Z:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Archives')


else:
    print("Ok, au revoir")