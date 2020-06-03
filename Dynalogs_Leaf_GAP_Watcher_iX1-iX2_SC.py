# -*- coding: utf-8 -*-

# Watch the dynalogs PFROTAT results excel file to see if the daily analysis is conform or not.
# If not, the program will open the excel file to see more details
# Author : Aurélien Corroyer-Dulmont
# Version : 03 may 2020


import time
import os


def Dynalogs_Leaf_GAP_Watcher(filepath):

	file = open(filepath, 'r')
	result = file.read()
	
	if result == "1":
		Result = "Hors Tolérance"
	else:
		Result = "Conforme"

	file.close()

	return Result




R1_path = "Q:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Results_iX1_iX2/temp_results_R1.txt"
R2_path = "Q:/Aurelien_Dynalogs/Results_Analyses_Dynalogs/LEAF_GAP_PFROTAT/Results_iX1_iX2/temp_results_R2.txt"

R1_result = Dynalogs_Leaf_GAP_Watcher(R1_path)
R2_result = Dynalogs_Leaf_GAP_Watcher(R2_path)


if R1_result == "Hors Tolérance":
	print("\n\nRESULTATS NON CONFORMES POUR LE RAPIDARC-1 \n\n")
	os.startfile('\\\\s-grp\\grp\\RADIOPHY\\Contrôle Qualité RTE\\Contrôle Qualité RTE-accélérateurs\\7_CLINAC iX 1\\7-3 CQ -EN\\7-1 CQ_quotidien\\CQ quotidien Dynalog PFROTAT - iX1.xlsm')
	os.system("pause")
else:
	print("\n\nRESULTATS CONFORMES POUR LE RAPIDARC-1 \n\n")

if R2_result == "Hors Tolérance":
	print("\n\nRESULTATS NON CONFORMES POUR LE RAPIDARC-2 \n\n")
	os.startfile('\\\\s-grp\\grp\\RADIOPHY\\Contrôle Qualité RTE\\Contrôle Qualité RTE-accélérateurs\\10_CLINAC iX 2\\10-3 CQ -EN\\10-1_CQ_quotidien\\CQ quotidien Dynalog PFROTAT - iX2.xlsm')
	os.system("pause")
else:
	print("\n\nRESULTATS CONFORMES POUR LE RAPIDARC-2 \n\n")
	time.sleep(3)
