:: Script pour connection du lecteur "qualité"" et déplacement des fichiers dynalogs des patients du jour précédent
:: La script est paramétré au niveau du R1 et R2 de façon à se lancer automatiquement lors de l'ouverture de la session

:: Auteur : Aurélien Corroyer-Dulmont
:: Version : 18 mai 2020

:: Update 09/03/2020


net use Q: \\172.27.22.37\Qualite /user:baclesse\radiot_rapidarc-ecli Max8Raw9 /persistent:yes

:: Pour RapidArc 1
:: move C:\Program File\Varian\oncology\MLC\controller\exec\dynalogs\*.dlg* "Q:\Aurelien_Dynalogs\0000_Fichiers_Dynalogs_A_Analyser\temp_export\"

:: Pour RapidArc 2
move D:\VMSOS\AppData\MLC\Controller\MLCDynalogs\dynalogs\*.dlg* "Q:\Aurelien_Dynalogs\0000_Fichiers_Dynalogs_A_Analyser\temp_export\"



