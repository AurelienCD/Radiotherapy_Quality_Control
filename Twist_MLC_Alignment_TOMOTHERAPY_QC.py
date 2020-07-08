# coding: utf-8


from __future__ import print_function
import importlib
import os, string
import unittest
import vtk, qt, ctk, slicer
from slicer.ScriptedLoadableModule import *
import logging
from __main__ import vtk, qt, ctk, slicer
from math import *
import numpy as np
from vtk.util import numpy_support
import SimpleITK as sitk
import sitkUtils as su
import time
import datetime
import codecs

#
# Twist_MLC_Alignment_TOMOTHERAPY_QC
#

class Twist_MLC_Alignment_TOMOTHERAPY_QC(ScriptedLoadableModule):
  
  def __init__(self, parent):
    ScriptedLoadableModule.__init__(self, parent)
    self.parent.title = "Twist_MLC_Alignment_TOMOTHERAPY_QC"
    self.parent.categories = ["QC Radiotherapy"]
    self.parent.dependencies = []
    self.parent.contributors = ["Aurelien CORROYER-DULMONT (Medical Physics department, Centre Francois Baclesse, CAEN, FRANCE)"]
    self.parent.helpText = """
The aim of this extension is to analyse automatically dosimetric films for quality control in clinical radiotherapy.
The quality control test concerned is the MLC Alignment test for annual control in tomotherapy (for more details please see: www.sfpm.fr (in french))
It performs a simple thresholding on the input volume, which allows to obtain 3 segments corresponding to the 3 blocks of irradiation. 
Then it calculates the distance in mm between the center (in X axis) of each block of irradiation to obtain the legal offset value.
With Offset value = (D1 - D2)/2
With D1 = distance between the center of the left block and the center of the block in the middle
With D2 = distance between the center of the righ block and the center of the block in the middle
If the offset is lower than 1.5mm (legal tolerance) the algorithm return "conforme"
Finally the extension return a text file containing all the information mentionned above.
"""
    self.parent.helpText += self.getDefaultModuleDocumentationLink()
    self.parent.acknowledgementText = """
"""


#
# Twist_MLC_Alignment_testWidget
#

class Twist_MLC_Alignment_TOMOTHERAPY_QCWidget(ScriptedLoadableModuleWidget):
  def setup(self):
    ScriptedLoadableModuleWidget.setup(self)

    # Instantiate and connect widgets ...
    
    #
    # Parameters Area
    #
    parametersCollapsibleButton = ctk.ctkCollapsibleButton()
    parametersCollapsibleButton.text = "Parameters"
    self.layout.addWidget(parametersCollapsibleButton)

    # Layout within the dummy collapsible button
    parametersFormLayout = qt.QFormLayout(parametersCollapsibleButton)

    #
    # input volume selector
    #
    self.inputSelector = slicer.qMRMLNodeComboBox()
    self.inputSelector.nodeTypes = ["vtkMRMLScalarVolumeNode"]
    self.inputSelector.selectNodeUponCreation = True
    self.inputSelector.addEnabled = False
    self.inputSelector.removeEnabled = False
    self.inputSelector.noneEnabled = False
    self.inputSelector.showHidden = False
    self.inputSelector.showChildNodeTypes = False
    self.inputSelector.setMRMLScene( slicer.mrmlScene )
    self.inputSelector.setToolTip( "Pick the input to the algorithm." )
    parametersFormLayout.addRow("Input Volume: ", self.inputSelector)
    
    #
    # Combo box to choose the machine
    # 
    self.machineNameSelector = qt.QComboBox()
    self.machineNameSelector.toolTip = "Select the name of the machine."
    self.machineNameSelector.addItem("Tomotherapy 1")
    self.machineNameSelector.addItem("Tomotherapy 2")
    parametersFormLayout.addRow("Machine's name: ", self.machineNameSelector)

    #
    # Apply Button
    #
    self.applyButton = qt.QPushButton("Apply")
    self.applyButton.toolTip = "Run the algorithm."
    self.applyButton.enabled = False
    parametersFormLayout.addRow(self.applyButton)

    # connections
    self.applyButton.connect('clicked(bool)', self.onApplyButton)
    self.inputSelector.connect("currentNodeChanged(vtkMRMLNode*)", self.onSelect)
    
    # Add vertical spacer
    self.layout.addStretch(1)

    # Refresh Apply button state
    self.onSelect()
    


  def cleanup(self):
    pass


  def onSelect(self):
    self.applyButton.enabled = self.inputSelector.currentNode()


  def onApplyButton(self):
    logic = Twist_MLC_Alignment_TOMOTHERAPY_QCLogic()
    Index = self.machineNameSelector.currentIndex
    logic.run(self.inputSelector.currentNode(), Index)

#
# Twist_MLC_Alignment_testLogic
#

class Twist_MLC_Alignment_TOMOTHERAPY_QCLogic(ScriptedLoadableModuleLogic):
  def hasImageData(self,volumeNode):
    """This is an example logic method that
    returns true if the passed in volume
    node has valid image data
    """
    if not volumeNode:
      logging.debug('hasImageData failed: no volume node')
      return False
    if volumeNode.GetImageData() is None:
      logging.debug('hasImageData failed: no image data in volume node')
      return False
    return True


  def run(self, inputVolume, Index):
    
    logging.info('Processing started') 
    
    DosiFilmImage = inputVolume
    displayNode = DosiFilmImage.GetDisplayNode()
    displayNode.AutoWindowLevelOff()
    displayNode.SetWindow(40000)
    displayNode.SetLevel(50000)

    logging.info(DosiFilmImage)
    date = datetime.datetime.now()
    savepath=u"//s-grp/grp/RADIOPHY/Personnel/Aurélien Corroyer-Dulmont/3dSlicer/Twist_MLC_Alignment_TOMOTHERAPY_QC_Results/Results_" + str(date.day) + str(date.month) + str(date.year) + ".txt"

    logging.info(savepath)
    logging.info(Index)

    # Stockage du nom de la machine en utilisant le choix de l'utilisateur dans la class Widget
    if Index == 0:
        machineName = 'Tomotherapy 1'
    else:
        machineName = 'Tomotherapy 2'

    # Création de la segmentation
    segmentationNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentationNode")
    segmentationNode.CreateDefaultDisplayNodes()
    segmentationNode.SetReferenceImageGeometryParameterFromVolumeNode(DosiFilmImage)

    logging.info(segmentationNode)

    # Création des segments editors temporaires
    segmentEditorWidget = slicer.qMRMLSegmentEditorWidget()
    segmentEditorWidget.setMRMLScene(slicer.mrmlScene)
    segmentEditorNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentEditorNode")
    segmentEditorWidget.setMRMLSegmentEditorNode(segmentEditorNode)
    segmentEditorWidget.setSegmentationNode(segmentationNode)
    segmentEditorWidget.setMasterVolumeNode(DosiFilmImage)

 
    # Création d'un segment après seuillage
    addedSegmentID = segmentationNode.GetSegmentation().AddEmptySegment("IrradiatedBlocks")
    segmentEditorNode.SetSelectedSegmentID(addedSegmentID)
    segmentEditorWidget.setActiveEffectByName("Threshold")
    effect = segmentEditorWidget.activeEffect()
    effect.setParameter("MinimumThreshold",str(30000))
    effect.setParameter("MaximumThreshold",str(45000))
    effect.self().onApply()
    
    # Passage en mode closed surface pour calcul des centres
    n = slicer.util.getNode('Segmentation')
    s = n.GetSegmentation()
    ss = s.GetSegment('IrradiatedBlocks')
    ss.AddRepresentation('Closed surface',vtk.vtkPolyData())


    # Division du segment en plusieurs segments (un par bloc d'irradiation)
    segmentEditorWidget.setActiveEffectByName("Islands")
    effect = segmentEditorWidget.activeEffect()
    effect.setParameter("Operation",str("SPLIT_ISLANDS_TO_SEGMENTS"))
    effect.setParameter("MinimumSize", 1000)
    effect.self().onApply()

    ######### Initialisation des variables fixes d'intérêt###########
    Segmentation_Name = 'Segmentation'
    Segment_Name = ["IrradiatedBlocks", "IrradiatedBlocks -_1", "IrradiatedBlocks -_2"]
    ListXaxisCenterOfBlock = [0,0,0] # initialisation de la liste contenant les valeurs Y centrales des blocs

    # Boucle de calcul des centres pour les 7 blocs (segment)
    for i in range(len(Segment_Name)): 
       n = slicer.util.getNode(Segmentation_Name)
       s = n.GetSegmentation()
       ss = s.GetSegment(Segment_Name[i])
       pd = ss.GetRepresentation('Closed surface')
       com = vtk.vtkCenterOfMass()
       com.SetInputData(pd)
       com.Update()
       CenterOfBlock = com.GetCenter() # CenterOfBlock est alors un tuple avec plusieurs variables (coordonées x,y,z)
       XaxisCenterOfBlock = (CenterOfBlock[0]) # Sélection de la 1ème valeur du tuple (indice 0) qui est la valeur dans l'axe X qui est l'unique valeur d'intérêt
       XaxisCenterOfBlock = abs(XaxisCenterOfBlock) # On passe en valeur absolue
       ListXaxisCenterOfBlock[i] = XaxisCenterOfBlock

    logging.info("X coordinates of the centre of blocks : " + str(ListXaxisCenterOfBlock))


    ######### Calcul de la distance en X entre les centres des différents blocs###########
    
    D1 = abs(ListXaxisCenterOfBlock[1]-ListXaxisCenterOfBlock[2]) # On récupère la distance entre le centre du bloc gauche et celui du milieu
    D2 = abs(ListXaxisCenterOfBlock[0]-ListXaxisCenterOfBlock[2]) # On récupère la distance entre le centre du bloc droit et celui du milieu
    Offset = D1 - D2
    OffsetInMm = Offset / 2 # Calcul de l'Offset qui est dans la réglementation
    logging.info("Offset in mm : " + str(OffsetInMm))

    ######### Création et remplissage fichier text pour stocker les résultats###########
    file = open(savepath, 'w')

    ### encodage du fichier pour écriture incluant les "é" ###
    file = codecs.open(savepath, encoding='utf-8')
    txt = file.read()
    file = codecs.open(savepath, "w", encoding='mbcs') 
    
    date = datetime.datetime.now()
    file.write(u"Résultats test -Twist MLC Alignment-")
    file.write("\n\n")
    file.write("Machine : " + str(machineName))
    file.write("\n\n")
    file.write("Date : " + str(date.day) +"/" + str(date.month) +"/" + str(date.year))
    file.write("\n\n")
    file.write("\n\n")
    i = 0

    for i in range(len(ListXaxisCenterOfBlock)): # Boucle pour obtenir les coordonées X des centres des 7 blocs
      file.write(u"Coordonnée X du centre du bloc n°" + str(i+1) + " : ")
      file.write(str(ListXaxisCenterOfBlock[i]))
      file.write("\n\n")  

    file.write("\n\n")
    file.write(u"Distance D1 entre centre du bloc de gauche et celui du milieu : " + str(D1))
    file.write("\n\n")
    file.write(u"Distance D2 entre centre du bloc de droite et celui du milieu : " + str(D2))
    file.write("\n\n")
    file.write(u"Valeur de l'offset voulue par le réglementation qui est égale à (D1-D2)/2 (en mm) : " + str(OffsetInMm))


    ######### Calcul de la conformité et mention dans le fichier résultats###########
    if  0 <= OffsetInMm < 1.5:
      Result = "Conforme"
    else:
      if OffsetInMm > 1.5:
        Result = "Hors tolerance"
      else:
        Result = "Limite" #car si pas < ou > à 1.5 alors = à 1.5
    
    if  OffsetInMm < 0:
      logging.info("Valeur de l'Offset négative, problème dans l'image ou dans le programme, contactez Aurélien Corroyer-Dulmont tel : 5768")
    
    logging.info(Result)
    
    file.write("\n\n")
    file.write("\n\n")
    file.write(u"Résultat : " + str(Result))
    file.close()
    
    logging.info('Processing completed')
    logging.info('\n\nResults are in the following file : ' + savepath) 
    return True