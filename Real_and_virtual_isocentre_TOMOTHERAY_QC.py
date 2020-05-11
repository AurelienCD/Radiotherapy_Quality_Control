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
# Real_and_virtual_isocentre_TOMOTHERAY_QC
#

class Real_and_virtual_isocentre_TOMOTHERAY_QC(ScriptedLoadableModule):
 
  def __init__(self, parent):
    ScriptedLoadableModule.__init__(self, parent)
    self.parent.title = "Real and Virtual Isocentre TOMOTHERAPY QC"
    self.parent.categories = ["QC Radiotherapy"]
    self.parent.dependencies = []
    self.parent.contributors = ["Aurelien CORROYER-DULMONT (Medical Physics department, Centre Francois Baclesse, CAEN, FRANCE)"]
    self.parent.helpText = """
The aim of this extension is to analyse automatically dosimetric films for quality control in clinical radiotherapy.
The quality control test concerned is the Real and Virtual Isocentre for annual control in tomotherapy (for more details please see: www.sfpm.fr (in french))

"""
    self.parent.helpText += self.getDefaultModuleDocumentationLink()
    self.parent.acknowledgementText = """"""
  
#
# Real_and_virtual_isocentre_TOMOTHERAY_QCWidget
#

class Real_and_virtual_isocentre_TOMOTHERAY_QCWidget(ScriptedLoadableModuleWidget):
    
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
    logic = Real_and_virtual_isocentre_TOMOTHERAY_QCLogic()
    Index = self.machineNameSelector.currentIndex
    logic.run(self.inputSelector.currentNode(), Index)


### message à l'écran ###
#slicer.util.confirmOkCancelDisplay("Cliquez sur les points gauche et droite du laser et sur celui en dessous de l'isocentre",windowTitle=None,parent=None)

        
#
# Real_and_virtual_isocentre_TOMOTHERAY_QCLogic
#
class Real_and_virtual_isocentre_TOMOTHERAY_QCLogic(ScriptedLoadableModuleLogic):
  
  def run(self, inputVolume, Index):
    
    logging.info('Processing started') 
    
    DosiFilmImage = inputVolume

    date = datetime.datetime.now()
    
    # Stockage du nom de la machine en utilisant le choix de l'utilisateur dans la class Widget
    if Index == 0:
        machineName = 'Tomotherapy 1'
    else:
        machineName = 'Tomotherapy 2'

    # To obtain fiducial position from user
    markups = slicer.util.getNode('F')
    FiducialCoordinatesF1 = [0,0,0,0]
    FiducialCoordinatesF2 = [0,0,0,0]
    FiducialCoordinatesF3 = [0,0,0,0]
    markups.GetNthFiducialWorldCoordinates(0,FiducialCoordinatesF1)
    markups.GetNthFiducialWorldCoordinates(1,FiducialCoordinatesF2)
    markups.GetNthFiducialWorldCoordinates(2,FiducialCoordinatesF3)

    # Création de la segmentation
    segmentationNode = slicer.mrmlScene.AddNewNodeByClass("vtkMRMLSegmentationNode")
    segmentationNode.CreateDefaultDisplayNodes()
    segmentationNode.SetReferenceImageGeometryParameterFromVolumeNode(DosiFilmImage)

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
    effect.setParameter("MinimumThreshold",str(20000))
    effect.setParameter("MaximumThreshold",str(30000))
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

 
    # Boucle de calcul des centres pour les 7 blocs (segment)
    n = slicer.util.getNode('Segmentation')
    s = n.GetSegmentation()
    ss = s.GetSegment("IrradiatedBlocks")
    pd = ss.GetRepresentation('Closed surface')
    com = vtk.vtkCenterOfMass()
    com.SetInputData(pd)
    com.Update()
    CenterOfBlock = com.GetCenter() # CenterOfBlock est alors un tuple avec plusieurs variables (coordonées x,y,z)
    XaxisCenterOfBlock = (CenterOfBlock[0]) # Sélection de la 2ème valeur du tuple (indice 1) qui est la valeur dans l'axe Y qui est l'unique valeure d'intérêt
    YaxisCenterOfBlock = (CenterOfBlock[1])
    YaxisCenterOfBlock = abs(YaxisCenterOfBlock) # On passe en valeur absolue
    XaxisCenterOfBlock = abs(XaxisCenterOfBlock) # On passe en valeur absolue


    ######### Calcul des éléments d'intérêt ###########
    VerticalOffsetInMm = (abs(((abs(FiducialCoordinatesF1[1]) + abs(FiducialCoordinatesF2[1]))/2) - YaxisCenterOfBlock))*0.3528
    VerticalDistanceInMm = (abs(abs(FiducialCoordinatesF1[1]) - abs(FiducialCoordinatesF2[1])))* 0.3528
    LateralOffsetInMm = (abs(FiducialCoordinatesF3[0]) - XaxisCenterOfBlock)* 0.3528

    ######### Enonciation des résultats ###########
    print(u"\n\nRésultats du test -Real and Virtual Isocentre-\n" )
    print("Machine : " + str(machineName) + "\n")
    print(u"Coordonnée X du centre du bloc d'irradiation: " + str(XaxisCenterOfBlock) + "\n")
    print(u"Coordonnée Y du centre du bloc d'irradiation: " + str(YaxisCenterOfBlock) + "\n")
    print(u"Coordonnée du fiducial F1: " + str(FiducialCoordinatesF1) + "\n")
    print(u"Coordonnée du fiducial F1: " + str(FiducialCoordinatesF2) + "\n")
    print(u"Coordonnée du fiducial F1: " + str(FiducialCoordinatesF3) + "\n\n")
    print(u"Vertical Offset (mm): " + str(VerticalOffsetInMm) + "\n")
    print(u"Vertical Distance (mm): " + str(VerticalDistanceInMm) + "\n")
    print(u"Lateral Offset (mm): " + str(LateralOffsetInMm) + "\n")

    ######### Calcul de la conformité et mention des résultats ###########
    if  0 <= VerticalOffsetInMm < 1:
        ResultVerticalOffset = "Conforme"
    elif VerticalOffsetInMm > 1:
        ResultVerticalOffset = "Hors tolerance"
    else:
        ResultVerticalOffset = "Limite" #car si pas < ou > à 1 alors = à 1
    
    print(u"Résultat pour Vertical Offset: " + str(ResultVerticalOffset))


    if  0 <= VerticalDistanceInMm < 1:
        ResultVerticalDistance = "Conforme"
    elif VerticalDistanceInMm > 1:
        ResultVerticalDistance = "Hors tolerance"
    else:
        ResultVerticalDistance = "Limite" #car si pas < ou > à 1 alors = à 1
    
    print(u"Résultat pour Vertical Distance: " + str(ResultVerticalDistance))


    if  0 <= LateralOffsetInMm < 1:
        ResultLateralOffset = "Conforme"
    elif LateralOffsetInMm > 1:
        ResultLateralOffset = "Hors tolerance"
    else:
        ResultLateralOffset = "Limite" #car si pas < ou > à 1 alors = à 1
    
    print(u"Résultat pour Lateral Offset: " + str(ResultLateralOffset) + "\n")

    """### Au cas ou problème dans l'image ##
    if  VerticalOffsetInMm or VerticalDistanceInMm or LateralOffsetInMm < 0:
      logging.info(u"Valeur de la différence négative, problème dans l'image ou dans le programme, contactez Aurélien Corroyer-Dulmont tel : 5768")"""

    
    logging.info('Processing completed')

    return True