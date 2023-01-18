#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Маркировка'
__doc__ = "Маркирует элементы систем"


import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit



from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *


doc = __revit__.ActiveUIDocument.Document # type: Document

uidoc = __revit__.ActiveUIDocument

view = doc.ActiveView


class elementOfBranch:
    def __init__(self):
        self.numberInLine = 0
        self.element = None
        self.width = 0
        self.height = 0
        self.diameter = 0
        self.flow = 0
        self.connectedElements = []
        self.XYZ = None

class branch:
    def __init__(self):
        self.curves = []
        self.fittings = []
        self.elementCount = 0

def getConnectors(element):
    connectors = []
    try:
        a = element.ConnectorManager.Connectors.ForwardIterator()
        while a.MoveNext():
            connectors.append(a.Current)
    except:
        try:
            a = element.MEPModel.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
        except:
            a = element.MEPSystem.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
    return connectors

def pick_elements(uidoc):
    selectedIds = uidoc.Selection.GetElementIds()
    if 0 == selectedIds.Count:
        print 'Выберите один элемент воздуховода или трубопровода'
        sys.exit()

    if selectedIds.Count > 1:
        print 'Выберите один элемент воздуховода или трубопровода'
        sys.exit()

    result = doc.GetElement(selectedIds[0])
    if selectedIds.Count == 1 and not isDuctOrPipe(result):
        print 'Выберите только одну секцию воздуховода или трубопровода'
        sys.exit()

    return result

def isDuctOrPipe(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves) or element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return True
    else:
        return False

def isDuctOrPipeConnector(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting) or element.Category.IsId(BuiltInCategory.OST_PipeFitting):
        return True
    else:
        return False

def isInsulation(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctInsulations) or element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
        return True

def isSystem(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctSystem) or element.Category.IsId(BuiltInCategory.OST_PipingSystem):
        return True

def getConnected(element):
    connected = []
    connectors = getConnectors(element)
    for connector in connectors:
        for ref in connector.AllRefs:
            if ref.Owner.Id not in trashElements:
                trashElements.append(ref.Owner.Id)
                if not isInsulation(ref.Owner) and ref.Owner.Id != element.Id and not isSystem(ref.Owner):
                    connected.append(ref.Owner)
    return connected

def optimizeList(list):
    calcList = []
    for element in list:
        if element not in calcList:
            calcList.append(element)
    return calcList

def findCurves(element, list):
    connected = getConnected(element)
    for el in connected:
        if isDuctOrPipe(el):
            list.append(el)
        else:
            list = list + findCurves(el, list)
    return list

def getConnectedCurves(curve):

    nextCurves = []  # type: List
    for element in curve.connectedElements:
        nextCurves = findCurves(element, nextCurves)

    if len(nextCurves) == 0:
        return None
    nextCurves = optimizeList(nextCurves)
    nextCurve = makeCurve(nextCurves[0], number=(curve.numberInLine + 1))  # type: elementOfBranch

    if len(nextCurves) > 1:
        nextCurves.pop(0)

        for curve in nextCurves:
            branchesStarts.append(curve)

    return nextCurve

def makeCurve(element, number):
    trashElements.append(element.Id)
    curve = elementOfBranch()  # type: elementOfBranch
    curve.element = element

    XYZ = element.Location.Curve.Origin

    start = element.Location.Curve.GetEndPoint(0)
    end = element.Location.Curve.GetEndPoint(1)

    midX = (start[0] + end[0])/2
    midY = (start[1] + end[1])/2
    midZ = (start[2] + end[2])/2

    pt = Autodesk.Revit.DB.XYZ(midX,midY,midZ)


    curve.XYZ = pt

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        curve.flow = element.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)

    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        curve.flow = element.GetParamValue(BuiltInParameter.RBS_PIPE_FLOW_PARAM)

    try:
        curve.width = element.Width
        curve.height = element.Height

    except:
        curve.diameter = element.Diameter

    connectors = getConnectors(element)

    for connector in connectors:
        for ref in connector.AllRefs:
            if not isInsulation(ref.Owner) and ref.Owner.Id != element.Id and not isSystem(ref.Owner):
                if ref.Owner not in trashElements:
                    curve.connectedElements.append(ref.Owner)
                    trashElements.append(ref.Owner.Id)
    curve.numberInLine = number

    return curve

def isGoOn(branch, curve):
    nextCurve = getConnectedCurves(curve)

    if nextCurve:
        branch.curves.append(nextCurve)
        isGoOn(branch, nextCurve)
    return branch

def buildBranch(element):
    newBranch = branch() # type: branch
    curve = makeCurve(element, number = 0) # type: elementOfBranch
    newBranch.curves.append(curve)
    newBranch = isGoOn(newBranch, curve)
    branches.append(newBranch)

def isSizeEqual(curve_1, curve_2):
    if curve_1.width != curve_2.width:
        return False

    if curve_1.height != curve_2.height:
        return False

    if curve_1.diameter != curve_2.diameter:
        return False

    return True

def isFlowEqual(curve_1, curve_2):
    pass

def isLevelEqual(curve_1, curve_2):
    pass

def placeMark(curve):

    ref = Reference(curve.element)
    tagMode = TagMode.TM_ADDBY_CATEGORY
    tagorn = TagOrientation.Horizontal


    IndependentTag.Create(doc, doc.ActiveView.Id, ref, True, tagMode, tagorn, curve.XYZ)


def compareCurves(branch, bySize, byFlow, byLevel):
    sizeStatus = True
    flowStatus = True
    levelStatus = True

    if len(branch.curves) == 0:
        return None

    if len(branch.curves) == 1:
        placeMark(branch.curves[0])


    curvesToMark = []

    for curve in branch.curves:
        index = curve.numberInLine

        try:
            compareCurve = branch.curves[index + 1]
        except:
            compareCurve = branch.curves[index - 1]

        if bySize:
            sizeStatus = isSizeEqual(curve, compareCurve)

        if not sizeStatus or not flowStatus or not levelStatus:
            if curve not in curvesToMark:
                curvesToMark.append(curve)
            if compareCurve not in curvesToMark:
                curvesToMark.append(compareCurve)
    for curve in curvesToMark:
        placeMark(curve)






branchesStarts = [] #список элементов-начал ответвлений

branches = [] #список обработанных ветвей

trashElements = [] #список уже обработных соеденителей которые не должны использоваться снова



def execute(bySize = False, byFlow = False, byLevel = False):

    #циклично перебираем начиная с стартового отрезка элементы и собираем в список ответвлений
    selectedIds = uidoc.Selection.GetElementIds()
    startElement = pick_elements(uidoc) #элемент-начало ответвления
    branchesStarts.append(startElement)

    while len(branchesStarts) > 0:
        buildBranch(branchesStarts[0])
        branchesStarts.pop(0)



    #выбираем марку
    mark = FilteredElementCollector(doc)
    if startElement.Category.IsId(BuiltInCategory.OST_DuctCurves):
        mark.OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_DuctTags).FirstElement()
    if startElement.Category.IsId(BuiltInCategory.OST_PipeCurves):
        mark.OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_GenericAnnotation).FirstElement()



    mark = doc.GetElement(ElementId(int(1673226)))
    # for el in mark:
    #     print el.Id
    #     mark = el

    with revit.Transaction("Маркировка элементов"):
        for branch in branches:
            #сравниваем и маркируем соседние участки воздуховодов
            compareCurves(branch, bySize, byFlow, byLevel)






def markBySizeAndFlow():
    execute(bySize=True,byFlow=True)

def markBySize():
    execute(bySize=True)



