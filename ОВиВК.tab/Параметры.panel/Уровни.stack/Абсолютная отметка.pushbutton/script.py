#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Абсолютная отметка'
__doc__ = "Обновляет значения параметров ADSK_Отметка оси от нуля и ADSK_Отметка низа от нуля"

import os.path as op
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
import sys
import paraSpec
from Autodesk.Revit.DB import *
from Redomine import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output

colFittings = make_col(BuiltInCategory.OST_DuctFitting)
colPipeFittings = make_col(BuiltInCategory.OST_PipeFitting)
colPipeCurves = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colFlexCurves = make_col(BuiltInCategory.OST_FlexDuctCurves)
colFlexPipeCurves = make_col(BuiltInCategory.OST_FlexPipeCurves)
colTerminals = make_col(BuiltInCategory.OST_DuctTerminal)
colAccessory = make_col(BuiltInCategory.OST_DuctAccessory)
colPipeAccessory = make_col(BuiltInCategory.OST_PipeAccessory)
colEquipment = make_col(BuiltInCategory.OST_MechanicalEquipment)
colInsulations = make_col(BuiltInCategory.OST_DuctInsulations)
colPipeInsulations = make_col(BuiltInCategory.OST_PipeInsulations)
colPlumbingFixtures = make_col(BuiltInCategory.OST_PlumbingFixtures)
colSprinklers = make_col(BuiltInCategory.OST_Sprinklers)

collections = [colCurves, colPipeCurves, colSprinklers, colAccessory,
               colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]

doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView


levelCol = make_col(BuiltInCategory.OST_Levels)

class level:
    def __init__(self, element):
        self.name = element.Name
        self.z = element.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()

report_rows = []
def isElementEditedBy(element):
    try:
        edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
    except Exception:
        edited_by = __revit__.Application.Username

    if edited_by and edited_by != __revit__.Application.Username:
        if edited_by not in report_rows:
            report_rows.append(edited_by)
        return True
    return False

def getElementLevelName(element):
    level = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
    baseLevel = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM)

    if element.Category.IsId(BuiltInCategory.OST_PipeInsulations) \
            or element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
        return None
    else:
        if level:
            level = level.AsValueString()
        if baseLevel:
            level = baseLevel.AsValueString()

    return level

def getElementBot(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        bot = element.get_Parameter(BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION).AsValueString()
        return bot

    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        bot = element.get_Parameter(BuiltInParameter.RBS_PIPE_BOTTOM_ELEVATION).AsValueString()
        return bot

    bot = element.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsValueString()
    return bot



def getElementMid(element):
    if element.Category.IsId(BuiltInCategory.OST_PipeCurves) or element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        mid = element.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsValueString()
        return mid
    else:
        return None

def summ(x, y):
    if ',' in x:
        x = x.replace(',', '.')
    if ',' in y:
        y = y.replace(',', '.')
    return float(x) + float(y)

def execute():
    with revit.Transaction("Обновление абсолютной отметки"):
        projectLevels = []
        for element in levelCol:
            newLevel = level(element)
            projectLevels.append(newLevel)

        for collection in collections:
            for element in collection:
                if not isElementEditedBy(element):
                    midParam = element.LookupParameter('ФОП_ВИС_Отметка оси от нуля')
                    botParam = element.LookupParameter('ФОП_ВИС_Отметка низа от нуля')
                    elementLV = getElementLevelName(element)
                    elementMid = getElementMid(element)
                    elementBot = getElementBot(element)




                    for projectLevel in projectLevels:
                        if projectLevel.name == elementLV:

                            if midParam:
                                midParam.Set(0)
                            if midParam and elementMid:
                                markMid = summ(projectLevel.z, elementMid)
                                midParam.Set(markMid/1000)

                            if botParam:
                                botParam.Set(0)
                            if botParam and elementBot:
                                markBot = summ(projectLevel.z, elementBot)
                                botParam.Set(markBot/1000)

                    try:
                        if element.Host:
                            elementBot = fromRevitToMilimeters(element.Location.Point[2])
                            botParam.Set(elementBot/1000)
                    except:
                        pass

    for report in report_rows:
        print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report


if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

parametersAdded = paraSpec.check_parameters()

if not parametersAdded:
    execute()

