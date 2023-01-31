#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Библиотека стандартных функций'
__doc__ = "pass"

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
from dosymep.Revit.Geometry import *
import sys
import paraSpec
from Autodesk.Revit.DB import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output
from pyrevit import script


doc = __revit__.ActiveUIDocument.Document  # type: Document
import os.path as op
import clr

output = script.get_output()


def getParameter(element, paraName):
    try:
        parameter = element.LookupParameter(paraName)
        return parameter
    except:
        return None

def setIfNotRO(parameter, value):
    if not parameter.IsReadOnly:
        parameter.Set(value)



def getDefCols():
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

    collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves, colInsulations,
                   colPipeInsulations, colPipeCurves, colSprinklers, colAccessory,
                   colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]

    return collections




output = script.get_output()

def isItFamily():
    try:
        manager = doc.FamilyManager
        return True
    except Exception:
        pass

def lookupCheck(element, paraName, isExit = True):
    type = 'экземпляра '
    try:
        if element.GetTypeId():
            type = 'типа '
    except:
        pass

    parameter = getParameter(element, paraName)

    if parameter:
        return parameter
    else:
        if isExit:
            print 'Параметр ' + type + paraName + ' не назначен для категории ' + element.Category.Name + ' (ID элемента на котором найдена ошибка ' + output.linkify(element.Id) +")"
            sys.exit()
        else:
            return None
def getSharedParameter (element, paraName, replaceName):
    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        if paraName == 'ADSK_Наименование':
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)
            if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента').AsInteger() == 1:

                parameter = element.LookupParameter('Описание сегмента').AsString()
                return parameter

    report = paraName.split('_')[1]
    replaceDef = isReplasing(replaceName)
    name = paraName
    if replaceDef:
            name = replaceIsValid(element, paraName, replaceName)
    try:
        parameter = element.LookupParameter(name).AsString()
        if parameter == None:
            parameter = 'None'
    except Exception:
        #try:

        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter(name) == None:
            parameter = "None"
        else:
            parameter = ElemType.LookupParameter(name).AsString()

    nullParas = ['ADSK_Завод-изготовитель', 'ADSK_Марка', 'ADSK_Код изделия']

    if not parameter:
        parameter = 'None'

    if parameter == 'None' or parameter == None:
        if paraName in nullParas:
            parameter = ''
        else:
            parameter = 'None'

    parameter = str(parameter)

    return parameter

def fromRevitToMeters(number):
    meters = (number * 304.8)/1000
    return meters

def fromRevitToMilimeters(number):
    mms = number * 304.8
    return mms
def fromRevitToSquareMeters(number):
    square = number * 0.092903
    return square


def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

def isReplasing(replaceName ):
    replaceDefiniton = doc.ProjectInformation.LookupParameter(replaceName).AsString()
    if replaceDefiniton == None:
        return False
    if replaceDefiniton == 'None':
        return False
    if replaceDefiniton == '':
        return False
    else:
        return replaceDefiniton

def replaceIsValid (element, paraName, replaceName):
    replaceparameter = doc.ProjectInformation.LookupParameter(replaceName).AsString()
    if not element.LookupParameter(replaceparameter):
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if not ElemType.LookupParameter(replaceparameter):
            print output.linkify(element.Id)
            print 'Назначеного параметра замены ' + replaceparameter + ' нет у одной из категорий, назначенных для исходного параметр ' + paraName
            sys.exit()

    return replaceparameter

def get_ADSK_Izm(element):
    paraName = 'ADSK_Единица измерения'
    replaceName = 'ФОП_ВИС_Замена параметра_Единица измерения'
    ADSK_Izm = getSharedParameter(element, paraName, replaceName)
    return ADSK_Izm

def get_ADSK_Name(element):
    paraName = 'ADSK_Наименование'
    replaceName = 'ФОП_ВИС_Замена параметра_Наименование'
    ADSK_Name = getSharedParameter(element, paraName, replaceName)
    return ADSK_Name

def get_ADSK_Maker(element):
    paraName = 'ADSK_Завод-изготовитель'
    replaceName = 'ФОП_ВИС_Замена параметра_Завод-изготовитель'
    ADSK_Maker = getSharedParameter(element, paraName, replaceName)
    return ADSK_Maker

def get_ADSK_Mark(element):
    paraName = 'ADSK_Марка'
    replaceName = 'ФОП_ВИС_Замена параметра_Марка'
    ADSK_Mark = getSharedParameter(element, paraName, replaceName)
    return ADSK_Mark

def get_ADSK_Code(element):
    paraName = 'ADSK_Код изделия'
    replaceName = 'ФОП_ВИС_Замена параметра_Код изделия'
    ADSK_Code = getSharedParameter(element, paraName, replaceName)
    return ADSK_Code
