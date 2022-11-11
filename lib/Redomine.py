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
import sys
import paraSpec
from Autodesk.Revit.DB import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output


doc = __revit__.ActiveUIDocument.Document  # type: Document
import os.path as op
import clr


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
    if parameter == 'None' or parameter == None:
        if paraName in nullParas:
            parameter = ''
        #except Exception:
        #    print 'Невозможно получить параметр ' + report
        #    sys.exit()
    return parameter
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
    if not element.LookupParameter(replaceName):
        print 'Назначеного параметра замены ' + replaceName + ' нет у одной из категорий, назначенных для исходного параметр ' + paraName
        sys.exit()
    else:
        return paraName

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

