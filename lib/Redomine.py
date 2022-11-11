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
    replaceDef = isReplasing(replaceName)
    if replaceDef:
            paraName = replaceIsValid(element, paraName, replaceName)
    try:
        ADSK_Izm = element.LookupParameter(paraName).AsString()
    except Exception:
        try:
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)

            if ElemType.LookupParameter(paraName) == None:
                ADSK_Izm = "None"
            else:
                ADSK_Izm = ElemType.LookupParameter(paraName).AsString()
        except Exception:
            print 'Невозможно получить параметр единицы измерения'
            sys.exit()
    return ADSK_Izm

def get_ADSK_Name(element):
    paraName = 'ADSK_Наименование'
    replaceName = 'ФОП_ВИС_Замена параметра_Наименование'
    replaceDef = isReplasing(replaceName)
    if replaceDef:
            paraName = replaceIsValid(element, paraName, replaceName)
    try:
        if element.LookupParameter(paraName):
            ADSK_Name = element.LookupParameter(paraName).AsString()
            if ADSK_Name == None or ADSK_Name == "":
                ADSK_Name = "None"
        else:
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)

            if ElemType.LookupParameter(paraName) == None\
                    or ElemType.LookupParameter(paraName).AsString() == None \
                    or ElemType.LookupParameter(paraName).AsString() == "":
                ADSK_Name = "None"
            else:
                ADSK_Name = ElemType.LookupParameter(paraName).AsString()
    except Exception:
        print 'Невозможно получить параметр наименования'
        sys.exit()

    if str(element.Category.Name) == 'Трубы':
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента'):
            if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента').AsInteger() == 1:
                ADSK_Name = element.LookupParameter('Описание сегмента').AsString()

    return ADSK_Name

def get_ADSK_Maker(element):
    paraName = 'ADSK_Завод-изготовитель'
    replaceName = 'ФОП_ВИС_Замена параметра_Завод-изготовитель'
    replaceDef = isReplasing(replaceName)
    if replaceDef:
            paraName = replaceIsValid(element, paraName, replaceName)

    try:
        if element.LookupParameter(paraName):
            ADSK_Maker = element.LookupParameter(paraName).AsString()
            if ADSK_Maker == None or ADSK_Maker == "":
                ADSK_Maker = ""
        else:
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)

            if ElemType.LookupParameter(paraName) == None\
                    or ElemType.LookupParameter(paraName).AsString() == None \
                    or ElemType.LookupParameter(paraName).AsString() == "":
                ADSK_Maker = ""
            else:
                ADSK_Maker = ElemType.LookupParameter(paraName).AsString()
    except Exception:
        print 'Невозможно получить параметр завода-изготовителя'
        sys.exit()

    return ADSK_Maker

def get_ADSK_Mark(element):
    paraName = 'ADSK_Марка'
    replaceName = 'ФОП_ВИС_Замена параметра_Марка'
    replaceDef = isReplasing(replaceName)
    if replaceDef:
            paraName = replaceIsValid(element, paraName, replaceName)

    try:
        if element.LookupParameter(paraName):
            ADSK_Mark = element.LookupParameter(paraName).AsString()
        else:
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)

            if ElemType.LookupParameter(paraName) == None \
                    or ElemType.LookupParameter(paraName).AsString() == None \
                    or ElemType.LookupParameter(paraName).AsString() == "":
                ADSK_Mark = "None"
            else:
                ADSK_Mark = ElemType.LookupParameter(paraName).AsString()
    except Exception:
        print 'Невозможно получить параметр марки'
        sys.exit()

    return ADSK_Mark

def get_ADSK_Code(element):
    paraName = 'ADSK_Код изделия'
    replaceName = 'ФОП_ВИС_Замена параметра_Код изделия'
    replaceDef = isReplasing(replaceName)
    if replaceDef:
            paraName = replaceIsValid(element, paraName, replaceName)

    if element.LookupParameter(paraName):
        ADSK_Code = element.LookupParameter(paraName).AsString()
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter(paraName) == None \
                or ElemType.LookupParameter(paraName).AsString() == None \
                or ElemType.LookupParameter(paraName).AsString() == "":
            ADSK_Code = ""
        else:
            ADSK_Code = ElemType.LookupParameter(paraName).AsString()
    return ADSK_Code




