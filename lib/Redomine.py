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

replacement = True

def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col


'ФОП_ВИС_Замена параметра_Единица измерения'
'ФОП_ВИС_Замена параметра_Завод-изготовитель'
'ФОП_ВИС_Замена параметра_Код изделия'
'ФОП_ВИС_Замена параметра_Марка'
'ФОП_ВИС_Замена параметра_Наименование'


def get_ADSK_Izm(element):
    Izm = 'ADSK_Единица измерения'
    #replacement_Izm = doc.ProjectInformation.LookupParameter('ФОП_ВИС_Замена параметра_Единица измерения').AsString()
    #if replacement_Izm != None or replacement_Izm != '':
        #Izm = 'ФОП_ВИС_Замена параметра_Единица измерения'
        #if not element.LookupParameter(Izm):
            #print 'Назначеного параметра замены единицы измерения нет в проекте'
            #replacement = False
            #return None
    try:
        ADSK_Izm = element.LookupParameter(Izm).AsString()
    except Exception:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter(Izm) == None:
            ADSK_Izm = "None"
        else:
            ADSK_Izm = ElemType.LookupParameter(Izm).AsString()
    return ADSK_Izm

def get_ADSK_Name(element):
    if element.LookupParameter('ADSK_Наименование'):
        ADSK_Name = element.LookupParameter('ADSK_Наименование').AsString()
        if ADSK_Name == None or ADSK_Name == "":
            ADSK_Name = "None"
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Наименование') == None\
                or ElemType.LookupParameter('ADSK_Наименование').AsString() == None \
                or ElemType.LookupParameter('ADSK_Наименование').AsString() == "":
            ADSK_Name = "None"
        else:
            ADSK_Name = ElemType.LookupParameter('ADSK_Наименование').AsString()

    if str(element.Category.Name) == 'Трубы':
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента'):
            if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента').AsInteger() == 1:
                ADSK_Name = element.LookupParameter('Описание сегмента').AsString()



    return ADSK_Name

def get_ADSK_Maker(element):
    if element.LookupParameter('ADSK_Завод-изготовитель'):
        ADSK_Maker = element.LookupParameter('ADSK_Завод-изготовитель').AsString()
        if ADSK_Maker == None or ADSK_Maker == "":
            ADSK_Maker = ""
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Завод-изготовитель') == None\
                or ElemType.LookupParameter('ADSK_Завод-изготовитель').AsString() == None \
                or ElemType.LookupParameter('ADSK_Завод-изготовитель').AsString() == "":
            ADSK_Maker = ""
        else:
            ADSK_Maker = ElemType.LookupParameter('ADSK_Завод-изготовитель').AsString()
    return ADSK_Maker

def get_ADSK_Mark(element):
    if element.LookupParameter('ADSK_Марка'):
        ADSK_Mark = element.LookupParameter('ADSK_Марка').AsString()
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Марка') == None \
                or ElemType.LookupParameter('ADSK_Марка').AsString() == None \
                or ElemType.LookupParameter('ADSK_Марка').AsString() == "":
            ADSK_Mark = "None"
        else:
            ADSK_Mark = ElemType.LookupParameter('ADSK_Марка').AsString()
    return ADSK_Mark

def get_ADSK_Code(element):
    if element.LookupParameter('ADSK_Код изделия'):
        ADSK_Code = element.LookupParameter('ADSK_Код изделия').AsString()
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Код изделия') == None \
                or ElemType.LookupParameter('ADSK_Код изделия').AsString() == None \
                or ElemType.LookupParameter('ADSK_Код изделия').AsString() == "":
            ADSK_Code = ""
        else:
            ADSK_Code = ElemType.LookupParameter('ADSK_Код изделия').AsString()
    return ADSK_Code


def replacement_status():
    if replacement:
        return True
    else:
        return False


