#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление ФОП ЭФ'
__doc__ = "Обновляет экономическую функцию"

import os.path as op

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
import sys
import paraSpec
from Autodesk.Revit.DB import *
from System import Guid
from pyrevit import revit
from Redomine import *
#test

doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView


def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

collections = getDefCols()


colDuctSystems = make_col(BuiltInCategory.OST_DuctSystem)
colPipeSystems = make_col(BuiltInCategory.OST_PipingSystem)

def getEFsystem(element):

    sys_name = element.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
    EF = None
    if sys_name != None:
        if element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
            sys_name = element.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
            sys_name = sys_name.split(',')
            sys_name = sys_name[0]

        if sys_name in ductDict:
            EF = ductDict[sys_name]
        if sys_name in pipeDict:
            EF = pipeDict[sys_name]



    if str(EF) == 'None':

        EF = lookupCheck(information, 'ФОП_Экономическая функция').AsString()

    return EF

def copyEF(collection):
    for element in collection:

        EF = getEFsystem(element)

        if EF != None:
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)

            typeEF = lookupCheck(ElemType, 'ФОП_ВИС_Экономическая функция').AsString()

            if typeEF:
                if str(typeEF) != 'None' or typeEF != "":
                    EF = typeEF


            parameter = lookupCheck(element, 'ФОП_Экономическая функция')
            setIfNotRO(parameter, EF)


def getDependent(collection):
    for element in collection:
        EF = lookupCheck(element, 'ФОП_Экономическая функция').AsString()

        dependent = None
        try:
            dependent = element.GetSubComponentIds()
        except:
            pass

        if dependent:
            for depend in dependent:
                for list in collections:
                    if depend in list:
                        parameter = lookupCheck(doc.GetElement(depend), 'ФОП_Экономическая функция')
                        setIfNotRO(parameter, EF)




def getSystemDict(collection):
    Dict = {}
    for system in collection:
        if system.Name not in Dict:
            ElemTypeId = system.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)

            typeEF = lookupCheck(ElemType, 'ФОП_ВИС_Экономическая функция', isExit = False)


            if typeEF:
                if typeEF != None:
                    if typeEF.AsString() != "":
                        EF = typeEF.AsString()
                        Dict[system.Name] = EF
            else:
                systemEF = lookupCheck(system, 'ФОП_ВИС_Экономическая функция')
                if systemEF.AsString() != None:
                    if systemEF.AsString() != "":
                        EF = systemEF.AsString()
                        Dict[system.Name] = EF
    return Dict

status = paraSpec.check_parameters()

if not status:
    information = doc.ProjectInformation
    try:
        if lookupCheck(information, 'ФОП_Экономическая функция').AsString() == None:
            print 'ФОП_Экономическая функция не заполнен в сведениях о проекте'
            sys.exit()
    except Exception:
        print 'Не найден параметр ФОП_Экономическая функция'
        sys.exit()

    ductDict = getSystemDict(colDuctSystems)

    pipeDict = getSystemDict(colPipeSystems)


    with revit.Transaction("Обновление общей спеки"):
        for collection in collections:
            copyEF(collection)
            getDependent(collection)