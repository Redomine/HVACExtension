#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление\nфункции'
__doc__ = "Обновляет экономическую функцию"

import os.path as op

import clr

import Redomine

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

Redomine.getDocIfItsWorkshared()
view = doc.ActiveView


def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

collections = getDefCols()
colGeneric = make_col(BuiltInCategory.OST_GenericModel)
collections.append(colGeneric)

colDuctSystems = make_col(BuiltInCategory.OST_DuctSystem)
colPipeSystems = make_col(BuiltInCategory.OST_PipingSystem)


report_rows = set()



def getEFsystem(element):
    if element not in colGeneric:
        sys_name = element.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
    else:
        sys_name = None
    EF = None

    if sys_name != None:
        if element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment) or element.Category.IsId(BuiltInCategory.OST_PipeAccessory)\
                or element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            sys_name = element.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
            sys_name = sys_name.split(',')
            sys_name = sys_name[0]

        if sys_name in ductDict:
            EF = ductDict[sys_name]
        if sys_name in pipeDict:
            EF = pipeDict[sys_name]



    if str(EF) == 'None':
        if not infEF:
            EF = 'None'
        else:
            EF = infEF

    return EF

def copyEF(collection):
    for element in collection:
        if not isElementEditedBy(element):
            EF = getEFsystem(element)
            if EF != None:
                ElemTypeId = element.GetTypeId()
                ElemType = doc.GetElement(ElemTypeId)

                typeEF = None

                try:
                    typeEF = ElemType.LookupParameter('ФОП_ВИС_Экономическая функция').AsString()
                except:
                    pass

                if typeEF:
                    if str(typeEF) != 'None' or typeEF != "":
                        EF = typeEF

                parameter = lookupCheck(element, 'ФОП_Экономическая функция')

                if element not in colGeneric:
                    setIfNotRO(parameter, EF)
                if element in colGeneric:
                    if "_Якорный" not in element.Symbol.FamilyName:
                        setIfNotRO(parameter, EF)

        else:
            fillReportRows(element, report_rows)



def getDependent(collection):
    for element in collection:
        if not isElementEditedBy(element):
            EF = lookupCheck(element, 'ФОП_Экономическая функция').AsString()
            dependent = None
            try:
                dependent = element.GetSubComponentIds()
            except:
                pass


            if dependent:
                for depend in dependent:
                    parameter = lookupCheck(doc.GetElement(depend), 'ФОП_Экономическая функция')
                    setIfNotRO(parameter, EF)
                    # for list in collections:
                    #     if depend in list:
                    #         parameter = lookupCheck(doc.GetElement(depend), 'ФОП_Экономическая функция')
                    #         setIfNotRO(parameter, EF)





def getSystemDict(collection):
    Dict = {}
    for system in collection:
        if system.Name not in Dict:
            ElemTypeId = system.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)

            typeEF = lookupCheck(ElemType, 'ФОП_ВИС_ЭФ для системы', isExit = False)

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



if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

status = paraSpec.check_parameters()

if not status:
    information = doc.ProjectInformation

    infEF = lookupCheck(information, 'ФОП_Экономическая функция').AsString()

    if lookupCheck(information, 'ФОП_Экономическая функция').AsString():
        if lookupCheck(information, 'ФОП_Экономическая функция').AsString() == '':
            infEF = None



    ductDict = getSystemDict(colDuctSystems)
    pipeDict = getSystemDict(colPipeSystems)


    with revit.Transaction("Обновление экономической функции"):
        for collection in collections:
            copyEF(collection)

    with revit.Transaction("Проверка вложенных"):
        for collection in collections:
            getDependent(collection)

