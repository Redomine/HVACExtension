#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление параметров'
__doc__ = "Добавление параметров в семейство для выдачи заданий электрикам"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
from System import Guid
from pyrevit import revit
from Autodesk.Revit.DB.Electrical import *
import os



doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView



def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
connectorCol = make_col(BuiltInCategory.OST_ConnectorElem)
loadsCol = make_col(BuiltInCategory.OST_ElectricalLoadClassifications)

try:
    manager = doc.FamilyManager
except Exception:
    print "Надстройка предназначена для работы с семействами"
    sys.exit()

def associate(param, famparam):
    manager.AssociateElementParameterToFamilyParameter(param, famparam)



set = doc.FamilyManager.Parameters



paraNames = ['ADSK_Полная мощность', 'ADSK_Коэффициент мощности', 'ADSK_Количество фаз', 'ADSK_Напряжение',
             'ADSK_Классификация нагрузок',  'ADSK_Номинальная мощность', 'mS_Имя нагрузки']

paraNames_V1 = ['ФОП_ВИС_Нагреватель или шкаф', 'ФОП_ВИС_Частотный регулятор', 'ФОП_ВИС_Мощность нагревателя', 'ФОП_ВИС_Напряжение нагревателя']

notFormula = ['ADSK_Полная мощность', 'ADSK_Коэффициент мощности', 'ADSK_Количество фаз']


def update_fop(version):

    if version == "ADSK":
        #проверяем тот ли файл общих параметров подгружен
        spFileName = str(doc.Application.SharedParametersFilename)
        spFileName = spFileName.split('\\')
        spFileName = spFileName[-1]
        if "ФОП_ADSK.txt" != spFileName:
            try:
                doc.Application.SharedParametersFilename = str(os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\ФОП_ADSK.txt"
            except Exception:
                print 'По стандартному пути не найден файл общих параметров, обратитесь в бим отдел или замените вручную на ФОП_ADSK.txt'
    else:
        spFileName = str(doc.Application.SharedParametersFilename)
        spFileName = spFileName.split('\\')
        spFileName = spFileName[-1]

        try:
            doc.Application.SharedParametersFilename = str(
                os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\ФОП_v1.txt"
        except Exception:
            print 'По стандартному пути не найден файл общих параметров, обратитесь в бим отдел или замените вручную на ФОП_v1.txt'

    spFile = doc.Application.OpenSharedParameterFile()
    return spFile



connectorNum = 0
try:
    for connector in connectorCol:
        if str(connector.Domain) == "DomainElectrical":
            connectorNum = connectorNum + 1
except Exception:
    print "Не найдено электрических коннекторов, должен быть один"
    sys.exit()


if connectorNum > 1:
    print "Электрических коннекторов больше одного, удалите лишние"
    sys.exit()
if connectorNum == 0:
    print "Не найдено электрических коннекторов, должен быть один"
    sys.exit()



with revit.Transaction("Добавление параметров"):
    #если в семействе нет никаких типоразмеров, ревит почему-то не даст создать формулы. Создаем хотя бы один тип
    typeNumber = 0
    for type in doc.FamilyManager.Types:
        typeNumber = typeNumber + 1
    if typeNumber < 1:
        doc.FamilyManager.NewType('Стандарт')


    #удаляем формулы на элементах, к которым будут присвоены свои, чтоб избежать конфликтов
    for param in set:
        if str(param.Definition.Name) in notFormula:
            manager.SetFormula(param, "1")

    #проверяем наличие параметров из списка имен в проекте, присваиваем им параметры типа или экземпляра
    for param in set:
        if str(param.Definition.Name) in paraNames:

            #мощность и напряжение иногда уже проставлены и могут быть любого типа, просто их не трогаю
            if str(param.Definition.Name) == 'ADSK_Номинальная мощность' or str(param.Definition.Name) == 'ADSK_Напряжение':
               pass
            else:
                manager.MakeInstance(param)
            paraNames.remove(param.Definition.Name)

        if str(param.Definition.Name) in paraNames_V1:
            manager.MakeInstance(param)
            paraNames_V1.remove(param.Definition.Name)



    #если в списке имен после проверки осталось что-то, добавляем параметры из списка
    if len(paraNames) > 0:
        spFile = update_fop("ADSK")
        for name in paraNames:
            for dG in spFile.Groups:
                group = "04 Обязательные ИНЖЕНЕРИЯ"
                if name == 'mS_Имя нагрузки': group = "mySchema"
                if str(dG.Name) == group:
                    myDefinitions = dG.Definitions
                    eDef = myDefinitions.get_Item(name)
                    if name == 'ADSK_Номинальная мощность' or name == 'ADSK_Напряжение' or name == 'ADSK_Классификация нагрузок':
                        manager.AddParameter(eDef, BuiltInParameterGroup.PG_ELECTRICAL_LOADS, False)
                    else:
                        manager.AddParameter(eDef, BuiltInParameterGroup.PG_ELECTRICAL_LOADS, True)

with revit.Transaction("Добавление параметров"):
    spFile = update_fop("v1")
    for dG in spFile.Groups:
        for name in paraNames_V1:

            group = "03_ВИС"
            if str(dG.Name) == group:
                myDefinitions = dG.Definitions
                eDef = myDefinitions.get_Item(name)
                manager.AddParameter(eDef, BuiltInParameterGroup.PG_ELECTRICAL_LOADS, False)









