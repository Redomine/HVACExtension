#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление параметров'
__doc__ = "Добавление параметров в семейство для выдачи заданий электрикам"


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
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
from System import Guid
from pyrevit import revit
from Redomine import *
from Autodesk.Revit.DB.Electrical import *
import os



doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

if not isItFamily():
    print 'Надстройка предназначена для работы с семействами'
    sys.exit()

def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

connectorCol = make_col(BuiltInCategory.OST_ConnectorElem)
loadsCol = make_col(BuiltInCategory.OST_ElectricalLoadClassifications)




def associate(param, famparam):
    manager = doc.FamilyManager
    manager.AssociateElementParameterToFamilyParameter(param, famparam)




fullParaList = ['ADSK_Полная мощность', 'ADSK_Коэффициент мощности', 'ADSK_Количество фаз', 'ADSK_Напряжение',
             'ADSK_Номинальная мощность', 'ФОП_ВИС_Нагреватель или шкаф', 'ФОП_ВИС_Частотный регулятор',
                'mS_Имя нагрузки', 'mS_Координация оборудования', 'mS_Имя нагрузки Владельца', 'mS_Имя нагрузки Суффикс']

paraNamesADSK = ['ADSK_Полная мощность', 'ADSK_Коэффициент мощности', 'ADSK_Количество фаз', 'ADSK_Напряжение',
             'ADSK_Номинальная мощность']

paraNamesFOP = ['ФОП_ВИС_Нагреватель или шкаф', 'ФОП_ВИС_Частотный регулятор']
paraNamesMS = ['mS_Имя нагрузки', 'mS_Координация оборудования', 'mS_Имя нагрузки Владельца', 'mS_Имя нагрузки Суффикс']


class familyParam:
    def __init__(self):
        self.Name = None
        self.Param = None
        self.Formula = None
        self.defaultValue = None
        self.connectorParam = None



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


def check_cons():
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

def isFormulaExist(name):
    if name == 'ADSK_Количество фаз':
        return  "if(ADSK_Напряжение < 250 В, 1, 3)"

    if name == 'ADSK_Коэффициент мощности':
        return 'if(ФОП_ВИС_Частотный регулятор, 0.95, if(ФОП_ВИС_Нагреватель или шкаф, 1, if(ADSK_Номинальная мощность < 1000 Вт, 0.65, if(ADSK_Номинальная мощность < 4000 Вт, 0.75, 0.85))))'

    if name == 'ADSK_Полная мощность':
        return "ADSK_Номинальная мощность / ADSK_Коэффициент мощности"
    return None

def isDefaultExist(name):
    if name == 'ФОП_ВИС_Нагреватель или шкаф':
        return 0
    if name == 'ФОП_ВИС_Частотный регулятор':
        return 0
    return None

def isConnectorParaExist(name):
    if name == 'ADSK_Коэффициент мощности':
        return BuiltInParameter.RBS_ELEC_POWER_FACTOR  # коэффициент мощности

    if name == 'ADSK_Напряжение':
        return BuiltInParameter.RBS_ELEC_VOLTAGE  # напряжение

    if name == 'ADSK_Полная мощность':
        return BuiltInParameter.RBS_ELEC_APPARENT_LOAD  # полная мощность

    if name == 'ADSK_Количество фаз':
        return BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES  # количество полюсов


    return None


def add_para_group(paraNames, group, fopName):
    manager = doc.FamilyManager

    spFile = update_fop(fopName)

    for dG in spFile.Groups:
        for name in paraNames:
            if str(dG.Name) == group:
                myDefinitions = dG.Definitions
                eDef = myDefinitions.get_Item(name)


                newPara = familyParam()

                newPara.Name = name
                newPara.Param = manager.AddParameter(eDef, BuiltInParameterGroup.PG_ELECTRICAL_LOADS, True)
                newPara.Formula = isFormulaExist(name)
                newPara.defaultValue = isDefaultExist(name)
                newPara.connectorParam = isConnectorParaExist(name)

                newParams.append(newPara)



newParams = []
def main():
    try:
        manager = doc.FamilyManager
        global manager
    except Exception:
        print "Надстройка предназначена для работы с семействами"
        sys.exit()

    projectParams = doc.FamilyManager.Parameters

    #проверяем количество электрических коннекторов - должен быть строго один
    check_cons()

    #если в семействе нет никаких типоразмеров, ревит почему-то не даст создать формулы. Создаем хотя бы один тип
    typeNumber = 0
    for type in doc.FamilyManager.Types:
        typeNumber = typeNumber + 1
    if typeNumber < 1:
        doc.FamilyManager.NewType('Стандарт')

    #проверяем наличие параметров из списка имен в проекте, присваиваем им параметры типа или экземпляра
    for param in projectParams:
        if str(param.Definition.Name) in paraNamesADSK:
            paraNamesADSK.remove(param.Definition.Name)
        if str(param.Definition.Name) in paraNamesFOP:
            paraNamesFOP.remove(param.Definition.Name)
        if str(param.Definition.Name) in paraNamesMS:
            paraNamesMS.remove(param.Definition.Name)

    #добавляем сами параметры
    with revit.Transaction("Добавление параметров"):
        add_para_group(paraNamesADSK, "04 Обязательные ИНЖЕНЕРИЯ", "ADSK")

        add_para_group(paraNamesFOP, "03_ВИС", "V1")

        add_para_group(paraNamesMS, "mySchema", "V1")

    #настраиваем
    with revit.Transaction("Присвоение к коннектору"):
        for connector in connectorCol:
            if str(connector.Domain) == "DomainElectrical":
                connector.SystemClassification = MEPSystemClassification.PowerBalanced
                connector.SetParamValue(BuiltInParameter.RBS_ELEC_POWER_FACTOR_STATE, 1)

        #присваеваем параметры коннектору. Не очень хорошо, что я системные параметры проверяю по имени, надо посмотреть какой-то другой вариант
        for connector in connectorCol:
            params = connector.GetOrderedParameters()
            for param in params: #type: Parameter

                conParaBuiltIn = param.Definition.BuiltInParameter

                for newParam in newParams: #type: familyParam
                    if newParam.connectorParam == conParaBuiltIn:

                        # print conParaName
                        # print param.Definition.ParameterType
                        associate(param, newParam.Param)



    #почему-то не получается одно действие добавить формулы и значения
    # with revit.Transaction("Назначение формул"):
    #     for newParam in newParams: #type: familyParam
    #         if newParam.defaultValue:
    #             print 1
    #             manager.SetFormula(newParam.Param, '1 = 0')
    #             manager.Set(newParam.Param, 0)

        # for newParam in newParams: #type: familyParam
        #     if newParam.Formula:
        #         print newParam.Formula
        #         manager.SetFormula(newParam.Param, newParam.Formula)








main()
