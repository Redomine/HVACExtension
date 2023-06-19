#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Классификация'
__doc__ = "Обновляет код классификатор элементов"

import os.path as op
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from Microsoft.Office.Interop import Excel
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
import sys
import os
import paraSpec
from Autodesk.Revit.DB import *
from Redomine import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output

doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView

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
colPlumbingFixtures= make_col(BuiltInCategory.OST_PlumbingFixtures)
colSprinklers = make_col(BuiltInCategory.OST_Sprinklers)

collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves,  colInsulations, colPipeInsulations, colPipeCurves, colSprinklers, colAccessory,
               colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]




report_rows = []
def isElementEditedBy(element):
    try:
        edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
    except Exception:
        edited_by = __revit__.Application.Username

    if edited_by and edited_by != __revit__.Application.Username:
        if edited_by not in report_rows:
            report_rows.add(edited_by)
        return True
    return False






class classString:
    def __init__(self, name, mrCode, waterCode, drainCode, stromdrainCode, heatCode, ventCode):
        self.name = str(name)
        self.mrCode = str(mrCode)
        self.waterCode = str(waterCode)
        self.drainCode = str(drainCode)
        self.stromdrainCode = str(stromdrainCode)
        self.heatCode = str(heatCode)
        self.ventCode = str(ventCode)


def getClass(element):
    elemTypeId = element.GetTypeId()
    typeId = doc.GetElement(elemTypeId)
    code = typeId.GetParamValue(BuiltInParameter.UNIFORMAT_CODE)
    return code

def getSystemName(element):
    system = element.LookupParameter('Имя системы').AsString()
    return system

def checkSystem(element, object, system, targetIdent, VK, OT, VN):
    stromdrainSystems = ['К2', 'К2.1', 'K2', 'K2.1']
    drainSystems = ['К1', 'К3', 'К4']

    if VK:
        notChecked = True

        if notChecked:
            for sysName in stromdrainSystems:
                if sysName in system:
                    targetIdent.Set(object.stromdrainCode)
                    notChecked = False

        if notChecked:
            for sysName in drainSystems:
                if sysName in system:

                    targetIdent.Set(object.drainCode)
                    notChecked = False

        if notChecked:
            targetIdent.Set(object.waterCode)

    if OT:
        targetIdent.Set(object.heatCode)

    if VN:
        targetIdent.Set(object.ventCode)

def checkDepend(element):
    if hasattr(element, "GetSubComponentIds"):
        workCode = element.LookupParameter('ФОП_Код работы').AsString()
        sub_elements = [doc.GetElement(element_id) for element_id in element.GetSubComponentIds()]
        for sub_element in sub_elements:
            targetParam = sub_element.LookupParameter('ФОП_Код работы')
            targetParam.Set(workCode)

def getTableData():
    numberCol = 0
    nameCol = 1
    mrCol = 2
    heatCol = 6
    ventCol = 7
    waterCol = 3
    drainCol = 4
    stromdrainCol = 5

    dataList = []

    exel = Excel.ApplicationClass()


    path = str(os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\Классификаторы ОВиВК.xlsx"

    try:
        workbook = exel.Workbooks.Open(path, ReadOnly=True)
    except Exception:
        print 'Ошибка открытия таблицы, проверьте ее целостность'
        exel.ActiveWorkbook.Close(True)
        sys.exit()

    try:
        worksheet = workbook.Sheets["Сопоставление"]
    except Exception:
        print 'Не найден лист с названием Импорт, проверьте файл формы.'
        exel.ActiveWorkbook.Close(True)
        sys.exit()

    xlrange = worksheet.Range["A1", "AZ500"]



    row = 3
    while True:
        if xlrange.value2[row, numberCol] == None:
            break

        nameClass = xlrange.value2[row, nameCol]
        mrClass = xlrange.value2[row, mrCol]
        heatClass = xlrange.value2[row, heatCol]
        ventClass = xlrange.value2[row, ventCol]
        waterClass = xlrange.value2[row, waterCol]
        drainClass = xlrange.value2[row, drainCol]
        stromdrainClass = xlrange.value2[row, stromdrainCol]

        newString = classString(name=nameClass, mrCode=mrClass, heatCode=heatClass, ventCode=ventClass,
                                waterCode=waterClass, drainCode=drainClass, stromdrainCode=stromdrainClass)
        dataList.append(newString)
        row+=1

    exel.ActiveWorkbook.Close(True)

    return dataList

def script_execute():
    VK = False
    VN = False
    OT = False

    dataList = getTableData()

    if "OV_VN" in doc.Title:
        VN = True

    elif "OV_OT" in doc.Title:
        OT = True
    else:
        VK = True

    #report_rows = set()
    for collection in collections:
        for element in collection:
            dataInList = False
            targetIdent = lookupCheck(element, 'ФОП_Код работы', isExit = True)
            classification = getClass(element)
            #print classification

            system = getSystemName(element)

            for object in dataList:
                if object.mrCode == classification:
                    dataInList = True
                    checkSystem(element, object, system, targetIdent, VK, OT, VN)
            if not dataInList:
                targetIdent.Set("Код по классификатору не опознан")

        for element in collection:
            checkDepend(element)




if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

parametersAdded = paraSpec.check_parameters()


if not parametersAdded:
    information = doc.ProjectInformation
    with revit.Transaction("Обновление классификатора"):
        #список элементов для перебора в вид узлов:
        vis_collectors = []
        script_execute()
        # for report in report_rows:
        #     print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report

    if lookupCheck(information, 'ФОП_ВИС_Нумерация позиций').AsInteger() == 1 or lookupCheck(information, 'ФОП_ВИС_Площади воздуховодов в примечания').AsInteger() == 1:
        import numerateSpec