#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Импорт немоделируемых'
__doc__ = "Генерирует в модели элементы в соответствии с их ведомостью"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import System
import paraSpec
import checkAnchor

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit


from Microsoft.Office.Interop import Excel
from Redomine import *
from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert

#Исходные данные
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)

nameOfModel = '_Якорный элемент'
description = 'Импорт немоделируемых'

def setElement(element, name, setting):
    if name == 'ФОП_ВИС_Число':
        try:
            element.LookupParameter(name).Set(setting)
        except:
            element.LookupParameter(name).Set(0)
    if name == 'ФОП_ВИС_Масса':
        pass
    else:
        if setting == None or setting == 'None':
            setting = ''
        element.LookupParameter(name).Set(str(setting))

temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, nameOfModel)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

if temporary == None:
    print 'Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()

def script_execute():

    exel = Excel.ApplicationClass()
    filepath = select_file()


    ADSK_System_Names = []
    System_Named = True

    try:
        workbook = exel.Workbooks.Open(filepath)
    except Exception:
        print 'Ошибка открытия таблицы, проверьте ее целостность'
        sys.exit()
    sheet_name = 'Импорт'

    try:
        worksheet = workbook.Sheets[sheet_name]
    except Exception:
        print 'Не найден лист с названием Импорт, проверьте файл формы.'
        sys.exit()

    xlrange = worksheet.Range["A1", "AZ500"]

    FOP_Corp = 0
    FOP_Sec = 1
    FOP_Floor = 2
    FOP_System = 3
    FOP_Group = 4
    FOP_Name = 5
    FOP_Mark = 6
    FOP_Art = 7
    FOP_Maker = 8
    FOP_Izm = 9
    FOP_Number = 10
    FOP_Mass = 11
    FOP_Comment = 12
    FOP_EF = 13

    report_rows = set()

    for element in colModel:
        if isElementEditedBy(element):
            fillReportRows(element, report_rows)
            for report in report_rows:
                print 'Якорные элементы не были обработаны, так как были заняты пользователями:' + report
            sys.exit()

    with revit.Transaction("Добавление расчетных элементов"):
        # при каждом повторе расчета удаляем старые версии
        remove_models(colModel, nameOfModel, description)
        #при каждом повторе расчета удаляем старые версии

        calculation_elements = []
        row = 2
        while True:
            if xlrange.value2[row, FOP_Name] == None:
                break
            newPos = rowOfSpecification(corp = xlrange.value2[row, FOP_Corp],
                            sec = xlrange.value2[row, FOP_Sec],
                            floor = xlrange.value2[row, FOP_Floor],
                            system = xlrange.value2[row, FOP_System],
                            group = xlrange.value2[row, FOP_Group],
                            name = xlrange.value2[row, FOP_Name],
                            mark = xlrange.value2[row, FOP_Mark],
                            art = xlrange.value2[row, FOP_Art],
                            maker = xlrange.value2[row, FOP_Maker],
                            unit = xlrange.value2[row, FOP_Izm],
                            number = xlrange.value2[row, FOP_Number],
                            mass = xlrange.value2[row, FOP_Mass],
                            comment = xlrange.value2[row, FOP_Comment],
                            EF = xlrange.value2[row, FOP_EF])


            row += 1

            calculation_elements.append(newPos)

        # в следующем блоке генерируем новые экземпляры пустых семейств куда уйдут расчеты
        new_position(calculation_elements, temporary, nameOfModel, description)

    exel.ActiveWorkbook.Close(True)
    Marshal.ReleaseComObject(worksheet)
    Marshal.ReleaseComObject(workbook)
    Marshal.ReleaseComObject(exel)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

status = paraSpec.check_parameters()

if not status:
    anchor = checkAnchor.check_anchor(showText = False)
    if anchor:
        script_execute()