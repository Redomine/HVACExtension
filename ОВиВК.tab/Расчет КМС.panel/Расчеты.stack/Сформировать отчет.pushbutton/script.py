#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Формирование отчета'
__doc__ = "Формирует отчет о расчете аэродинамики"


import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selectedIds = uidoc.Selection.GetElementIds()
if 0 == selectedIds.Count:
    print 'Выделите систему для формирования отчета перед запуском плагина'
if selectedIds.Count > 1:
    print 'Нужно выделить только одну систему'

system = doc.GetElement(selectedIds[0])
if selectedIds.Count == 1 and system.Category.IsId(BuiltInCategory.OST_DuctSystem) == False:
    print 'Обработке подлежат только системы воздуховодов'

view = doc.ActiveView



def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
path_numbers = system.GetCriticalPathSectionNumbers()

data = []
count = 0
summ_pressure = 0
system_name = system.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)

def optimise_list(data_list):
    old_count = ''
    old_name = ''
    old_lenght = ''
    old_size = ''
    old_pressure_drop = ''
    old_summ_pressure = ''
    old_elementId = ''
    #for data in data_list:
    #    count, name, lenght, size, pressure_drop, summ_pressure, elementId

output = script.get_output()
for number in path_numbers:
    count += 1
    section = system.GetSectionByNumber(number)
    elementsIds = section.GetElementIds()
    for elementId in elementsIds:
        element = doc.GetElement(elementId)
        name = ''
        if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
            name = 'Воздуховод'
        elif element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
            name = 'Воздухораспределитель'
        elif element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
            name = 'Оборудование'
        elif element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            name = 'Фасонный элемент воздуховода'
            if str(element.MEPModel.PartType) == 'Elbow':
                name = 'Отвод воздуховода'
            if str(element.MEPModel.PartType) == 'Transition':
                name = 'Переход между сечениями'
            if str(element.MEPModel.PartType) == 'Tee':
                name = 'Тройник'
            if str(element.MEPModel.PartType) == 'TapAdjustable':
                name = 'Врезка'

        else:
            name = 'Арматура'

        size = '-'
        try:
            size = element.GetParamValue(BuiltInParameter.RBS_CALCULATED_SIZE)
        except Exception:
            pass

        lenght = '-'
        try:
            lenght = section.GetSegmentLength(elementId)
        except Exception:
            pass

        coef = '-'
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            try:
                coef = section.GetCoefficient(elementId)
            except Exception:
                pass

        flow = 0
        try:
            flow = section.Flow
        except Exception:
            pass

        pressure_drop = 0
        try:
            pressure_drop = section.GetPressureDrop(elementId) * 3.280839895
            summ_pressure += pressure_drop
        except Exception:
            pass
        if pressure_drop == 0:
            continue
        else:
            data.append([count, name, lenght, size, flow, coef, pressure_drop, summ_pressure, output.linkify(elementId)])



output.print_table(table_data=data,
                   title=("Отчет о расчете аэродинамики системы " + system_name),
                   columns=["Номер участка", "Наименование элемента", "Длина, м.п.","Размер", "Расход, м3/ч", "КМС", "Потери напора элемента, Па", "Суммарные потери напора, Па", "Id элемента"],
                   formats=['', '', ''],
                   )