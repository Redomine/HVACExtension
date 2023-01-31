#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление ФОП_Этаж'
__doc__ = "Перенос значения уровня в ФОП_Этаж"

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

def getConnectors(element):
    connectors = []
    try:
        a = element.ConnectorManager.Connectors.ForwardIterator()
        while a.MoveNext():
            connectors.append(a.Current)
    except:
        try:
            a = element.MEPModel.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
        except:
            a = element.MEPSystem.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
    return connectors


def script_execute():
    with revit.Transaction("Перенос уровня"):
        report_rows = set()
        for collection in collections:
            for element in collection:
                try:
                    edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
                except Exception:
                    print element.Id

                if edited_by and edited_by != __revit__.Application.Username:
                    report_rows.add(edited_by)
                    continue

                FOP_Level = element.LookupParameter('ФОП_Этаж')

                level = 'None'
                if element.Category.IsId(BuiltInCategory.OST_PipeInsulations) \
                        or element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
                    connectors = getConnectors(element)
                    for connector in connectors:
                        for el in connector.AllRefs:
                            if el.Owner.LookupParameter('Уровень'):
                                level = el.Owner.LookupParameter('Уровень').AsValueString()
                            if el.Owner.LookupParameter('Базовый уровень'):
                                level = el.Owner.LookupParameter('Базовый уровень').AsValueString()
                else:
                    if element.LookupParameter('Уровень'):
                        level = element.LookupParameter('Уровень').AsValueString()
                    if element.LookupParameter('Базовый уровень'):
                        level = element.LookupParameter('Базовый уровень').AsValueString()



                try:

                    FOP_Level.Set(level)
                except:
                    print element.Id
                    print level
                    print element.LookupParameter('Базовый уровень').AsValueString()
                    sys.Exit()
        if report_rows:
            print "Некоторые элементы не были обработаны, так как были заняты пользователями:"
            print "\r\n".join(report_rows)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

parametersAdded = paraSpec.check_parameters()

if not parametersAdded:
        script_execute()
