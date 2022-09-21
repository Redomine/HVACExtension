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
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)


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

def getDpTapAdjustable(element):
    conSet = getConnectors(element)

    try:
        Fo = conSet[0].Height*0.3048*conSet[0].Width*0.3048
        form = "Прямоугольный отвод"
    except:
        Fo = 3.14*0.3048*0.3048*conSet[0].Radius**2
        form = "Круглый отвод"

    mainCon = []
    #if element.Id.IntegerValue == 2752996:
    #    for con in conSet:
    #        print con.AllRefs.ForwardIterator().Current
            #.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)


    connectorSet_0 = conSet[0].AllRefs.ForwardIterator()

    connectorSet_1 = conSet[1].AllRefs.ForwardIterator()

    old_flow = 0
    for con in conSet:
        connectorSet = con.AllRefs.ForwardIterator()
        while connectorSet.MoveNext():
            flow = connectorSet.Current.Owner.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
            if flow > old_flow:
                mainCon = []
                mainCon.append(connectorSet.Current)
                old_flow = flow



    duct = mainCon[0].Owner


    ductCons = getConnectors(duct)
    Flow = []

    for ductCon in ductCons:
        Flow.append(ductCon.Flow*101.94)
        try:
            Fc = conSet[0].Height * 0.3048 * ductCon.Width * 0.3048
            Fp = Fc
        except:
            Fc = 3.14 * 0.3048 * 0.3048 * ductCon.Radius ** 2
            Fp = Fc

    Lc = max(Flow)
    Lo = conSet[0].Flow*101.94

    f0 = Fo / Fc
    l0 = Lo / Lc
    fp = Fp / Fc

    if str(conSet[0].DuctSystemType) == "ExhaustAir" or str(conSet[0].DuctSystemType) == "ReturnAir":
        if form == "Круглый отвод":
            if Lc > Lo:
                K = ((1-Fp**0.5)+0.5*l0+0.05)*(1.7+(1/(2*f0)-1)*l0-((Fp+f0)*l0)**0.5)*(Fp/(1-l0))**2
            else:
                K = (-0.7-6.05*(1-Fp)**3)*(f0/l0)**2+(1.32+3.23*(1-Fp)**2)*f0/l0+(0.5+0.42*Fp)-0.167*l0/f0
        else:
            if Lc > Lo:
                K = (Fp/(1-l0))**2*((1-Fp)+0.5*l0+0.05)*(1.5+(1/(2*f0)-1)*l0-((Fp+f0)*l0)**0.5)
            else:
                K = (f0/l0)**2*(4.1*(Fp/f0)**1.25*l0**1.5*(Fp+f0)**(0.3*(f0/Fp)**0.5/l0-2)-0.5*Fp/f0)


    if str(conSet[0].DuctSystemType) == "SupplyAir":
        if form == "Круглый отвод":
            if Lc > Lo:
                K = 0.45*(Fp/(1-l0))**2+(0.6-1.7*Fp)*Fp/(1-l0)-(0.25-0.9*Fp**2)+0.19*(1-l0)/Fp
            else:
                K = (f0/l0)**2-0.58*f0/l0+0.54+0.025*l0/f0
        else:
            if Lc > Lo:
                K = 0.45*(Fp/(1-l0))**2+(0.6-1.7*Fp)*Fp/(1-l0)-(0.25-0.9*Fp**2)+0.19*(1-l0)/Fp
            else:
                K = (f0/l0)**2-0.42*f0/l0+0.81-0.06*l0/f0

    return K


passed_taps = []
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

        velocity = 0
        try:
            velocity = section.Velocity
        except Exception:
            pass

        pressure_drop = 0
        try:
            pressure_drop = section.GetPressureDrop(elementId) * 3.280839895

        except Exception:
            pass

        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            if str(element.MEPModel.PartType) == 'TapAdjustable':
                if element.Id not in passed_taps:
                    Pd = (1.21 * velocity * velocity)/2 #Динамическое давление
                    K = getDpTapAdjustable(element) #КМС
                    Z = Pd * K
                    pressure_drop = Z
                    passed_taps.append(element.Id)

        summ_pressure += pressure_drop
        if pressure_drop == 0:
            continue
        else:
            data.append([count, name, lenght, size, flow, velocity, coef, pressure_drop, summ_pressure, output.linkify(elementId)])



output.print_table(table_data=data,
                   title=("Отчет о расчете аэродинамики системы " + system_name),
                   columns=["Номер участка", "Наименование элемента", "Длина, м.п.","Размер", "Расход, м3/ч", "Скорость, м/с", "КМС", "Потери напора элемента, Па", "Суммарные потери напора, Па", "Id элемента"],
                   formats=['', '', ''],
                   )