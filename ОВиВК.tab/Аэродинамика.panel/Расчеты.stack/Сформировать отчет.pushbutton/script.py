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
    print 'Для формирования отчета выделите систему перед запуском плагина'
    sys.exit()
if selectedIds.Count > 1:
    print 'Нужно выделить только одну систему'
    sys.exit()
system = doc.GetElement(selectedIds[0])
if selectedIds.Count == 1 and system.Category.IsId(BuiltInCategory.OST_DuctSystem) == False:
    print 'Обработке подлежат только системы воздуховодов'
    sys.exit()

if len(system.GetCriticalPathSectionNumbers()) == 0 or system.PressureLossOfCriticalPath == 0:
    print 'У выделенной системы не ведется расчет статического давления или оно зануляется'
    sys.exit()

view = doc.ActiveView



def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    


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


    old_flow = 0
    for con in conSet:
        connectorSet = con.AllRefs.ForwardIterator()
        while connectorSet.MoveNext():
            try:
                flow = connectorSet.Current.Owner.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
            except Exception:
                flow = 0
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
            Fc = ductCon.Height * 0.3048 * ductCon.Width * 0.3048
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
                K = ((1-fp**0.5)+0.5*l0+0.05)*(1.7+(1/(2*f0)-1)*l0-((fp+f0)*l0)**0.5)*(fp/(1-l0))**2
            else:
                K = (-0.7-6.05*(1-fp)**3)*(f0/l0)**2+(1.32+3.23*(1-fp)**2)*f0/l0+(0.5+0.42*fp)-0.167*l0/f0
        else:
            if Lc > Lo:
                K = (fp/(1-l0))**2*((1-fp)+0.5*l0+0.05)*(1.5+(1/(2*f0)-1)*l0-((fp+f0)*l0)**0.5)
            else:
                K = (f0/l0)**2*(4.1*(fp/f0)**1.25*l0**1.5*(fp+f0)**(0.3*(f0/fp)**0.5/l0-2)-0.5*fp/f0)



    if str(conSet[0].DuctSystemType) == "SupplyAir":
        if form == "Круглый отвод":
            if Lc > Lo:
                K = 0.45*(fp/(1-l0))**2+(0.6-1.7*fp)*fp/(1-l0)-(0.25-0.9*fp**2)+0.19*(1-l0)/fp
            else:
                K = (f0/l0)**2-0.58*f0/l0+0.54+0.025*l0/f0
        else:
            if Lc > Lo:
                K = 0.45*(fp/(1-l0))**2+(0.6-1.7*fp)*fp/(1-l0)-(0.25-0.9*fp**2)+0.19*(1-l0)/fp
            else:
                K = (f0/l0)**2-0.42*f0/l0+0.81-0.06*l0/f0

    return K

path_numbers = system.GetCriticalPathSectionNumbers()
path = []
for number in path_numbers:
    path.append(number)
if str(system.SystemType) == "SupplyAir":
    path.reverse()

passed_taps = []
output = script.get_output()
old_flow = 0
for number in path:

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
            lenght = section.GetSegmentLength(elementId) * 304.8 / 1000
            lenght = float('{:.2f}'.format(lenght))
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
            flow = section.Flow * 101.941317259
            flow = int(flow)
        except Exception:
            pass
        if old_flow < flow:
            old_flow = flow
            count += 1
        velocity = '-'
        try:
            velocity = section.Velocity * 0.30473037475
            velocity = float('{:.2f}'.format(velocity))
        except Exception:
            pass
        if velocity == 0:
            velocity = '-'

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

                    K = float('{:.2f}'.format(K))
                    Z = Pd * K
                    coef = K
                    pressure_drop = Z
                    passed_taps.append(element.Id)

        pressure_drop = float('{:.2f}'.format(pressure_drop))

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