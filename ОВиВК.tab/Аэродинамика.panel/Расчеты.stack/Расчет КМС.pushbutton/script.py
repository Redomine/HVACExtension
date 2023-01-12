#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пересчет КМС'
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"


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
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig


doc = __revit__.ActiveUIDocument.Document # type: Document

uidoc = __revit__.ActiveUIDocument




view = doc.ActiveView



def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 



colFittings = make_col(BuiltInCategory.OST_DuctFitting)
colSystem = make_col(BuiltInCategory.OST_DuctSystem)



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

def getDpElbow(element):
    a = getConnectors(element)
    angle = a[1].Angle

    try:
        sizes = [a[0].Height* 304.8, a[0].Width* 304.8]
        H = max(sizes)
        B = min(sizes)
        E = 2.71828182845904

        K = (0.25 * (B/H) ** 0.25) * (1.07 * E**(2/(2*(100 + B/2)/B +1)) - 1) ** 2

        if angle <= 1:
            K = K * 0.708

    except:
        if angle > 1:
            K = 0.33
        else:
            K = 0.18

    return K

def getDpTransition(element):
    a = getConnectors(element)
    try:
        S1 = a[0].Height*304.8*a[0].Width*304.8
    except:
        S1 = 3.14*304.8*304.8*a[0].Radius**2
    try:
        S2 = a[1].Height*304.8*a[1].Width*304.8
    except:
        S2 = 3.14*304.8*304.8*a[1].Radius**2


    #проверяем в какую сторону дует воздух чтоб выяснить расширение это или заужение
    if str(a[0].Direction) == "In":
        if S1 > S2:
            transition = 'Заужение'
            F0 = S2
            F1 = S1
        else:
            transition = 'Расширение'
            F0 = S1
            F1 = S2
    if str(a[0].Direction) == "Out":
        if S1 < S2:
            transition = 'Заужение'
            F0 = S1
            F1 = S2
        else:
            transition = 'Расширение'
            F0 = S2
            F1 = S1

    F = F0/F1



    if transition == 'Расширение':
        if F < 0.11: K = 0.81
        elif F < 0.21: K = 0.64
        elif F < 0.31: K = 0.5
        elif F < 0.41: K = 0.36
        elif F < 0.51: K = 0.26
        elif F < 0.61: K = 0.16
        elif F < 0.71: K = 0.09
        else: K = 0.04
    if transition == 'Заужение':
        if F < 0.11: K = 0.45
        elif F < 0.21: K = 0.4
        elif F < 0.31: K = 0.35
        elif F < 0.41: K = 0.3
        elif F < 0.51: K = 0.25
        elif F < 0.61: K = 0.2
        elif F < 0.71: K = 0.15
        else: K = 0.1

    return K

def getConCoords(connector):
    a0 = connector.Origin.ToString()
    a0 = a0.replace("(", "")
    a0 = a0.replace(")", "")
    a0 = a0.split(",")
    for x in a0:
        x = float(x)
    return a0

def getDuctCoords(inTeeCon, connector):
    mainCon = []
    connectorSet = connector.AllRefs.ForwardIterator()
    while connectorSet.MoveNext():
        mainCon.append(connectorSet.Current)
    duct = mainCon[0].Owner
    ductCons = getConnectors(duct)
    for ductCon in ductCons:
        if getConCoords(ductCon) != inTeeCon:
            inDuctCon = getConCoords(ductCon)
            return inDuctCon



def getTeeOrient(element):
    connectors = getConnectors(element)
    exitCons = []
    exhaustAirCons = []
    for connector in connectors:
        if str(connectors[0].DuctSystemType) == "SupplyAir":
            if connector.Flow != max(connectors[0].Flow, connectors[1].Flow, connectors[2].Flow):
                exitCons.append(connector)
        if str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir":
            #а что делать если на на разветвлении расход одинаковы?
            if connector.Flow == max(connectors[0].Flow, connectors[1].Flow, connectors[2].Flow):
                exitCons.append(connector)
            else:
                exhaustAirCons.append(connector)
            #для входа в тройник ищем координаты начала входящего воздуховода чтоб построить прямую через эти две точки

        if str(connectors[0].DuctSystemType) == "SupplyAir":
            if str(connector.Direction) == "In":
                inTeeCon = getConCoords(connector)
                # выбираем из коннектора подключенный воздуховод
                inDuctCon = getDuctCoords(inTeeCon, connector)

    #в случе вытяжной системы, чтоб выбрать коннектор с выходящим воздухом из второстепенных, берем два коннектора у которых расход не максимальны
    #(максимальный точно выходной у вытяжной системы) и сравниваем. Тот что самый малый - ответветвление
    #а второй - точка вхождения потока воздуха из которой берем координаты для построения вектора



    if str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir":

        if exhaustAirCons[0].Flow < exhaustAirCons[1].Flow:
            exitCons.append(exhaustAirCons[0])
            inTeeCon = getConCoords(exhaustAirCons[1])
            inDuctCon = getDuctCoords(inTeeCon, exhaustAirCons[1])
        else:
            exitCons.append(exhaustAirCons[1])
            inTeeCon = getConCoords(exhaustAirCons[0])
            inDuctCon = getDuctCoords(inTeeCon, exhaustAirCons[0])


    # среди выходящих коннекторов ищем диктующий по большему расходу
    if exitCons[0].Flow > exitCons[1].Flow:
        exitCon = exitCons[0]
        secondaryCon = exitCons[1]
    else:
        exitCon = exitCons[1]
        secondaryCon = exitCons[0]


    #диктующий коннектор
    exitCon = getConCoords(exitCon)

    #вторичный коннектор
    secondaryCon = getConCoords(secondaryCon)

    #найдем вектор по координатам точек AB = {Bx - Ax; By - Ay; Bz - Az}
    ductToTee = [(float(inDuctCon[0]) - float(inTeeCon[0])), (float(inDuctCon[1]) - float(inTeeCon[1])),
                 (float(inDuctCon[2]) - float(inTeeCon[2]))]

    teeToExit = [(float(inTeeCon[0]) - float(exitCon[0])), (float(inTeeCon[1]) - float(exitCon[1])),
               (float(inTeeCon[2]) - float(exitCon[2]))]

    #то же самое для вторичного отвода
    teeToMinor = [(float(inTeeCon[0]) - float(secondaryCon[0])), (float(inTeeCon[1]) - float(secondaryCon[1])),
               (float(inTeeCon[2]) - float(secondaryCon[2]))]

    #найдем скалярное произведение векторов AB · CD = ABx · CDx + ABy · CDy + ABz · CDz
    teeToExit_ductToTee = ductToTee[0]*teeToExit[0] + ductToTee[1]*teeToExit[1] + ductToTee[2]*teeToExit[2]

    #то же самое с вторичным коннектором
    teeToMinor_ductToTee = ductToTee[0]*teeToMinor[0] + ductToTee[1]*teeToMinor[1] + ductToTee[2]*teeToMinor[2]

    #найдем длины векторов
    len_ductToTee = ((ductToTee[0])**2 + (ductToTee[1])**2 + (ductToTee[2])**2)**0.5
    len_teeToExit = ((teeToExit[0])**2 + (teeToExit[1])**2 + (teeToExit[2])**2)**0.5

    #то же самое для вторичного вектора
    len_teeToMinor = ((teeToMinor[0])**2 + (teeToMinor[1])**2 + (teeToMinor[2])**2)**0.5


    #найдем косинус
    cosMain = (teeToExit_ductToTee)/(len_ductToTee * len_teeToExit)

    #то же самое с вторичным вектором
    cosMinor = (teeToMinor_ductToTee) / (len_ductToTee * len_teeToMinor)

    #Если угол расхождения между вектором входа воздуха и выхода больше 10 градусов(цифра с потолка) то считаем что идет буквой L
    #Если нет, то считаем что идет по прямой буквой I

    #тип 1
    #вытяжной воздуховод zп
    if math.acos(cosMain) < 0.10 and (str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir"):
        type = 1

    #тип 2
    #вытяжной воздуховод, zо
    elif math.acos(cosMain) > 0.10 and (str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir"):
        type = 2

    #тип 3
    #подающий воздуховод, zп
    elif math.acos(cosMain) < 0.10 and str(connectors[0].DuctSystemType) == "SupplyAir":
        type = 3

    #тип 4
    #подающий воздуховод, zо
    elif math.acos(cosMain) > 0.10 and str(connectors[0].DuctSystemType) == "SupplyAir":
        type = 4

    return type

def getDpTee(element):
    conSet = getConnectors(element)

    try:
        type = getTeeOrient(element)
    except Exception:
        type = 'Ошибка'

    try:
        S1 = conSet[0].Height*0.3048*conSet[0].Width*0.3048
    except:
        S1 = 3.14*0.3048*0.3048*conSet[0].Radius**2
    try:
        S2 = conSet[1].Height*0.3048*conSet[1].Width*0.3048
    except:
        S2 = 3.14*0.3048*0.3048*conSet[1].Radius**2
    try:
        S3 = conSet[2].Height*0.3048*conSet[2].Width*0.3048
    except:
        S3 = 3.14*0.3048*0.3048*conSet[2].Radius**2

    if type != 'Ошибка':
        v1 = conSet[0].Flow*101.94
        v2 = conSet[1].Flow*101.94
        v3 = conSet[2].Flow*101.94
        Vset = [v1, v2, v3]
        Lc = max(Vset)
        Vset.remove(Lc)
        if Vset[0] > Vset[1]:
            Lo = Vset[1]
        else:
            Lo = Vset[0]

        Fc = max([S1, S2, S3])
        Fo = min([S1, S2, S3])
        Fp = Fo

        f0 = Fo/Fc
        l0 = Lo/Lc
        fp = Fp/Fc

        if f0 < 0.351:  koef_1 = 0.8 * l0
        elif l0 < 0.61: koef_1 = 0.5
        else: koef_1 = 0.8*l0

        if f0 < 0.351:  koef_2 = 1
        elif l0 < 0.41: koef_2 = 0.9*(1-l0)
        else: koef_2 = 0.55

        if f0 < 0.41:  koef_3 = 0.4
        elif l0 < 0.51: koef_3 = 0.5
        else: koef_3 = 2*l0-1

        if f0 < 0.36:
            if l0 < 0.41: koef_4 = 1.1-0.7*l0
            else: koef_4 = 0.85
        else:
            if l0 < 0.61: koef_4 = 1 - 0.6*l0
            else: koef_4 = 0.6

    if type == 1:
        K = (1-(1-l0)**2-(1.4-l0)*l0**2)/((1-l0)**2/fp**2)

    if type == 2:
        K = koef_2*(1+(l0/f0)**2-2*(1-l0)**2)/(l0/f0)**2

    if type == 3:
        K = koef_3*l0**2/((1-l0)**2/Fp**2)

    if type == 4:
        K = (koef_4*(1+(l0/f0)**2))/(l0/f0)**2
    
    if type == 'Ошибка':
        K = 0

    return K

def getDpTapAdjustable(element):
    conSet = getConnectors(element)

    try:
        Fo = conSet[0].Height*0.3048*conSet[0].Width*0.3048
        form = "Прямоугольный отвод"
    except:
        Fo = 3.14*0.3048*0.3048*conSet[0].Radius**2
        form = "Круглый отвод"

    mainCon = []

    connectorSet_0 = conSet[0].AllRefs.ForwardIterator()

    connectorSet_1 = conSet[1].AllRefs.ForwardIterator()

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
            if Lc > Lo * 2:
                K = ((1-fp**0.5)+0.5*l0+0.05)*(1.7+(1/(2*f0)-1)*l0-((fp+f0)*l0)**0.5)*(fp/(1-l0))**2
            else:
                K = (-0.7-6.05*(1-fp)**3)*(f0/l0)**2+(1.32+3.23*(1-fp)**2)*f0/l0+(0.5+0.42*fp)-0.167*l0/f0
        else:
            if Lc > Lo * 2:
                K = (fp/(1-l0))**2*((1-fp)+0.5*l0+0.05)*(1.5+(1/(2*f0)-1)*l0-((fp+f0)*l0)**0.5)
            else:
                K = (f0/l0)**2*(4.1*(fp/f0)**1.25*l0**1.5*(fp+f0)**(0.3*(f0/fp)**0.5/l0-2)-0.5*fp/f0)



    if str(conSet[0].DuctSystemType) == "SupplyAir":
        if form == "Круглый отвод":
            if Lc > Lo * 2:
                K = 0.45*(fp/(1-l0))**2+(0.6-1.7*fp)*fp/(1-l0)-(0.25-0.9*fp**2)+0.19*(1-l0)/fp
            else:
                K = (f0/l0)**2-0.58*f0/l0+0.54+0.025*l0/f0
        else:
            if Lc > Lo * 2:
                K = 0.45*(fp/(1-l0))**2+(0.6-1.7*fp)*fp/(1-l0)-(0.25-0.9*fp**2)+0.19*(1-l0)/fp
            else:
                K = (f0/l0)**2-0.42*f0/l0+0.81-0.06*l0/f0

    return K

def getLossMethods(serviceId):
    lc=[]
    service = ExternalServiceRegistry.GetService(serviceId)
    serverIds = service.GetRegisteredServerIds()
    list=List[ElementId]()
    for serverId in serverIds:
        server = getServerById(serverId, serviceId)
        id=serverId
        name=server.GetName()
        lc.append(id)
        lc.append(name)
        lc.append(server)
    return lc

def getServerById(serverGUID, serviceId):
    service = ExternalServiceRegistry.GetService(serviceId)
    if service != "null" and serverGUID != "null":
        server = service.GetServer(serverGUID)
        if server != "null":
            return server
    return None

def script_execute():
    report_rows = set()
    with revit.Transaction("Пересчет потерь напора"):
        for element in colFittings:
            try:
                edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
            except Exception:
                print element.Id
            if edited_by and edited_by != __revit__.Application.Username:
                report_rows.add(edited_by)
                continue

            K = 0

            try:
                if str(element.MEPModel.PartType) == 'Elbow':
                    K = getDpElbow(element)
            except Exception:
                pass

            try:
                if str(element.MEPModel.PartType) == 'Transition':
                    K = getDpTransition(element)
            except Exception:
                pass

            try:
                if str(element.MEPModel.PartType) == 'Tee':
                    K = getDpTee(element)
            except Exception:
                pass


            try:
                if str(element.MEPModel.PartType) == 'TapAdjustable':
                    K = getDpTapAdjustable(element)
            except Exception:
                pass

            #print element.Id.IntegerValue


            #выбираем метод потерь по гуиду, "определенный коэффициент"
            # schema = Schema.Lookup(Guid("13ded697-d107-4b0d-8dc4-2a2e4c870096"))
            # fitting = doc.GetElement(element.Id)
            # param = fitting.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
            # param.Set("Coefficient")
            # entity = element.GetEntity(schema)
            # field = schema.GetField("Coefficient")
            # entity.Set[field.ValueType](field, str(K))
            # element.SetEntity(entity)

            eleId = element.Id
            fitting = doc.GetElement(eleId)
            param = fitting.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
            lc = getLossMethods(ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService)
            param.Set(lc[6].ToString()) # установка метода потерь
            schema = lc[8].GetDataSchema()

            field = schema.GetField("Coefficient")
            entity=fitting.GetEntity(schema)

            try:
                entity.Set[field.ValueType](field, str(K))
                fitting.SetEntity(entity)
            except Exception:
                pass

    if report_rows:
        print "Некоторые элементы не были обработаны, так как были заняты пользователями:"
        print "\r\n".join(report_rows)

    with revit.Transaction("Выключение систем"):
        colSystems = make_col(BuiltInCategory.OST_DuctSystem)
        for el in colSystems:
            sysType = doc.GetElement(el.GetTypeId())
            sysType.CalculationLevel = sysType.CalculationLevel.None

    with revit.Transaction("Включение систем"):
        colSystems = make_col(BuiltInCategory.OST_DuctSystem)
        for el in colSystems:
            sysType = doc.GetElement(el.GetTypeId())
            sysType.CalculationLevel = sysType.CalculationLevel.All

script_execute()

script_execute()