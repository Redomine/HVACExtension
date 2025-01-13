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
from pyrevit import forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, Selection
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from collections import namedtuple

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig

class Aerodinamiccoefficientcalculator:
    LOSS_GUID_CONST = "46245996-eebb-4536-ac17-9c1cd917d8cf" # Гуид для удельных потерь
    COEFF_GUID_CONST = "5a598293-1504-46cc-a9c0-de55c82848b9" # Это - Гуид "Определенный коэффициент". Вроде бы одинаков всегда

    doc = None
    uidoc = None
    view = None

    def __init__(self, doc, uidoc, view):
        self.doc = doc
        self.uidoc = uidoc
        self.view = view

    def is_supply_air(self, connector):
        return connector.DuctSystemType == DuctSystemType.SupplyAir

    def is_exhaust_air(self, connector):
        return (connector.DuctSystemType == DuctSystemType.ExhaustAir
                or connector.DuctSystemType == DuctSystemType.ReturnAir)

    def is_direction_inside(self, connector):
        return connector.Direction == FlowDirectionType.In

    def is_direction_bidirectonal(self, connector):
        return connector.Direction == FlowDirectionType.Bidirectional

    def is_direction_outside(self, connector):
        return connector.Direction == FlowDirectionType.Out

    def get_connectors(self, element):
        connectors = []

        if isinstance(element, FamilyInstance) and element.MEPModel.ConnectorManager is not None:
            connectors.extend(element.MEPModel.ConnectorManager.Connectors)

        if element.InAnyCategory([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves]) and \
                isinstance(element, MEPCurve) and element.ConnectorManager is not None:
            connectors.extend(element.ConnectorManager.Connectors)

        return connectors

    def get_con_coords(self, connector):
        a0 = connector.Origin.ToString()
        a0 = a0.replace("(", "")
        a0 = a0.replace(")", "")
        a0 = a0.split(",")
        for x in a0:
            x = float(x)
        return a0

    def get_coef_elbow(self, element):
        a = self.get_connectors(element)
        angle = a[1].Angle

        try:
            sizes = [a[0].Height * 304.8, a[0].Width * 304.8]
            H = max(sizes)
            B = min(sizes)
            E = 2.71828182845904

            K = (0.25 * (B / H) ** 0.25) * (1.07 * E ** (2 / (2 * (100 + B / 2) / B + 1)) - 1) ** 2

            if angle <= 1:
                K = K * 0.708

        except:
            if angle > 1:
                K = 0.33
            else:
                K = 0.18

        return K

    def get_connector_area(self, connector):
        if connector.Shape == ConnectorProfileType.Round:
            radius = self.convert_to_milimeters(connector.Radius)
            area = math.pi * radius ** 2
        else:
            height = self.convert_to_milimeters(connector.Height)
            width = self.convert_to_milimeters(connector.Width)
            area = height * width
        return area

    def convert_to_milimeters(self, value):
        return  UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.Millimeters)

    def convert_to_square_meters(self, value):
        return  UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.SquareMeters)

    def get_coef_transition(self, element):
        connectors = self.get_connectors(element)

        con1_area = self.get_connector_area(connectors[0])
        con2_area = self.get_connector_area(connectors[1])

        transition_extension = None # False - заужение, True - расширение
        # Определяем тип перехода (расширение или сужение)
        if self.is_direction_inside(connectors[0]):
            transition_extension = True if con1_area <= con2_area else False
            F0, F1 = (con1_area, con2_area) if con1_area <= con2_area else (con2_area, con1_area)
        elif self.is_direction_outside(connectors[0]):
            transition_extension = True if con1_area >= con2_area else False
            F0, F1 = (con2_area, con1_area) if con1_area >= con2_area else (con1_area, con2_area)
        else: # Если вызывается эта часть - поток двунаправленный. Принимаем для притока всегда заужение,
            # для вытяжки всегда расширение
            if self.is_supply_air(connectors[0]):
                transition_extension = False
                F0, F1 = (con2_area, con1_area) if con1_area > con2_area else (con1_area, con2_area)
            else:
                transition_extension = True
                F0, F1 = (con1_area, con2_area) if con1_area < con2_area else (con2_area, con1_area)

        F = F0 / F1

        if transition_extension:
            if F < 0.11:
                coefficient = 0.81
            elif F < 0.21:
                coefficient = 0.64
            elif F < 0.31:
                coefficient = 0.5
            elif F < 0.41:
                coefficient = 0.36
            elif F < 0.51:
                coefficient = 0.26
            elif F < 0.61:
                coefficient = 0.16
            elif F < 0.71:
                coefficient = 0.09
            else:
                coefficient = 0.04
        else:
            if F < 0.11:
                coefficient = 0.45
            elif F < 0.21:
                coefficient = 0.4
            elif F < 0.31:
                coefficient = 0.35
            elif F < 0.41:
                coefficient = 0.3
            elif F < 0.51:
                coefficient = 0.25
            elif F < 0.61:
                coefficient = 0.2
            elif F < 0.71:
                coefficient = 0.15
            else:
                coefficient = 0.1

        return coefficient

    def get_duct_coords(self, in_tee_con, connector):
        main_con = []
        connector_set = connector.AllRefs.ForwardIterator()
        while connector_set.MoveNext():
            main_con.append(connector_set.Current)
        duct = main_con[0].Owner
        duct_cons = self.get_connectors(duct)
        for duct_con in duct_cons:
            if self.get_con_coords(duct_con) != in_tee_con:
                in_duct_con = self.get_con_coords(duct_con)
                return in_duct_con

    def get_tee_orient(self, element):
        connectors = self.get_connectors(element)
        exit_cons = []
        exhaust_air_cons = []
        for connector in connectors:
            if self.is_supply_air(connectors[0]):
                if connector.Flow != max(connectors[0].Flow, connectors[1].Flow, connectors[2].Flow):
                    exit_cons.append(connector)
            if self.is_exhaust_air(connectors[0]):
                # а что делать если на на разветвлении расход одинаковы?
                if connector.Flow == max(connectors[0].Flow, connectors[1].Flow, connectors[2].Flow):
                    exit_cons.append(connector)
                else:
                    exhaust_air_cons.append(connector)
                # для входа в тройник ищем координаты начала входящего воздуховода чтоб построить прямую через эти две точки

            if connectors[0].DuctSystemType == DuctSystemType.SupplyAir:
                if str(connector.Direction) == "In":
                    in_tee_con = self.get_con_coords(connector)
                    # выбираем из коннектора подключенный воздуховод
                    in_duct_con = get_duct_coords(in_tee_con, connector)

        # в случе вытяжной системы, чтоб выбрать коннектор с выходящим воздухом из второстепенных, берем два коннектора у которых расход не максимальны
        # (максимальный точно выходной у вытяжной системы) и сравниваем. Тот что самый малый - ответветвление
        # а второй - точка вхождения потока воздуха из которой берем координаты для построения вектора

        if (self.is_exhaust_air(connectors[0])):
            if exhaust_air_cons[0].Flow < exhaust_air_cons[1].Flow:
                exit_cons.append(exhaust_air_cons[0])
                in_tee_con = self.get_con_coords(exhaust_air_cons[1])
                in_duct_con = get_duct_coords(in_tee_con, exhaust_air_cons[1])
            else:
                exit_cons.append(exhaust_air_cons[1])
                in_tee_con = self.get_con_coords(exhaust_air_cons[0])
                in_duct_con = get_duct_coords(in_tee_con, exhaust_air_cons[0])

        # среди выходящих коннекторов ищем диктующий по большему расходу
        if exit_cons[0].Flow > exit_cons[1].Flow:
            exit_con = exit_cons[0]
            secondary_con = exit_cons[1]
        else:
            exit_con = exit_cons[1]
            secondary_con = exit_cons[0]

        # диктующий коннектор
        exit_con = self.get_con_coords(exit_con)

        # вторичный коннектор
        secondary_con = self.get_con_coords(secondary_con)

        # найдем вектор по координатам точек AB = {Bx - Ax; By - Ay; Bz - Az}
        duct_to_tee = [(float(in_duct_con[0]) - float(in_tee_con[0])), (float(in_duct_con[1]) - float(in_tee_con[1])),
                       (float(in_duct_con[2]) - float(in_tee_con[2]))]

        tee_to_exit = [(float(in_tee_con[0]) - float(exit_con[0])), (float(in_tee_con[1]) - float(exit_con[1])),
                       (float(in_tee_con[2]) - float(exit_con[2]))]

        # то же самое для вторичного отвода
        tee_to_minor = [(float(in_tee_con[0]) - float(secondary_con[0])),
                        (float(in_tee_con[1]) - float(secondary_con[1])),
                        (float(in_tee_con[2]) - float(secondary_con[2]))]

        # найдем скалярное произведение векторов AB · CD = ABx · CDx + ABy · CDy + ABz · CDz
        tee_to_exit_duct_to_tee = duct_to_tee[0] * tee_to_exit[0] + duct_to_tee[1] * tee_to_exit[1] + duct_to_tee[2] * \
                                  tee_to_exit[2]

        # то же самое с вторичным коннектором
        tee_to_minor_duct_to_tee = duct_to_tee[0] * tee_to_minor[0] + duct_to_tee[1] * tee_to_minor[1] + duct_to_tee[
            2] * tee_to_minor[2]

        # найдем длины векторов
        len_duct_to_tee = ((duct_to_tee[0]) ** 2 + (duct_to_tee[1]) ** 2 + (duct_to_tee[2]) ** 2) ** 0.5
        len_tee_to_exit = ((tee_to_exit[0]) ** 2 + (tee_to_exit[1]) ** 2 + (tee_to_exit[2]) ** 2) ** 0.5

        # то же самое для вторичного вектора
        len_tee_to_minor = ((tee_to_minor[0]) ** 2 + (tee_to_minor[1]) ** 2 + (tee_to_minor[2]) ** 2) ** 0.5

        # найдем косинус
        cos_main = (tee_to_exit_duct_to_tee) / (len_duct_to_tee * len_tee_to_exit)

        # то же самое с вторичным вектором
        cos_minor = (tee_to_minor_duct_to_tee) / (len_duct_to_tee * len_tee_to_minor)

        # Если угол расхождения между вектором входа воздуха и выхода больше 10 градусов(цифра с потолка) то считаем что идет буквой L
        # Если нет, то считаем что идет по прямой буквой I

        # тип 1
        # вытяжной воздуховод zп
        if math.acos(cos_main) < 0.10 and (self.is_exhaust_air(connectors[0])):
            type = 1

        # тип 2
        # вытяжной воздуховод, zо
        elif math.acos(cos_main) > 0.10 and (self.is_exhaust_air(connectors[0])):
            type = 2

        # тип 3
        # подающий воздуховод, zп
        elif math.acos(cos_main) < 0.10 and self.is_supply_air(connectors[0]):
            type = 3

        # тип 4
        # подающий воздуховод, zо
        elif math.acos(cos_main) > 0.10 and self.is_supply_air(connectors[0]):
            type = 4

        return type

    def get_coef_tee(self, element):
        con_set = self.get_connectors(element)

        try:
            type = get_tee_orient(element)
        except Exception:
            type = 'Ошибка'

        try:
            S1 = con_set[0].Height * 0.3048 * con_set[0].Width * 0.3048
        except:
            S1 = 3.14 * 0.3048 * 0.3048 * con_set[0].Radius ** 2
        try:
            S2 = con_set[1].Height * 0.3048 * con_set[1].Width * 0.3048
        except:
            S2 = 3.14 * 0.3048 * 0.3048 * con_set[1].Radius ** 2
        try:
            S3 = con_set[2].Height * 0.3048 * con_set[2].Width * 0.3048
        except:
            S3 = 3.14 * 0.3048 * 0.3048 * con_set[2].Radius ** 2

        if type != 'Ошибка':
            v1 = con_set[0].Flow * 101.94
            v2 = con_set[1].Flow * 101.94
            v3 = con_set[2].Flow * 101.94
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

            f0 = Fo / Fc
            l0 = Lo / Lc
            fp = Fp / Fc

            if f0 < 0.351:
                koef_1 = 0.8 * l0
            elif l0 < 0.61:
                koef_1 = 0.5
            else:
                koef_1 = 0.8 * l0

            if f0 < 0.351:
                koef_2 = 1
            elif l0 < 0.41:
                koef_2 = 0.9 * (1 - l0)
            else:
                koef_2 = 0.55

            if f0 < 0.41:
                koef_3 = 0.4
            elif l0 < 0.51:
                koef_3 = 0.5
            else:
                koef_3 = 2 * l0 - 1

            if f0 < 0.36:
                if l0 < 0.41:
                    koef_4 = 1.1 - 0.7 * l0
                else:
                    koef_4 = 0.85
            else:
                if l0 < 0.61:
                    koef_4 = 1 - 0.6 * l0
                else:
                    koef_4 = 0.6

        if type == 1:
            K = (1 - (1 - l0) ** 2 - (1.4 - l0) * l0 ** 2) / ((1 - l0) ** 2 / fp ** 2)

        if type == 2:
            K = koef_2 * (1 + (l0 / f0) ** 2 - 2 * (1 - l0) ** 2) / (l0 / f0) ** 2
            if K < 0:  # для деталей см. диалог с Коценко, ключевые слова (2653148)
                K = (1 - (1 - l0) ** 2 - (1.4 - l0) * l0 ** 2) / ((1 - l0) ** 2 / fp ** 2)

        if type == 3:
            K = koef_3 * l0 ** 2 / ((1 - l0) ** 2 / Fp ** 2)

        if type == 4:
            K = (koef_4 * (1 + (l0 / f0) ** 2)) / (l0 / f0) ** 2

        if type == 'Ошибка':
            K = 0

        return K

    def get_coef_tap_adjustable(self, element):
        con_set = self.get_connectors(element)

        try:
            Fo = con_set[0].Height * 0.3048 * con_set[0].Width * 0.3048
            form = "Прямоугольный отвод"
        except:
            Fo = 3.14 * 0.3048 * 0.3048 * con_set[0].Radius ** 2
            form = "Круглый отвод"

        main_con = []

        connector_set_0 = con_set[0].AllRefs.ForwardIterator()

        connector_set_1 = con_set[1].AllRefs.ForwardIterator()

        old_flow = 0
        for con in con_set:
            connector_set = con.AllRefs.ForwardIterator()
            while connector_set.MoveNext():
                try:
                    flow = connector_set.Current.Owner.GetParamValue(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
                except Exception:
                    flow = 0
                if flow > old_flow:
                    main_con = []
                    main_con.append(connector_set.Current)
                    old_flow = flow

        duct = main_con[0].Owner

        duct_cons = self.get_connectors(duct)
        Flow = []

        for duct_con in duct_cons:
            Flow.append(duct_con.Flow * 101.94)
            try:
                Fc = duct_con.Height * 0.3048 * duct_con.Width * 0.3048
                Fp = Fc
            except:
                Fc = 3.14 * 0.3048 * 0.3048 * duct_con.Radius ** 2
                Fp = Fc

        Lc = max(Flow)
        Lo = con_set[0].Flow * 101.94

        f0 = Fo / Fc
        l0 = Lo / Lc
        fp = Fp / Fc

        if self.is_exhaust_air(con_set[0]):

            if form == "Круглый отвод":
                if Lc >= Lo * 2:
                    K = ((1 - fp ** 0.5) + 0.5 * l0 + 0.05) * (
                                1.7 + (1 / (2 * f0) - 1) * l0 - ((fp + f0) * l0) ** 0.5) * (
                                fp / (1 - l0)) ** 2
                else:
                    K = (-0.7 - 6.05 * (1 - fp) ** 3) * (f0 / l0) ** 2 + (1.32 + 3.23 * (1 - fp) ** 2) * f0 / l0 + (
                            0.5 + 0.42 * fp) - 0.167 * l0 / f0
            else:
                if Lc >= Lo * 2:
                    K = (fp / (1 - l0)) ** 2 * ((1 - fp) + 0.5 * l0 + 0.05) * (
                            1.5 + (1 / (2 * f0) - 1) * l0 - ((fp + f0) * l0) ** 0.5)
                else:
                    K = (f0 / l0) ** 2 * (4.1 * (fp / f0) ** 1.25 * l0 ** 1.5 * (fp + f0) ** (
                            0.3 * (f0 / fp) ** 0.5 / l0 - 2) - 0.5 * fp / f0)

        if self.is_supply_air(con_set[0]):
            if form == "Круглый отвод":
                if Lc >= Lo * 2:
                    K = 0.45 * (fp / (1 - l0)) ** 2 + (0.6 - 1.7 * fp) * fp / (1 - l0) - (
                                0.25 - 0.9 * fp ** 2) + 0.19 * (
                                1 - l0) / fp
                else:
                    K = (f0 / l0) ** 2 - 0.58 * f0 / l0 + 0.54 + 0.025 * l0 / f0
            else:
                if Lc >= Lo * 2:
                    K = 0.45 * (fp / (1 - l0)) ** 2 + (0.6 - 1.7 * fp) * fp / (1 - l0) - (
                                0.25 - 0.9 * fp ** 2) + 0.19 * (
                                1 - l0) / fp
                else:
                    K = (f0 / l0) ** 2 - 0.42 * f0 / l0 + 0.81 - 0.06 * l0 / f0

        return K
