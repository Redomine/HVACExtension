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
    doc = None
    uidoc = None
    view = None

    def __init__(self, doc, uidoc, view):
        self.doc = doc
        self.uidoc = uidoc
        self.view = view

    def get_connectors(self, element):
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

    def get_coef_transition(self, element):
        a = self.get_connectors(element)
        try:
            S1 = a[0].Height * 304.8 * a[0].Width * 304.8
        except:
            S1 = 3.14 * 304.8 * 304.8 * a[0].Radius ** 2
        try:
            S2 = a[1].Height * 304.8 * a[1].Width * 304.8
        except:
            S2 = 3.14 * 304.8 * 304.8 * a[1].Radius ** 2

        # проверяем в какую сторону дует воздух чтоб выяснить расширение это или заужение
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

        F = F0 / F1

        if transition == 'Расширение':
            if F < 0.11:
                K = 0.81
            elif F < 0.21:
                K = 0.64
            elif F < 0.31:
                K = 0.5
            elif F < 0.41:
                K = 0.36
            elif F < 0.51:
                K = 0.26
            elif F < 0.61:
                K = 0.16
            elif F < 0.71:
                K = 0.09
            else:
                K = 0.04
        if transition == 'Заужение':
            if F < 0.11:
                K = 0.45
            elif F < 0.21:
                K = 0.4
            elif F < 0.31:
                K = 0.35
            elif F < 0.41:
                K = 0.3
            elif F < 0.51:
                K = 0.25
            elif F < 0.61:
                K = 0.2
            elif F < 0.71:
                K = 0.15
            else:
                K = 0.1

        return K

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
            if str(connectors[0].DuctSystemType) == "SupplyAir":
                if connector.Flow != max(connectors[0].Flow, connectors[1].Flow, connectors[2].Flow):
                    exit_cons.append(connector)
            if str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir":
                # а что делать если на на разветвлении расход одинаковы?
                if connector.Flow == max(connectors[0].Flow, connectors[1].Flow, connectors[2].Flow):
                    exit_cons.append(connector)
                else:
                    exhaust_air_cons.append(connector)
                # для входа в тройник ищем координаты начала входящего воздуховода чтоб построить прямую через эти две точки

            if str(connectors[0].DuctSystemType) == "SupplyAir":
                if str(connector.Direction) == "In":
                    in_tee_con = self.get_con_coords(connector)
                    # выбираем из коннектора подключенный воздуховод
                    in_duct_con = get_duct_coords(in_tee_con, connector)

        # в случе вытяжной системы, чтоб выбрать коннектор с выходящим воздухом из второстепенных, берем два коннектора у которых расход не максимальны
        # (максимальный точно выходной у вытяжной системы) и сравниваем. Тот что самый малый - ответветвление
        # а второй - точка вхождения потока воздуха из которой берем координаты для построения вектора

        if str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir":

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
        if math.acos(cos_main) < 0.10 and (
                str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir"):
            type = 1

        # тип 2
        # вытяжной воздуховод, zо
        elif math.acos(cos_main) > 0.10 and (
                str(connectors[0].DuctSystemType) == "ExhaustAir" or str(connectors[0].DuctSystemType) == "ReturnAir"):
            type = 2

        # тип 3
        # подающий воздуховод, zп
        elif math.acos(cos_main) < 0.10 and str(connectors[0].DuctSystemType) == "SupplyAir":
            type = 3

        # тип 4
        # подающий воздуховод, zо
        elif math.acos(cos_main) > 0.10 and str(connectors[0].DuctSystemType) == "SupplyAir":
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

        if str(con_set[0].DuctSystemType) == "ExhaustAir" or str(con_set[0].DuctSystemType) == "ReturnAir":
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

        if str(con_set[0].DuctSystemType) == "SupplyAir":
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
