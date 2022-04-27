#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление спецификации'
__doc__ = "Обновляет число подсчетных элементов"

import os.path as op

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from rpw.ui.forms import SelectFromList
from System import Guid
from itertools import groupby

from pyrevit import revit
from pyrevit import forms
from pyrevit import script
from pyrevit.forms import Reactive, reactive
from pyrevit.revit import selection, Transaction

doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView


def get_ADSK_Izm(element):
    try:
        ADSK_Izm = element.LookupParameter('ADSK_Наименование').AsString()
    except Exception:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        ADSK_Izm = ElemType.get_Parameter(Guid('4289cb19-9517-45de-9c02-5a74ebf5c86d')).AsString()

    return ADSK_Izm

def get_ADSK_Name(element):
    if element.LookupParameter('ADSK_Наименование'):
        ADSK_Name = element.LookupParameter('ADSK_Наименование').AsString()
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        ADSK_Name = ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString()

    return ADSK_Name

def get_ADSK_Mark(element):
    if element.LookupParameter('ADSK_Наименование'):
        ADSK_Mark = element.LookupParameter('ADSK_Наименование').AsString()
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        ADSK_Mark = ElemType.get_Parameter(Guid('2204049c-d557-4dfc-8d70-13f19715e46d')).AsString()

    return ADSK_Mark

def check_collection(collection):
    for element in collection:
        ADSK_Izm = get_ADSK_Izm(element)
        if ADSK_Izm not in Izm_names or ADSK_Izm == None:
            error = 'Для категории ' + str(element.Category.Name) + ' есть элементы без подходящей единицы измерения(м.п., м., мп, м , м.п, шт, шт., м2)' \
                                                                    '\n В такой ситуаци для труб или воздуховодов будут приняты м.п., для изоляции м2'
            if error not in errors_list:
                errors_list.append(error)


# Переменные для расчета
length_reserve = 1.2 #запас длин
area_reserve = 1.2 #запас площадей
sort_dependent_by_equipment = True #включаем или выключаем сортировку вложенных семейств по их родителям

def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

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

collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves, colTerminals, colAccessory,
               colPipeAccessory, colEquipment, colInsulations, colPipeInsulations, colPipeCurves, colPlumbingFixtures]

def getNumber(Number1, Number2):
    if Number1 > Number2:
        Number = Number1
    else:
        Number = Number2
    return Number

def duct_thickness(element):
    mode = ''

    if str(element.Category.Name) == 'Воздуховоды':
        a = getConnectors(element)
        try:
            SizeA = a[0].Width * 304.8
            SizeB = a[0].Height * 304.8
            Size = getNumber(SizeA, SizeB)
            mode = 'W'
        except:
            Size = a[0].Radius*2 * 304.8
            mode = 'R'

    if str(element.Category.Name) == 'Соединительные детали воздуховодов':
        a = getConnectors(element)
        if str(element.MEPModel.PartType) == 'Elbow':
            try:
                SizeA = a[0].Width * 304.8
                SizeB = a[0].Height * 304.8
                Size = getNumber(SizeA, SizeB)
                mode = 'W'
            except:
                Size = a[0].Radius*2 * 304.8
                mode = 'R'

        if str(element.MEPModel.PartType) == 'Transition':
            circles = []
            squares = []
            try:
                SizeA = a[0].Height * 304.8
                SizeA_1 = a[0].Width * 304.8
                squares.append(SizeA)
                squares.append(SizeA_1)
                SizeA = getNumber(SizeA, SizeA_1)
            except:
                SizeA = a[0].Radius*2*304.8
                circles.append(SizeA)
            try:
                SizeB = a[1].Height * 304.8
                SizeB_1 = a[1].Width * 304.8
                squares.append(SizeB)
                squares.append(SizeB_1)
                SizeB = getNumber(SizeB, SizeB_1)
            except:
                SizeB = a[1].Radius*2* 304.8
                circles.append(SizeB)


            Size = getNumber(SizeA, SizeB)
            if Size in squares: mode = 'W'
            if Size in circles: mode = 'R'

        if str(element.MEPModel.PartType) == 'Tee':
            try:
                SizeA = a[0].Width * 304.8
                SizeA_1 = a[0].Height * 304.8
                SizeB = a[1].Width * 304.8
                SizeB_1 = a[1].Height * 304.8
                SizeC = a[2].Width * 304.8
                SizeC_1 = a[2].Height * 304.8

                Size = getNumber(SizeA, SizeB)
                Size = getNumber(Size, SizeC)
                mode = 'W'
            except:
                SizeA = a[0].Radius*2 * 304.8
                SizeB = a[1].Radius*2 * 304.8
                SizeC = a[2].Radius*2 * 304.8

                Size = getNumber(SizeA, SizeB)
                Size = getNumber(Size, SizeC)
                mode = 'R'

    if mode == 'R':
        if Size < 201:
            thickness = '0.5'
        elif Size < 451:
            thickness = '0.6'
        elif Size < 801:
            thickness = '0.7'
        elif Size < 1251:
            thickness = '1.0'
        elif Size < 1601:
            thickness = '1.2'
        else:
            thickness = '1.4'
    if mode == 'W':
        if Size < 251:
            thickness = '0.5'
        elif Size < 1001:
            thickness = '0.7'
        elif Size < 2001:
            thickness = '0.9'
        else:
            thickness = '1.4'

    dependent = element.GetDependentElements(ElementCategoryFilter(BuiltInCategory.OST_DuctInsulations))
    for elid in dependent:
        el = doc.GetElement(elid)
        ElemTypeId = el.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if ElemType.get_Parameter(Guid('7af80795-5115-46e4-867f-f276a2510250')):
            min_thickness = ElemType.get_Parameter(Guid('7af80795-5115-46e4-867f-f276a2510250')).AsDouble()
            if min_thickness == None: min_thickness = 0
            if float(thickness) < min_thickness: thickness = str(min_thickness)

    ElemTypeId = element.GetTypeId()
    ElemType = doc.GetElement(ElemTypeId)
    if ElemType.get_Parameter(Guid('7af80795-5115-46e4-867f-f276a2510250')):
        min_thickness = ElemType.get_Parameter(Guid('7af80795-5115-46e4-867f-f276a2510250')).AsDouble()
        if min_thickness == None: min_thickness = 0
        if float(thickness) < min_thickness: thickness = str(min_thickness)

    return thickness


def make_new_name(collection):
    for element in collection:
        Spec_Name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное')
        ADSK_Name = get_ADSK_Name(element)

        if ADSK_Name == None:
            ADSK_Name = "Не заполнен ADSK_Наименование"


        New_Name = ADSK_Name

        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '4. Трубопроводы':
            external_size = element.LookupParameter('Внешний диаметр').AsDouble() * 304.8
            internal_size = element.LookupParameter('Внутренний диаметр').AsDouble() * 304.8
            pipe_thickness = (external_size - internal_size)/2
            Dy = element.LookupParameter('Диаметр').AsDouble() * 304.8
            New_Name = ADSK_Name + ' ' + 'Ду='+ str(Dy) + ' (Днар. х т.с. ' + str(external_size) + 'x' + str(pipe_thickness) + ')'

        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '4. Воздуховоды':
            thickness = duct_thickness(element)
            New_Name = ADSK_Name + ' толщиной ' + thickness + ' мм'

        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '6. Материалы трубопроводной изоляции':
            ADSK_Izm = get_ADSK_Izm(element)
            if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
                L = element.LookupParameter('Длина').AsDouble() * 304.8
                S = element.LookupParameter('Площадь').AsDouble() * 0.092903

                pipe = doc.GetElement(element.HostElementId)

                #это на случай если(каким-то образом) изоляция трубы висит без трубы
                try:
                    if pipe.LookupParameter('Внешний диаметр') != None:
                        d = pipe.LookupParameter('Внешний диаметр').AsDouble() * 304.8
                        New_Name = ADSK_Name + ' внутренним диаметром Ø' + str(d)
                except Exception:
                    pass

        if element.LookupParameter('ФОП_ВИС_Группирование').AsString() == '5. Фасонные детали воздуховодов':

            New_Name = ADSK_Name + ' ' + element.LookupParameter('Размер').AsString()
            if str(element.MEPModel.PartType) == 'Elbow' or str(element.MEPModel.PartType) == 'Transition' or str(element.MEPModel.PartType) == 'Tee':
                thickness = duct_thickness(element)
                New_Name = ADSK_Name + ' ' + element.LookupParameter('Размер').AsString() + ' толщиной ' + thickness + ' мм'

        Spec_Name.Set(New_Name)

        ADSK_Izm = get_ADSK_Izm(element)
        if ADSK_Izm != None:
            element.LookupParameter('ФОП_ВИС_Единица измерения').Set(ADSK_Izm)


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

#в этом блоке получаем размеры воздуховодов и труб для наименования в спеке
def getElementSize(element):
    Size = ''
    if element.LookupParameter('Внешний диаметр'):
        outer_size = element.LookupParameter('Внешний диаметр').AsDouble() * 304.8
        interior_size = element.LookupParameter('Внутренний диаметр').AsDouble() * 304.8
        thickness = (float(outer_size) - float(interior_size))/2
        outer_size = str(outer_size)

        a = outer_size.split('.') #убираем 0 после запятой в наружном диаметре если он не имеет значения
        if a[1] == '0':
            outer_size = outer_size.replace(".0","")
        Size = "Ø" + outer_size + "x" + str(thickness)
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)
    elif element.LookupParameter('Размер'):
        Size = element.LookupParameter('Размер').AsString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)

    elif element.LookupParameter('Диаметр'):
        Size = element.LookupParameter('Диаметр').AsValueString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)

    elif element.LookupParameter('Размер трубы'):
        Size = element.LookupParameter('Размер трубы').AsString()
        Spec_Size = element.LookupParameter('ИОС_Размер')
        Spec_Size.Set(Size)

#этот блок для элементов с длиной или площадью(учесть что в единицах измерения проекта должны стоять милимметры для длины и м2 для площади)
def getCapacityParam(collection, position):
    for element in collection:
            if element.LookupParameter('ФОП_ВИС_Группирование'):
                Pos = element.LookupParameter('ФОП_ВИС_Группирование')
                Pos.Set(position)
            ADSK_Izm = get_ADSK_Izm(element)
            if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
                param = 'Длина'
            elif ADSK_Izm == 'шт' or ADSK_Izm == 'шт.':
                amount = element.LookupParameter('ФОП_ВИС_Число')
                amount.Set(1)
                continue
            elif ADSK_Izm == 'м2':
                param = 'Площадь'
            else:
                if position == '6. Материалы трубопроводной изоляции' or position == '6. Материалы изоляции воздуховодов':
                    param = 'Площадь'
                else:
                    param = 'Длина'


            if element.LookupParameter(param):
                if param == 'Длина':
                    CapacityParam = ((element.LookupParameter(param).AsDouble() * 304.8)/1000) * length_reserve
                    CapacityParam = round(CapacityParam, 2)
                else:
                    CapacityParam = (element.LookupParameter(param).AsDouble() * 0.092903) * area_reserve
                    CapacityParam = round(CapacityParam, 2)

                if element.LookupParameter('ФОП_ВИС_Число'):
                    if CapacityParam == None: continue
                    Spec_Length = element.LookupParameter('ФОП_ВИС_Число')
                    Spec_Length.Set(CapacityParam)

#этот блок для элементов которые идут поштучно
def getNumericalParam(collection, position):
    for element in collection:
        try:
            if element.Location:
                if element.LookupParameter('ФОП_ВИС_Группирование'):
                    Pos = element.LookupParameter('ФОП_ВИС_Группирование')
                    Pos.Set(position)

                if element.LookupParameter('ФОП_ВИС_Число'):
                    amount = element.LookupParameter('ФОП_ВИС_Число')
                    amount.Set(1)
        except Exception:
            pass

def getDependent(collection):

    d = {}
    for element in collection:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        chain = ElemType.get_Parameter(Guid('c39ded76-eb20-4a21-abf3-6db36aca369b')).AsValueString()

        if chain == "Нет":
            pass
        else:
            ADSK_Mark = get_ADSK_Mark(element)
            new_group =  element.LookupParameter('ФОП_ВИС_Группирование').AsString() + " " + element.LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString() + " " + ADSK_Mark
            Pos = element.LookupParameter('ФОП_ВИС_Группирование')
            Pos.Set(new_group + " 0")


            dependent = element.GetSubComponentIds()
            numbering = []
            for x in dependent:
                #перебираем вложенные семейства, но вложены могут быть даже обобщенные модели, которые спекой не обрабатываем
                #пока что если выпадает ошибка в считывании наименования просто пропуск, если что будет видно по пустым
                #строкам в спеке
                try:
                    numbering.append(doc.GetElement(x).LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString())
                except Exception:
                    pass
            numbering.sort()
            numbering = [el for el, _ in groupby(numbering)]
            numbering_d = {}

            number = 1
            for name in numbering:
                if name not in numbering_d:
                    numbering_d[name] = number
                    number = number + 1

            for x in dependent:
                #то же что и выше
                try:
                    name = doc.GetElement(x).LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString()
                    new_name = doc.GetElement(x).LookupParameter('ФОП_ВИС_Наименование комбинированное')
                    new_name.Set(str(numbering_d[name]) +'. ' + name)

                    Pos = doc.GetElement(x).LookupParameter('ФОП_ВИС_Группирование')
                    Pos.Set(new_group + " " + str(numbering_d[name]))
                except Exception:
                    pass



paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения', 'ФОП_ВИС_Масса', 'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число']

#проверка на наличие нужных параметров
map = doc.ParameterBindings
it = map.ForwardIterator()
while it.MoveNext():
    newProjectParameterData = it.Key.Name
    if str(newProjectParameterData) in paraNames:
        paraNames.remove(str(newProjectParameterData))
if len(paraNames) > 0:
    print 'Необходимо добавить параметры'
    for name in paraNames:
        print name
    sys.exit()

#проверяем заполненность параметров ADSK_Наименование и ADSK_ед. измерения. Единицы еще сверяем со списком допустимых.
errors_list = []
Izm_names = ['м.п.', 'м.', 'мп', 'м', 'м.п', 'шт', 'шт.' 'м2']
check_izm = [colPipeCurves, colCurves, colFlexCurves, colFlexPipeCurves, colInsulations, colPipeInsulations]
for izm in check_izm:
    check_collection(izm)


if len(errors_list) > 0:
    for error in errors_list:
        print error
    sys.exit()

def execute():
    getNumericalParam(colEquipment, '1. Оборудование')
    getNumericalParam(colPlumbingFixtures, '1. Оборудование')
    getNumericalParam(colAccessory, '2. Арматура')
    getNumericalParam(colTerminals, '3. Воздухораспределители')
    getNumericalParam(colPipeAccessory, '2. Трубопроводная арматура')
    getNumericalParam(colPipeFittings, '5. Фасонные детали трубопроводов')
    getNumericalParam(colFittings, '5. Фасонные детали воздуховодов')
    getCapacityParam(colCurves, '4. Воздуховоды')
    getCapacityParam(colFlexCurves, '4. Гибкие воздуховоды')
    getCapacityParam(colPipeCurves, '4. Трубопроводы')
    getCapacityParam(colFlexPipeCurves, '4. Гибкие трубопроводы')
    getCapacityParam(colPipeInsulations, '6. Материалы трубопроводной изоляции')
    getCapacityParam(colInsulations, '6. Материалы изоляции воздуховодов')

    for collection in collections:
        make_new_name(collection)

    if len(errors_list) > 0:
        for error in errors_list:
            print error

    if sort_dependent_by_equipment == True:
        getDependent(colEquipment)


with revit.Transaction("Обновление общей спеки"):
    execute()