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
from Autodesk.Revit.DB import *
from System import Guid
from itertools import groupby

from pyrevit import revit
from pyrevit.script import output


doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView


def get_ADSK_Izm(element):
    try:
        ADSK_Izm = element.LookupParameter('ADSK_Единица измерения').AsString()
    except Exception:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.get_Parameter(Guid('4289cb19-9517-45de-9c02-5a74ebf5c86d')) == None:
            ADSK_Izm = "None"
        else:
            ADSK_Izm = ElemType.get_Parameter(Guid('4289cb19-9517-45de-9c02-5a74ebf5c86d')).AsString()

    return ADSK_Izm

def get_ADSK_Name(element):
    if element.LookupParameter('ADSK_Наименование'):
        ADSK_Name = element.LookupParameter('ADSK_Наименование').AsString()
        if ADSK_Name == None or ADSK_Name == "":
            ADSK_Name = "None"
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')) == None\
                or ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString() == None \
                or ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString() == "":
            ADSK_Name = "None"
        else:
            ADSK_Name = ElemType.get_Parameter(Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')).AsString()


    return ADSK_Name

def get_ADSK_Mark(element):
    if element.LookupParameter('ADSK_Марка'):
        ADSK_Mark = element.LookupParameter('ADSK_Марка').AsString()
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.get_Parameter(Guid('2204049c-d557-4dfc-8d70-13f19715e46d')) == None \
                or ElemType.get_Parameter(Guid('2204049c-d557-4dfc-8d70-13f19715e46d')).AsString() == None \
                or ElemType.get_Parameter(Guid('2204049c-d557-4dfc-8d70-13f19715e46d')).AsString() == "":
            ADSK_Mark = "None"
        else:
            ADSK_Mark = ElemType.get_Parameter(Guid('2204049c-d557-4dfc-8d70-13f19715e46d')).AsString()
    return ADSK_Mark

def check_collection(collection):
    for element in collection:
        ADSK_Izm = get_ADSK_Izm(element)
        if str(ADSK_Izm) not in Izm_names or ADSK_Izm == None:
            error = 'Для категории ' + str(element.Category.Name) + ' есть элементы без подходящей единицы измерения(м.п., м., мп, м , м.п, шт, шт., м2)' \
                                                                    '\n В такой ситуаци для труб или воздуховодов будут приняты м.п., для изоляции м2'

            if error not in errors_list:
                errors_list.append(error)


def get_D_type(element):
    ElemTypeId = element.GetTypeId()
    ElemType = doc.GetElement(ElemTypeId)

    if ElemType.LookupParameter('ФОП_ВИС_Ду').AsInteger() == 1: type = "Ду"
    elif ElemType.LookupParameter('ФОП_ВИС_Ду х Стенка').AsInteger() == 1: type = "Ду х Стенка"
    else: type = "Днар х Стенка"
    return type

paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Масса', 'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число', 'ФОП_ВИС_Узел', 'ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка',
             'ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ТИП_Назначение', 'ФОП_ТИП_Число', 'ФОП_ТИП_Единица измерения', 'ФОП_ТИП_Код', 'ФОП_ТИП_Наименование работы']

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


# Переменные для расчета
length_reserve = 1 + (doc.ProjectInformation.LookupParameter('ФОП_ВИС_Запас воздуховодов/труб').AsDouble()/100) #запас длин
area_reserve = 1 + (doc.ProjectInformation.LookupParameter('ФОП_ВИС_Запас изоляции').AsDouble()/100)#запас площадей
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
colSprinklers = make_col(BuiltInCategory.OST_Sprinklers)

collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves, colTerminals, colAccessory,
               colPipeAccessory, colEquipment, colInsulations, colPipeInsulations, colPipeCurves, colPlumbingFixtures, colSprinklers]

def duct_thickness(element):
    mode = ''
    if str(element.Category.Name) == 'Воздуховоды':

        a = getConnectors(element)
        try:
            SizeA = a[0].Width * 304.8
            SizeB = a[0].Height * 304.8
            Size = max(SizeA, SizeB)
            mode = 'W'
        except:
            Size = a[0].Radius*2 * 304.8
            mode = 'R'

    if str(element.Category.Name) == 'Соединительные детали воздуховодов':
        a = getConnectors(element)
        if str(element.MEPModel.PartType) == 'Elbow' or str(element.MEPModel.PartType) == 'Cap' \
                or str(element.MEPModel.PartType) == 'TapAdjustable' or str(element.MEPModel.PartType) == 'Union':
            try:
                SizeA = a[0].Width * 304.8
                SizeB = a[0].Height * 304.8
                Size = max(SizeA, SizeB)
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
                SizeA = max(SizeA, SizeA_1)
            except:
                SizeA = a[0].Radius*2*304.8
                circles.append(SizeA)
            try:
                SizeB = a[1].Height * 304.8
                SizeB_1 = a[1].Width * 304.8
                squares.append(SizeB)
                squares.append(SizeB_1)
                SizeB = max(SizeB, SizeB_1)
            except:
                SizeB = a[1].Radius*2* 304.8
                circles.append(SizeB)


            Size = max(SizeA, SizeB)
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

                Size = max(SizeA, SizeB)
                Size = max(Size, SizeC)
                mode = 'W'
            except:
                SizeA = a[0].Radius*2 * 304.8
                SizeB = a[1].Radius*2 * 304.8
                SizeC = a[2].Radius*2 * 304.8

                Size = max(SizeA, SizeB)
                Size = max(Size, SizeC)
                mode = 'R'

        if str(element.MEPModel.PartType) == 'Cross':
            try:
                SizeA = a[0].Width * 304.8
                SizeA_1 = a[0].Height * 304.8
                SizeB = a[1].Width * 304.8
                SizeB_1 = a[1].Height * 304.8
                SizeC = a[2].Width * 304.8
                SizeC_1 = a[2].Height * 304.8
                SizeD = a[3].Width * 304.8
                SizeD_1 = a[3].Height * 304.8

                Size = max(SizeA, SizeB)
                Size = max(Size, SizeC)
                Size = max(Size, SizeD)
                mode = 'W'
            except:
                SizeA = a[0].Radius*2 * 304.8
                SizeB = a[1].Radius*2 * 304.8
                SizeC = a[2].Radius*2 * 304.8
                SizeD = a[3].Radius*2 * 304.8

                Size = max(SizeA, SizeB)
                Size = max(Size, SizeC)
                Size = max(Size, SizeD)
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


def make_new_name(element):
    Spec_Name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное')
    ADSK_Name = get_ADSK_Name(element)

    New_Name = ADSK_Name



    if str(element.Category.Name) == 'Трубы':
        external_size = element.LookupParameter('Внешний диаметр').AsDouble() * 304.8
        internal_size = element.LookupParameter('Внутренний диаметр').AsDouble() * 304.8
        pipe_thickness = (external_size - internal_size)/2
        Dy = str(element.LookupParameter('Диаметр').AsDouble() * 304.8)


        if Dy[-2:] == '.0':
            Dy=Dy[:-2]

        external_size = str(external_size)
        if external_size[-2:] == '.0':
            external_size=external_size[:-2]


        d_type = get_D_type(element)

        if d_type == "Ду":
            New_Name = ADSK_Name + ' ' + 'DN' + Dy
        elif d_type == "Ду х Стенка":
            New_Name = ADSK_Name + ' ' + 'DN' + Dy + 'x' + str(pipe_thickness)
        else:
            New_Name = ADSK_Name + ' ' + '⌀' + external_size + 'x' + str(pipe_thickness)



    if str(element.Category.Name) == 'Воздуховоды':
        thickness = duct_thickness(element)
        try:
            New_Name = ADSK_Name + ', толщиной ' + thickness + ' мм,' + " " + element.LookupParameter('Размер').AsString()
        except Exception:
            New_Name = ADSK_Name + ', толщиной ' + thickness + ' мм,' + " " + element.LookupParameter(
                'Свободный размер').AsString()


    if str(element.Category.Name) == 'Материалы изоляции труб':
        ADSK_Izm = get_ADSK_Izm(element)
        if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
            L = element.LookupParameter('Длина').AsDouble() * 304.8
            S = element.LookupParameter('Площадь').AsDouble() * 0.092903

            pipe = doc.GetElement(element.HostElementId)

            #это на случай если(каким-то образом) изоляция трубы висит без трубы
            try:
                if pipe.LookupParameter('Внешний диаметр') != None:
                    d = pipe.LookupParameter('Внешний диаметр').AsDouble() * 304.8
                    d = str(d)
                    if d[-2:] == '.0':
                        d = d[:-2]
                    New_Name = ADSK_Name + ' внутренним диаметром Ø' + d
            except Exception:
                pass

    if str(element.Category.Name) == 'Соединительные детали воздуховодов':

        try:
            thickness = duct_thickness(element)
        except Exception:
            print str(element.MEPModel.PartType)
            print element.Id

        try:
            connectors = getConnectors(element)
            for connector in connectors:
                if getDuct(connector) != None:
                    ElemTypeId = getDuct(connector).GetTypeId()
                    ElemType = doc.GetElement(ElemTypeId)
                    if ElemType.get_Parameter(Guid('7af80795-5115-46e4-867f-f276a2510250')):
                        min_thickness = ElemType.get_Parameter(Guid('7af80795-5115-46e4-867f-f276a2510250')).AsDouble()
                        if float(min_thickness) > float(thickness):
                            thickness = min_thickness
        except Exception:
            pass



        New_Name = 'Металл для фасонных деталей воздуховодов толщиной ' + str(thickness) + ' мм'


    Spec_Name.Set(str(New_Name))


def getDuct(connector):
    mainCon = []
    connectorSet = connector.AllRefs.ForwardIterator()
    while connectorSet.MoveNext():
        mainCon.append(connectorSet.Current)

    for con in mainCon:
        if con.Owner.LookupParameter('ФОП_ВИС_Группирование'):
            if str(con.Owner.Category.Name) == 'Воздуховоды':
                duct = con.Owner
                return duct





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

#этот блок для элементов с длиной или площадью(учесть что в единицах измерения проекта должны стоять милимметры для длины и м2 для площади) и для расстановки позиции
def getCapacityParam(element, position):

    if element.LookupParameter('ФОП_ВИС_Группирование'):
        Pos = element.LookupParameter('ФОП_ВИС_Группирование')
        Pos.Set(position)
    ADSK_Izm = get_ADSK_Izm(element)

    if ADSK_Izm == 'шт' or ADSK_Izm == 'шт.' or ADSK_Izm == 'Шт.' or ADSK_Izm == 'Шт':
        amount = element.LookupParameter('ФОП_ВИС_Число')
        amount.Set(1)
    else:
        if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
            param = 'Длина'

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
                if CapacityParam == None: pass
                Spec_Length = element.LookupParameter('ФОП_ВИС_Число')
                Spec_Length.Set(CapacityParam)



#этот блок для элементов которые идут поштучно и для расстановки позиции
def getNumericalParam(element, position):
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

    try:
        if str(element.Category.Name) == 'Соединительные детали воздуховодов':
            options = Options()
            geoms = element.get_Geometry(options)

            for g in geoms:
                solids = g.GetInstanceGeometry()
            area = 0

            for solid in solids:
                if isinstance(solid, Line) or isinstance(solid, Arc):
                    continue
                for face in solid.Faces:
                    area = area + face.Area

            connectors = getConnectors(element)

            for connector in connectors:
                try:
                    H = connector.Height
                    B = connector.Width
                    S = H * B
                    area = area - S
                except Exception:
                    R = connector.Radius
                    S = 3.14 * R * R
                    area = area - S

            Spec_Length = element.LookupParameter('ФОП_ВИС_Число')

            Spec_Length.Set(area * 0.092903)

    except Exception:
        Spec_Length = element.LookupParameter('ФОП_ВИС_Число')
        Spec_Length.Set(0)



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




#проверяем заполненность параметров ADSK_Наименование и ADSK_ед. измерения. Единицы еще сверяем со списком допустимых.
errors_list = []
Izm_names = ['м.п.', 'м.', 'мп', 'м', 'м.п', 'шт', 'шт.', 'м2']
check_izm = [colPipeCurves, colCurves, colFlexCurves, colFlexPipeCurves, colInsulations, colPipeInsulations]
for izm in check_izm:
    check_collection(izm)


def update_element(element):
    if element in colEquipment: getNumericalParam(element, '1. Оборудование')
    if element in colPlumbingFixtures:  getNumericalParam(element, '1. Оборудование')
    if element in colSprinklers: getNumericalParam(element, '1. Оборудование')
    if element in colAccessory: getNumericalParam(element, '2. Арматура')
    if element in colTerminals: getNumericalParam(element, '3. Воздухораспределители')
    if element in colPipeAccessory: getNumericalParam(element, '2. Трубопроводная арматура')
    if element in colPipeFittings: getNumericalParam(element, '5. Фасонные детали трубопроводов')
    if element in colFittings: getNumericalParam(element, '5. Фасонные детали воздуховодов')
    if element in colCurves: getCapacityParam(element, '4. Воздуховоды')
    if element in colFlexCurves: getCapacityParam(element, '4. Гибкие воздуховоды')
    if element in colPipeCurves: getCapacityParam(element, '4. Трубопроводы')
    if element in colFlexPipeCurves: getCapacityParam(element, '4. Гибкие трубопроводы')
    if element in colPipeInsulations: getCapacityParam(element, '6. Материалы трубопроводной изоляции')
    if element in colInsulations: getCapacityParam(element, '6. Материалы изоляции воздуховодов')


def update_boq(element):
    fop_name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString()
    adsk_mark = get_ADSK_Mark(element)

    boq_name = element.LookupParameter('ФОП_ТИП_Назначение')
    if adsk_mark == "None":
        boq_name.Set(fop_name)
    else:
        boq_name.Set(fop_name + ' ' + adsk_mark)

    fop_izm = get_ADSK_Izm(element)


    if str(element.Category.Name) == 'Воздуховоды' \
            or str(element.Category.Name) == 'Соединительные детали воздуховодов':
        fop_izm = "м2"

    boq_izm = element.LookupParameter('ФОП_ТИП_Единица измерения')
    if fop_izm == None:
        fop_izm = "None"
    boq_izm.Set(fop_izm)

    fop_number = element.LookupParameter('ФОП_ВИС_Число').AsDouble()

    if str(element.Category.Name) == 'Воздуховоды':
        fop_number = (element.LookupParameter('Площадь').AsDouble() * 0.092903) * area_reserve
        fop_number = round(fop_number, 2)

    boq_number = element.LookupParameter('ФОП_ТИП_Число')
    boq_number.Set(fop_number)

    ElemTypeId = element.GetTypeId()
    ElemType = doc.GetElement(ElemTypeId)
    code = ElemType.LookupParameter('Код по классификатору').AsString()
    boq_code = element.LookupParameter('ФОП_ТИП_Код')
    boq_code.Set(code)

    work_name = ElemType.LookupParameter('Описание по классификатору').AsString()
    boq_work = element.LookupParameter('ФОП_ТИП_Наименование работы')
    boq_work.Set(work_name)

def regroop(element):
    ADSK_Mark = get_ADSK_Mark(element)
    ADSK_Name = get_ADSK_Name(element)
    FOP_Name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString()
    if str(element.Category.Name) != 'Соединительные детали воздуховодов':
        element.LookupParameter('ФОП_ВИС_Группирование').Set(element.LookupParameter('ФОП_ВИС_Группирование').AsString() + " " + FOP_Name + " " + ADSK_Mark)
    else:
        element.LookupParameter('ФОП_ВИС_Группирование').Set(element.LookupParameter('ФОП_ВИС_Группирование').AsString() + " " + FOP_Name)




def script_execute():
    report_rows = set()

    for collection in collections:
        for element in collection:

            edited_by = element.LookupParameter('Редактирует').AsString()
            if edited_by and edited_by != __revit__.Application.Username:
                report_rows.add(edited_by)
                continue

            update_element(element)
            make_new_name(element)
            update_boq(element)

    if len(errors_list) > 0:
        for error in errors_list:
            print error

    for collection in collections:
        for element in collection:
            regroop(element)

    if sort_dependent_by_equipment == True:
        getDependent(colEquipment)
        getDependent(colPlumbingFixtures)
        getDependent(colPipeAccessory)
        getDependent(colAccessory)
        getDependent(colTerminals)



    if report_rows:
        print "Некоторые элементы не были обработаны, так как были заняты пользователями:"
        print "\r\n".join(report_rows)



with revit.Transaction("Обновление общей спеки"):
    script_execute()