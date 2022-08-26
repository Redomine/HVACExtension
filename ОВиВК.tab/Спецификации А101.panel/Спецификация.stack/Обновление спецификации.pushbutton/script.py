#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Обновление спецификации'
__doc__ = "Обновляет число подсчетных элементов"

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

from Autodesk.Revit.DB import *
from System import Guid
from itertools import groupby

from pyrevit import revit
from pyrevit.script import output






doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView


def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

def get_ADSK_Izm(element):
    try:
        ADSK_Izm = element.LookupParameter('ADSK_Единица измерения').AsString()
    except Exception:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Единица измерения') == None:
            ADSK_Izm = "None"
        else:
            ADSK_Izm = ElemType.LookupParameter('ADSK_Единица измерения').AsString()

    return ADSK_Izm

def get_ADSK_Name(element):
    if element.LookupParameter('ADSK_Наименование'):
        ADSK_Name = element.LookupParameter('ADSK_Наименование').AsString()
        if ADSK_Name == None or ADSK_Name == "":
            ADSK_Name = "None"
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Наименование') == None\
                or ElemType.LookupParameter('ADSK_Наименование').AsString() == None \
                or ElemType.LookupParameter('ADSK_Наименование').AsString() == "":
            ADSK_Name = "None"
        else:
            ADSK_Name = ElemType.LookupParameter('ADSK_Наименование').AsString()

    if str(element.Category.Name) == 'Трубы':
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента'):
            if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента').AsInteger() == 1:
                ADSK_Name = element.LookupParameter('Описание сегмента').AsString()



    return ADSK_Name

def get_ADSK_Mark(element):
    if element.LookupParameter('ADSK_Марка'):
        ADSK_Mark = element.LookupParameter('ADSK_Марка').AsString()
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Марка') == None \
                or ElemType.LookupParameter('ADSK_Марка').AsString() == None \
                or ElemType.LookupParameter('ADSK_Марка').AsString() == "":
            ADSK_Mark = "None"
        else:
            ADSK_Mark = ElemType.LookupParameter('ADSK_Марка').AsString()
    return ADSK_Mark

def get_D_type(element):
    ElemTypeId = element.GetTypeId()
    ElemType = doc.GetElement(ElemTypeId)

    if ElemType.LookupParameter('ФОП_ВИС_Ду').AsInteger() == 1: type = "Ду"
    elif ElemType.LookupParameter('ФОП_ВИС_Ду х Стенка').AsInteger() == 1: type = "Ду х Стенка"
    else: type = "Днар х Стенка"
    return type

def duct_thickness(element):
    mode = ''
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        a = getConnectors(element)
        try:
            SizeA = a[0].Width * 304.8
            SizeB = a[0].Height * 304.8
            Size = max(SizeA, SizeB)
            mode = 'W'
        except:
            Size = a[0].Radius*2 * 304.8
            mode = 'R'

    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
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

        if ElemType.LookupParameter('ФОП_ВИС_Минимальная толщина воздуховода'):
            min_thickness = ElemType.LookupParameter('ФОП_ВИС_Минимальная толщина воздуховода').AsDouble()
            if min_thickness == None: min_thickness = 0
            if float(thickness) < min_thickness: thickness = str(min_thickness)

    ElemTypeId = element.GetTypeId()
    ElemType = doc.GetElement(ElemTypeId)
    if ElemType.LookupParameter('ФОП_ВИС_Минимальная толщина воздуховода'):
        min_thickness = ElemType.LookupParameter('ФОП_ВИС_Минимальная толщина воздуховода').AsDouble()
        if min_thickness == None: min_thickness = 0
        if float(thickness) < min_thickness: thickness = str(min_thickness)

    return thickness

def make_new_name(element):
    Spec_Name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное')
    ADSK_Name = get_ADSK_Name(element)
    New_Name = ADSK_Name


    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        external_size = element.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) * 304.8
        internal_size = element.GetParamValue(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM) * 304.8
        pipe_thickness = (external_size - internal_size)/2

        Dy = str(element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) * 304.8)

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

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        thickness = duct_thickness(element)
        try:
            New_Name = ADSK_Name + ', толщиной ' + thickness + ' мм,' + " " + element.GetParamValue(BuiltInParameter.RBS_CALCULATED_SIZE)
        except Exception:
            New_Name = ADSK_Name + ', толщиной ' + thickness + ' мм,' + " " + element.GetParamValue(BuiltInParameter.RBS_REFERENCE_FREESIZE)

    if element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
        ADSK_Izm = get_ADSK_Izm(element)
        if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
            lenght = element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH)
            if lenght == None:
                lenght = 0
            area = element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA)
            if area == None:
                area = 0
            L = lenght * 304.8
            S = area * 0.092903

            pipe = doc.GetElement(element.HostElementId)

            #это на случай если(каким-то образом) изоляция трубы висит без трубы
            try:
                if pipe.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) != None:
                    d = pipe.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) * 304.8
                    d = str(d)
                    if d[-2:] == '.0':
                        d = d[:-2]
                    New_Name = ADSK_Name + ' внутренним диаметром Ø' + d
            except Exception:
                pass


    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        thickness = duct_thickness(element)

        try:
            connectors = getConnectors(element)
            for connector in connectors:
                if getDuct(connector) != None:
                    ElemTypeId = getDuct(connector).GetTypeId()
                    ElemType = doc.GetElement(ElemTypeId)
                    if ElemType.LookupParameter('ФОП_ВИС_Минимальная толщина воздуховода'):
                        min_thickness = ElemType.LookupParameter('ФОП_ВИС_Минимальная толщина воздуховода').AsDouble()
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
            if str(con.Owner.Category.IsId(BuiltInCategory.OST_DuctCurves)):
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

#этот блок для элементов с длиной или площадью(учесть что в единицах измерения проекта должны стоять милимметры для длины и м2 для площади) и для расстановки позиции
def getCapacityParam(element, position):
    fop_izm = element.LookupParameter('ФОП_ВИС_Единица измерения')

    if element.LookupParameter('ФОП_ВИС_Группирование'):
        Pos = element.LookupParameter('ФОП_ВИС_Группирование')
        Pos.Set(position)
    ADSK_Izm = get_ADSK_Izm(element)

    if ADSK_Izm == 'шт' or ADSK_Izm == 'шт.' or ADSK_Izm == 'Шт.' or ADSK_Izm == 'Шт':
        fop_izm.Set('шт.')
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
                fop_izm.Set('м.п.')

                lenght = element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH)
                if lenght == None:
                    lenght = 0

                CapacityParam = ((lenght * 304.8)/1000) * length_reserve
                CapacityParam = round(CapacityParam, 2)
            else:
                fop_izm.Set('м²')
                area = element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA)
                if area == None:
                    area = 0

                CapacityParam = (area * 0.092903) * area_reserve
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

    fop_izm = element.LookupParameter('ФОП_ВИС_Единица измерения')
    fop_izm.Set('шт.')

    try:
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
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

#Обновляем параметры для ВОР и перебиваем единицы измерения
def update_boq(element):


    fop_name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString()

    adsk_mark = get_ADSK_Mark(element)

    boq_name = element.LookupParameter('ФОП_ТИП_Назначение')
    if adsk_mark == "None":
        boq_name.Set(fop_name)
    else:
        boq_name.Set(fop_name + ' ' + adsk_mark)

    boq_izm = element.LookupParameter('ФОП_ТИП_Единица измерения')
    fop_izm = element.LookupParameter('ФОП_ВИС_Единица измерения')


    adsk_izm = get_ADSK_Izm(element)

    if adsk_izm == None:
        adsk_izm = "None"


    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        fop_izm.Set("м²")
        boq_izm.Set("м²")
    elif element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        boq_izm.Set("м²")
    else:
        boq_izm.Set(fop_izm.AsString())


    fop_number = element.LookupParameter('ФОП_ВИС_Число').AsDouble()

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        fop_number = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903) * area_reserve
        fop_number = round(fop_number, 2)

    boq_number = element.LookupParameter('ФОП_ТИП_Число')
    boq_number.Set(fop_number)

    ElemTypeId = element.GetTypeId()
    ElemType = doc.GetElement(ElemTypeId)

    fop_number = round(fop_number, 2)
    code = str(ElemType.GetParamValue(BuiltInParameter.UNIFORMAT_CODE))
    if code == None:
        code = ""
    boq_code = element.LookupParameter('ФОП_ТИП_Код')
    boq_code.Set(code)

    work_name = ElemType.GetParamValue(BuiltInParameter.UNIFORMAT_DESCRIPTION)
    if work_name == None:
        work_name = ""
    boq_work = element.LookupParameter('ФОП_ТИП_Наименование работы')
    boq_work.Set(work_name)

def regroop(element):
    ADSK_Mark = get_ADSK_Mark(element)
    ADSK_Name = get_ADSK_Name(element)
    FOP_Name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString()
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        element.LookupParameter('ФОП_ВИС_Группирование').Set(
            element.LookupParameter('ФОП_ВИС_Группирование').AsString() + " " + FOP_Name)
    else:
        element.LookupParameter('ФОП_ВИС_Группирование').Set(
            element.LookupParameter('ФОП_ВИС_Группирование').AsString() + " " + FOP_Name + " " + ADSK_Mark)



def script_execute():
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

            update_element(element)
            make_new_name(element)
            update_boq(element)



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


paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения' ,'ФОП_ВИС_Масса', 'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число', 'ФОП_ВИС_Узел', 'ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка',
             'ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ТИП_Назначение', 'ФОП_ТИП_Число', 'ФОП_ТИП_Единица измерения',
             'ФОП_ТИП_Код', 'ФОП_ТИП_Наименование работы', 'ФОП_ВИС_Имя трубы из сегмента', 'ФОП_ВИС_Позиция', 'ФОП_ВИС_Площади воздуховодов в примечания',
             'ФОП_ВИС_Нумерация позиций']



#проверка на наличие нужных параметров
map = doc.ParameterBindings
it = map.ForwardIterator()
while it.MoveNext():
    newProjectParameterData = it.Key.Name
    if str(newProjectParameterData) in paraNames:
        paraNames.remove(str(newProjectParameterData))
if len(paraNames) > 0:
    try:
        import paraSpec
        print 'Были добавлен параметры, перезапустите скрипт'
    except Exception:
        print 'Не удалось добавить параметры'
else:
    # Переменные для расчета
    length_reserve = 1 + (doc.ProjectInformation.LookupParameter('ФОП_ВИС_Запас воздуховодов/труб').AsDouble()/100) #запас длин
    area_reserve = 1 + (doc.ProjectInformation.LookupParameter('ФОП_ВИС_Запас изоляции').AsDouble()/100)#запас площадей
    sort_dependent_by_equipment = True #включаем или выключаем сортировку вложенных семейств по их родителям


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


    with revit.Transaction("Обновление общей спеки"):
        script_execute()


    if doc.ProjectInformation.LookupParameter('ФОП_ВИС_Нумерация позиций').AsInteger() == 1 or doc.ProjectInformation.LookupParameter('ФОП_ВИС_Нумерация позиций').AsInteger() == 1:
        import numerateSpec