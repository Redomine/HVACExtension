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
import paraSpec
from Autodesk.Revit.DB import *
from Redomine import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output

doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView

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

def get_fitting_area(element):
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

        area = area * 0.092903
        connectors = getConnectors(element)

        for connector in connectors:
            try:
                H = connector.Height
                B = connector.Width
                S = (H * B) * 0.092903
                area = area - S
            except Exception:
                R = connector.Radius
                S = (3.14 * R * R) * 0.092903
                area = (area - S)
    return area


def get_depend(element):
    parent = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    parent_group = element.LookupParameter('ФОП_ВИС_Группирование').AsString()
    subIds = element.GetSubComponentIds()
    vkheat_collector = []
    for subId in subIds:
        subElement = doc.GetElement(subId)
        part = vkheat_collector_part(element = subElement, ADSK_name= get_ADSK_Name(subElement), ADSK_mark= get_ADSK_Mark(subElement), ADSK_maker = get_ADSK_Maker(subElement), parent = parent, parent_group = parent_group)
        vkheat_collector.append(part)
    vkheat_collector.sort(key=lambda x: x.group)

    number = 0
    old_group = ''

    for part in vkheat_collector:
        if part.isKit:
            if old_group != part.group:
                number += 1
                old_group = part.group


            part.reinsert(number)

class settings:
    def __init__(self,
                 Collection,
                 Group,
                 isSingle):
        self.Collection = Collection
        self.Group = Group
        self.isSingle = isSingle


class vkheat_collector_part:
    def __init__(self, element, ADSK_name, ADSK_mark, ADSK_maker, parent, parent_group):
        self.parent_group = parent_group
        self.element = element
        self.ADSK_maker = ADSK_maker
        self.ADSK_name = ADSK_name
        self.ADSK_mark = ADSK_mark
        self.parent = parent
        self.group = '_Узел_'+self.parent+self.ADSK_name+self.ADSK_mark+self.ADSK_maker
        self.isKit = True

        ElemTypeId = self.element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if ElemType.LookupParameter('ФОП_ВИС_Поставляется отдельно от узла'):
            self.isKit = False
    def reinsert(self, number):
        self.FOP_name = self.element.LookupParameter('ФОП_ВИС_Наименование комбинированное')
        self.FOP_group = self.element.LookupParameter('ФОП_ВИС_Группирование')
        new_group = self.parent_group + self.group

        if (str(number) + '. ') not in self.FOP_name.AsString():
            new_name = "‎    " + str(number) + '. ' + self.FOP_name.AsString()
            self.FOP_name.Set(new_name)


        self.FOP_group.Set(new_group)

def pipe_optimization(size):
    size = str(size)

    old_char = ''
    ind = 0
    for char in size:
        if char == '0' and old_char == '0':
            size = size[:(ind-1)]
            if size[-1] == '.':
                size = size[:-1]
            return size
        old_char = char
        ind += 1
    return size


class shedule_position:
    def shedName(self, element):
        ADSK_Name = self.ADSK_name
        New_Name = ADSK_Name

        try:
            if element.Category.IsId(BuiltInCategory.OST_PipeFitting):
                cons = getConnectors(element)
                for con in cons:

                        for el in con.AllRefs:
                            if el.Owner.Category.IsId(BuiltInCategory.OST_PipeCurves):
                                ElemTypeId = el.Owner.GetTypeId()
                                ElemType = doc.GetElement(ElemTypeId)
                                if ElemType.LookupParameter('ФОП_ВИС_Учитывать фитинги').AsInteger() != 1:
                                    New_Name = '!Не учитывать'
        except Exception:
            pass


        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            external_size = element.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) * 304.8
            internal_size = element.GetParamValue(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM) * 304.8
            pipe_thickness = (external_size - internal_size) / 2

            Dy = str(element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) * 304.8)

            if Dy[-2:] == '.0':
                Dy = Dy[:-2]

            external_size = str(external_size)
            if external_size[-2:] == '.0':
                external_size = external_size[:-2]

            d_type = get_D_type(element)

            if d_type == "Ду":
                New_Name = ADSK_Name + ' ' + 'DN' + pipe_optimization(Dy)
            elif d_type == "Ду х Стенка":
                New_Name = ADSK_Name + ' ' + 'DN' + pipe_optimization(Dy) + 'x' + pipe_optimization(str(pipe_thickness))
            else:
                New_Name = ADSK_Name + ' ' + '⌀' + pipe_optimization(external_size) + 'x' + pipe_optimization(str(pipe_thickness))

        if element.Category.IsId(BuiltInCategory.OST_FlexDuctCurves):
            New_Name = ADSK_Name  + ' ' + element.GetParamValue(BuiltInParameter.RBS_CALCULATED_SIZE)



        if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
            thickness = duct_thickness(element)
            try:
                New_Name = ADSK_Name + ', толщиной ' + thickness + ' мм,' + " " + element.GetParamValue(
                    BuiltInParameter.RBS_CALCULATED_SIZE)
            except Exception:
                New_Name = ADSK_Name + ', толщиной ' + thickness + ' мм,' + " " + element.GetParamValue(
                    BuiltInParameter.RBS_REFERENCE_FREESIZE)

            cons = getConnectors(element)
            for con in cons:
                for el in con.AllRefs:
                    if el.Owner.Category.IsId(BuiltInCategory.OST_DuctInsulations):
                        insType = doc.GetElement(el.Owner.GetTypeId())
                        if insType.LookupParameter('ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
                            insName = get_ADSK_Name(el.Owner)
                            if insName == 'None':
                                insName = 'None_Изоляция'
                            if insName not in New_Name:
                                New_Name = New_Name + " в изоляции " + insName

        if element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
            New_Name = ADSK_Name
            connectors = getConnectors(element)
            for connector in connectors:
                for el in connector.AllRefs:
                    if el.Owner.Category.IsId(BuiltInCategory.OST_PipeFitting):
                        New_Name = '!Не учитывать'
                    if el.Owner.Category.IsId(BuiltInCategory.OST_PipeCurves):
                        pipe_name = el.Owner.LookupParameter('ФОП_ВИС_Наименование комбинированное').AsString()
                        if pipe_name[-1] == '.':
                            pipe_name = pipe_name[:-1]

                        New_Name = ADSK_Name + ' (' + pipe_name + ')'




        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            thickness = duct_thickness(element)
            try:
                connectors = getConnectors(element)
                for connector in connectors:
                    for el in connector.AllRefs:
                        if el.Owner.Category.IsId(BuiltInCategory.OST_DuctCurves):
                            ductType = doc.GetElement(el.Owner.GetTypeId())
                            if ductType.LookupParameter('ФОП_ВИС_Минимальная толщина воздуховода'):
                                min_thickness = ductType.LookupParameter(
                                    'ФОП_ВИС_Минимальная толщина воздуховода').AsDouble()
                                if float(min_thickness) > float(thickness):
                                    thickness = min_thickness

            except Exception:
                pass

            New_Name = 'Металл для фасонных деталей воздуховодов толщиной ' + str(thickness) + ' мм'

            cons = getConnectors(element)
            for con in cons:
                for el in con.AllRefs:
                    if el.Owner.Category.IsId(BuiltInCategory.OST_DuctInsulations):
                        insType = doc.GetElement(el.Owner.GetTypeId())
                        try:
                            if insType.LookupParameter('ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
                                if get_ADSK_Name(el.Owner) not in New_Name:
                                    New_Name = New_Name + " в изоляции " + get_ADSK_Name(el.Owner)
                        except Exception:
                            print
                            insType.Id

        if element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
            insType = doc.GetElement(element.GetTypeId())
            try:
                if insType.LookupParameter('ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
                    New_Name = '!Не учитывать'
            except Exception:
                print
                insType.Id

        return New_Name

    def shedMark(self, element):
        mark = self.ADSK_mark
        if self.ADSK_mark == 'None':
            mark = ''
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            mark = ''
        return mark

    def shedIzm(self, element, ADSK_Izm, isSingle):
        if isSingle:
            if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
                return 'м²'
            return 'шт.'
        else:
            if ADSK_Izm == 'шт' or ADSK_Izm == 'шт.' or ADSK_Izm == 'Шт.' or ADSK_Izm == 'Шт':
                return 'шт.'
            if element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
                if ADSK_Izm == 'м.п.' or ADSK_Izm == 'м.' or ADSK_Izm == 'мп' \
                        or ADSK_Izm == 'м' or ADSK_Izm == 'м.п':
                    return 'м.п.'
                return 'м²'
            if element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
                if ADSK_Izm == 'м2' or ADSK_Izm == 'м²':
                    return 'м²'
                return 'м.п.'
            return 'м.п.'

    def shedNumber(self, element):
        FOP_izm = self.FOP_izm.AsString()
        if FOP_izm == 'шт.':
            return 1
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            return get_fitting_area(element)
        if FOP_izm == 'м.п.':
            length = element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH)
            if length == None:
                length = 0
            else:
                length = (length * 304.8)/1000
            if element.Category.IsId(BuiltInCategory.OST_PipeInsulations) or element.Category.IsId(
                    BuiltInCategory.OST_DuctInsulations):
                length = round((length * isol_reserve), 2)
            else:
                length = round((length * length_reserve), 2)

            return length
        if FOP_izm == 'м²':
            area = element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA)
            if area == None:
                area = 0
            else:
                area = area * 0.092903

            if element.Category.IsId(BuiltInCategory.OST_PipeInsulations) or element.Category.IsId(
                    BuiltInCategory.OST_DuctInsulations):
                area = round((area * isol_reserve), 2)
            else:
                area = round((area * length_reserve), 2)
            return area
        print FOP_izm

    def regroop(self, element):
        new_group = self.paraGroup + "_" + self.FOP_name.AsString() + "_" + self.FOP_Mark.AsString()
        return new_group
    def insert(self):
        if not self.FOP_izm.IsReadOnly:
            self.FOP_izm.Set(self.shedIzm(self.element, self.ADSK_izm, self.isSingle))
        if not self.FOP_name.IsReadOnly:
            self.FOP_name.Set(self.shedName(self.element))
        if not self.FOP_Mark.IsReadOnly:
            self.FOP_Mark.Set(self.shedMark(self.element))
        if not self.FOP_pos.IsReadOnly:
            self.FOP_pos.Set('')
        self.FOP_number.Set(self.shedNumber(self.element))

        if self.FOP_EF == None:
            if not self.FOP_EF.IsReadOnly:
                self.FOP_EF.Set('None')

        if not self.FOP_group.IsReadOnly:
            self.FOP_group.Set(self.regroop(self.element))
        if not self.FOP_maker.IsReadOnly:
            self.FOP_maker.Set(self.ADSK_maker)
        if not self.FOP_code.IsReadOnly:
            self.FOP_code.Set(self.ADSK_code)

    def __init__(self, element, collection, parametric):
        for params in parametric:
            if collection == params.Collection:
                #paraCol = collection
                self.paraGroup = params.Group
                self.isSingle = params.isSingle

        self.element = element
        self.FOP_EF = element.LookupParameter('ФОП_Экономическая функция')
        self.FOP_group = element.LookupParameter('ФОП_ВИС_Группирование')
        self.FOP_name = element.LookupParameter('ФОП_ВИС_Наименование комбинированное')
        self.FOP_number = element.LookupParameter('ФОП_ВИС_Число')
        self.FOP_izm = element.LookupParameter('ФОП_ВИС_Единица измерения')
        self.FOP_Mark = element.LookupParameter('ФОП_ВИС_Марка')
        self.FOP_pos = element.LookupParameter('ФОП_ВИС_Позиция')
        self.FOP_code = element.LookupParameter('ФОП_ВИС_Код изделия')
        self.FOP_maker = element.LookupParameter('ФОП_ВИС_Завод-изготовитель')


        self.ADSK_maker = get_ADSK_Maker(element)
        self.ADSK_name = get_ADSK_Name(element)
        self.ADSK_mark = get_ADSK_Mark(element)
        self.ADSK_izm = get_ADSK_Izm(element)
        self.ADSK_code = get_ADSK_Code(element)

        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ФОП_ВИС_Узел'):
            if ElemType.LookupParameter('ФОП_ВИС_Узел').AsInteger() == 1:
                vis_collectors.append(element)


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

#Коллекция, Категория первичной группы, Единичный элемент?
parametric = [
    settings(colEquipment, '1. Оборудование', True),
    settings(colPlumbingFixtures, '1. Оборудование', True),
    settings(colSprinklers, '1. Оборудование', True),

    settings(colAccessory, '2. Арматура воздуховодов', True),
    settings(colTerminals, '3. Воздухораспределители', True),
    settings(colCurves, '4. Воздуховоды', False),
    settings(colFlexCurves, '4. Гибкие воздуховоды', False),
    settings(colFittings, '5. Фасонные детали воздуховодов', True),
    settings(colInsulations, '6. Материалы изоляции воздуховодов', False),

    settings(colPipeCurves, '7. Трубопроводы', False),
    settings(colFlexPipeCurves, '8. Гибкие трубопроводы', False),
    settings(colPipeAccessory, '9. Трубопроводная арматура', True),
    settings(colPipeFittings, '10. Фасонные детали трубопроводов', True),
    settings(colPipeInsulations, '11. Материалы трубопроводной изоляции', False)
]




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

            data = shedule_position(element, collection, parametric)
            data.insert()


        for element in vis_collectors:
            try:
                edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
            except Exception:
                print element.Id

            if edited_by and edited_by != __revit__.Application.Username:
                report_rows.add(edited_by)
                continue

            get_depend(element)

    if report_rows:
        print "Некоторые элементы не были обработаны, так как были заняты пользователями:"
        print "\r\n".join(report_rows)

parametersAdded = paraSpec.check_parameters()

if not parametersAdded:
    with revit.Transaction("Обновление общей спеки"):
        #список элементов для перебора в вид узлов:
        vis_collectors = []
        # Переменные для расчета
        length_reserve = 1 + (doc.ProjectInformation.LookupParameter(
            'ФОП_ВИС_Запас воздуховодов/труб').AsDouble() / 100)  # запас длин
        isol_reserve = 1 + (
                doc.ProjectInformation.LookupParameter('ФОП_ВИС_Запас изоляции').AsDouble() / 100)  # запас площадей
        script_execute()

    if doc.ProjectInformation.LookupParameter('ФОП_ВИС_Нумерация позиций').AsInteger() == 1 or doc.ProjectInformation.LookupParameter('ФОП_ВИС_Площади воздуховодов в примечания').AsInteger() == 1:
        import numerateSpec