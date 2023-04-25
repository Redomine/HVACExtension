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

    if lookupCheck(ElemType, 'ФОП_ВИС_Ду').AsInteger() == 1: type = "Ду"
    elif lookupCheck(ElemType, 'ФОП_ВИС_Ду х Стенка').AsInteger() == 1: type = "Ду х Стенка"
    else: type = "Днар х Стенка"
    return type




def duct_thickness(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        a = getConnectors(element)
        try:
            SizeA = fromRevitToMilimeters(a[0].Width)
            SizeB = fromRevitToMilimeters(a[0].Height)
            Size = max(SizeA, SizeB)

            if Size < 251:
                thickness = '0.5'
            elif Size < 1001:
                thickness = '0.7'
            elif Size < 2001:
                thickness = '0.9'
            else:
                thickness = '1.4'

        except:
            Size = fromRevitToMilimeters(a[0].Radius*2)

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

    dependent = element.GetDependentElements(ElementCategoryFilter(BuiltInCategory.OST_DuctInsulations))
    for elid in dependent:
        el = doc.GetElement(elid)
        ElemTypeId = el.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if getParameter(ElemType, 'ФОП_ВИС_Минимальная толщина воздуховода'):
            min_thickness = lookupCheck(ElemType, 'ФОП_ВИС_Минимальная толщина воздуховода').AsDouble()
            if min_thickness == None: min_thickness = 0
            if float(thickness) < min_thickness: thickness = str(min_thickness)

    ElemTypeId = element.GetTypeId()
    ElemType = doc.GetElement(ElemTypeId)
    if getParameter(ElemType, 'ФОП_ВИС_Минимальная толщина воздуховода'):
        min_thickness = lookupCheck(ElemType, 'ФОП_ВИС_Минимальная толщина воздуховода').AsDouble()
        if min_thickness == None: min_thickness = 0
        if float(thickness) < min_thickness: thickness = str(min_thickness)

    return thickness

def getDuct(connector):
    mainCon = []
    connectorSet = connector.AllRefs.ForwardIterator()
    while connectorSet.MoveNext():
        mainCon.append(connectorSet.Current)

    for con in mainCon:
        if getParameter(con.Owner, 'ФОП_ВИС_Группирование'):
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

        area = fromRevitToSquareMeters(area)
        connectors = getConnectors(element)

        for connector in connectors:
            try:
                H = connector.Height
                B = connector.Width
                S = (H * B)
                S = fromRevitToSquareMeters(S)
                area = area - S
            except Exception:
                R = connector.Radius
                S = (3.14 * R * R)
                S = fromRevitToSquareMeters(S)
                area = (area - S)
        area = round(area, 2)
    return area


def get_depend(element):
    parent = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    parent_group = lookupCheck(element,'ФОП_ВИС_Группирование').AsString()

    subIds = element.GetSubComponentIds()
    vkheat_collector = []
    for subId in subIds:
        subElement = doc.GetElement(subId)
        part = vkheat_collector_part(element = subElement, ADSK_name= get_ADSK_Name(subElement),
                                     ADSK_mark= get_ADSK_Mark(subElement), ADSK_maker = get_ADSK_Maker(subElement),
                                     parent = parent, parent_group = parent_group)
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

        if getParameter(ElemType, 'ФОП_ВИС_Поставляется отдельно от узла'):
            self.isKit = False
    def reinsert(self, number):
        self.FOP_name = lookupCheck(self.element, 'ФОП_ВИС_Наименование комбинированное')
        self.FOP_group =lookuoCheck(self.element, 'ФОП_ВИС_Группирование')
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
                                if lookupCheck(ElemType, 'ФОП_ВИС_Учитывать фитинги').AsInteger() != 1:
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

        if element.Category.IsId(BuiltInCategory.OST_FlexPipeCurves):
            Dy = str(element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) * 304.8)

            if Dy[-2:] == '.0':
                Dy = Dy[:-2]
            New_Name = ADSK_Name + ' ' + 'DN' + pipe_optimization(Dy)


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
                        if lookupCheck(insType, 'ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
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
                    if el.Owner.Category.IsId(BuiltInCategory.OST_PipeCurves):
                        pipe_name = lookupCheck(el.Owner, 'ФОП_ВИС_Наименование комбинированное').AsString()
                        if pipe_name == None:
                            continue
                        if pipe_name[-1] == '.':
                            pipe_name = pipe_name[:-1]

                        New_Name = ADSK_Name + ' (' + pipe_name + ')'
                    else:
                        New_Name = '!Не учитывать'



        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            cons = getConnectors(element)

            for con in cons:
                ductAround = False
                for el in con.AllRefs:

                    if el.Owner.Category.IsId(BuiltInCategory.OST_DuctCurves):
                        ductAround = True
                        thickness = duct_thickness(el.Owner)
                        New_Name = 'Металл для фасонных деталей воздуховодов толщиной ' + str(thickness) + ' мм'


                    if el.Owner.Category.IsId(BuiltInCategory.OST_DuctInsulations):
                        insType = doc.GetElement(el.Owner.GetTypeId())
                        try:
                            if lookupCheck(insType, 'ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
                                if get_ADSK_Name(el.Owner) not in New_Name:
                                    New_Name = New_Name + " в изоляции " + get_ADSK_Name(el.Owner)
                                    ductAround = True
                        except Exception:
                            pass

            for con in cons:
                if not ductAround:
                    for el in con.AllRefs:
                        if el.Owner.Category.IsId(BuiltInCategory.OST_DuctFitting):
                            New_Name = el.Owner.LookupParameter("ФОП_ВИС_Наименование комбинированное").AsString()






        if element.Category.IsId(BuiltInCategory.OST_DuctInsulations):

            insType = doc.GetElement(element.GetTypeId())
            try:
                if lookupCheck(insType, 'ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
                    New_Name = '!Не учитывать'
            except Exception:
                pass

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
            if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
                if ADSK_Izm == 'м2' or ADSK_Izm == 'м²':
                    return 'м²'
                return 'м.п.'
            return 'м.п.'

    def shedNumber(self, element):
        # Переменные для расчета

        length_reserve = 1 + (lookupCheck(information, 'ФОП_ВИС_Запас воздуховодов/труб').AsDouble() / 100)  # запас длин
        isol_reserve = 1 + (lookupCheck(information, 'ФОП_ВИС_Запас изоляции').AsDouble() / 100)  # запас площадей


        if self.stock != 0:
            isol_reserve = 1 + self.stock / 100
            length_reserve = 1 + self.stock / 100





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
                length = fromRevitToMeters(length)
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
                area = fromRevitToSquareMeters(area)

            if element.Category.IsId(BuiltInCategory.OST_PipeInsulations) or element.Category.IsId(
                    BuiltInCategory.OST_DuctInsulations):
                area = round((area * isol_reserve), 2)
            else:
                area = round((area * length_reserve), 2)

            if element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
                connectors = getConnectors(element)
                for connector in connectors:
                    for el in connector.AllRefs:
                        if el.Owner.Category.IsId(BuiltInCategory.OST_DuctFitting):

                            return round((get_fitting_area(el.Owner)* isol_reserve), 2)
            return area


    def regroop(self, element):
        new_group = self.paraGroup + "_" + self.FOP_name.AsString() + "_" + self.FOP_Mark.AsString()
        return new_group

    def isDataToInsert(self, param, value):
        if param:
            if not param.IsReadOnly:
                if value == None:
                    value = 'None'
                param.Set(value)


    def insert(self):

        code = self.ADSK_code
        unit = self.shedIzm(self.element, self.ADSK_izm, self.isSingle)
        self.isDataToInsert(self.FOP_izm, unit)
        name = self.shedName(self.element)
        self.isDataToInsert(self.FOP_name, name)
        mark = self.shedMark(self.element)
        self.isDataToInsert(self.FOP_Mark, mark)
        self.isDataToInsert(self.FOP_pos, '')
        number = self.shedNumber(self.element)

        self.isDataToInsert(self.FOP_number, number)
        group = self.regroop(self.element)
        self.isDataToInsert(self.FOP_group, group)
        maker = self.ADSK_maker
        self.isDataToInsert(self.FOP_maker, maker)
        self.isDataToInsert(self.FOP_code, code)

        if self.FOP_EF:
            if self.FOP_EF.AsString() == None or self.FOP_EF.AsString() == '':
                self.isDataToInsert(self.FOP_EF, 'None')

        if self.FOP_System:
            if self.FOP_System.AsString() == None or self.FOP_System.AsString() == '':
                self.isDataToInsert(self.FOP_System, 'None')



    def __init__(self, element, collection, parametric):
        for params in parametric:
            if collection == params.Collection:
                #paraCol = collection
                self.paraGroup = params.Group
                self.isSingle = params.isSingle

        self.element = element


        self.FOP_System = lookupCheck(element, 'ФОП_ВИС_Имя системы')
        self.FOP_EF = lookupCheck(element,'ФОП_Экономическая функция')
        self.FOP_group = lookupCheck(element, 'ФОП_ВИС_Группирование')
        self.FOP_name = lookupCheck(element, 'ФОП_ВИС_Наименование комбинированное')
        self.FOP_number = lookupCheck(element, 'ФОП_ВИС_Число')
        self.FOP_izm = lookupCheck(element, 'ФОП_ВИС_Единица измерения')
        self.FOP_Mark = lookupCheck(element, 'ФОП_ВИС_Марка')
        self.FOP_pos = lookupCheck(element, 'ФОП_ВИС_Позиция')
        self.FOP_code = lookupCheck(element, 'ФОП_ВИС_Код изделия')
        self.FOP_maker = lookupCheck(element, 'ФОП_ВИС_Завод-изготовитель')



        self.ADSK_maker = get_ADSK_Maker(element)
        self.ADSK_name = get_ADSK_Name(element)

        self.ADSK_mark = get_ADSK_Mark(element)
        self.ADSK_izm = get_ADSK_Izm(element)
        self.ADSK_code = get_ADSK_Code(element)
        self.stock = 0

        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if getParameter(ElemType, 'ФОП_ВИС_Индивидуальный запас'):
            self.stock = lookupCheck(ElemType, 'ФОП_ВИС_Индивидуальный запас').AsDouble()

        if getParameter(ElemType, 'ФОП_ВИС_Узел'):
            if lookupCheck(ElemType, 'ФОП_ВИС_Узел').AsInteger() == 1:
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

    settings(colPipeAccessory, '7. Трубопроводная арматура', True),
    settings(colPipeCurves, '8. Трубопроводы', False),
    settings(colFlexPipeCurves, '8. Гибкие трубопроводы', False),
    settings(colPipeFittings, '10. Фасонные детали трубопроводов', True),
    settings(colPipeInsulations, '11. Материалы трубопроводной изоляции', False)
]

report_rows = []
def isElementEditedBy(element):
    try:
        edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
    except Exception:
        edited_by = __revit__.Application.Username

    if edited_by and edited_by != __revit__.Application.Username:
        if edited_by not in report_rows:
            report_rows.append(edited_by)
        return True
    return False


def script_execute():
    for collection in collections:
        for element in collection:
            if not isElementEditedBy(element):
                data = shedule_position(element, collection, parametric)
                data.insert()


        for element in vis_collectors:
            if not isElementEditedBy(element):
                get_depend(element)



if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

parametersAdded = paraSpec.check_parameters()



if not parametersAdded:
    information = doc.ProjectInformation
    with revit.Transaction("Обновление девятиграфной формы"):
        #список элементов для перебора в вид узлов:
        vis_collectors = []
        script_execute()
        for report in report_rows:
            print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report

    if lookupCheck(information, 'ФОП_ВИС_Нумерация позиций').AsInteger() == 1 or lookupCheck(information, 'ФОП_ВИС_Площади воздуховодов в примечания').AsInteger() == 1:
        import numerateSpec