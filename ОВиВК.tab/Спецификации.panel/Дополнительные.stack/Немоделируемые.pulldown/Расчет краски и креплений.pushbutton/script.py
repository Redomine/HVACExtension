#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Расчет краски и креплений'
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")


import sys
import System
import dosymep
import paraSpec


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)


from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from Redomine import *


from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert





doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

    
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)
colSystems = make_col(BuiltInCategory.OST_DuctSystem)
colInsul = make_col(BuiltInCategory.OST_DuctInsulations)


class generationElement:
    def __init__(self, group, name, mark, art, maker, unit, method, collection, isType):
        self.group = group
        self.name = name
        self.mark = mark
        self.maker = maker
        self.unit = unit
        self.collection = collection
        self.method = method
        self.isType = isType
        self.art = art


genList = [
    generationElement(group = '12. Расчетные элементы', name = "Металлические крепления для воздуховодов", mark = '', art = '', unit = 'кг.', maker = '',method = 'ФОП_ВИС_Расчет металла для креплений', collection=colCurves,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Металлические крепления для трубопроводов", mark = '', art = '', unit = 'кг.', maker = '', method =  'ФОП_ВИС_Расчет металла для креплений', collection= colPipes,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Изоляция для фланцев и стыков", mark = '', art = '', unit = 'м².', maker = '', method =  'ФОП_ВИС_Совместно с воздуховодом', collection= colInsul,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Краска антикоррозионная за два раза", mark = 'БТ-177', art = '', unit = 'кг.', maker = '', method =  'ФОП_ВИС_Расчет краски и грунтовки', collection= colPipes,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Грунтовка для стальных труб", mark = 'ГФ-031', art = '', unit = 'кг.', maker = '', method =  'ФОП_ВИС_Расчет краски и грунтовки', collection= colPipes,isType= False)
]


class calculation_element:
    def duct_material(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903) / 100
        if element.GetParamValue(BuiltInParameter.RBS_EQ_DIAMETER_PARAM) == element.GetParamValue(
                BuiltInParameter.RBS_HYDRAULIC_DIAMETER_PARAM):
            D = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            P = 3.14 * D
        else:
            A = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            B = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            P = 2 * (A + B)

        if P < 1001:
            kg = area * 65
        elif P < 1801:
            kg = area * 122
        else:
            kg = area * 225

        return kg

    def pipe_material(self, element):
        lenght = (304.8 * element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH)) / 1000

        D = 304.8 * element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)

        if D < 25:
            kg = 0.9 * lenght
        elif D < 33:
            kg = 0.73 * lenght
        elif D < 41:
            kg = 0.64 * lenght
        elif D < 51:
            kg = 0.67 * lenght
        elif D < 66:
            kg = 0.53 * lenght
        elif D < 81:
            kg = 0.7 * lenght
        elif D < 101:
            kg = 0.64 * lenght
        elif D < 126:
            kg = 1.16 * lenght
        else:
            kg = 0.96 * lenght
        return kg

    def insul_stock(self, element):
        area = element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA)
        if area == None:
            area = 0
        area = area * 0.092903 * 0.03
        return area

    def grunt(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area / 10
        return number

    def colorBT(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area * 0.2 * 2
        return number

    def get_number(self, element, name):
        Number = 1
        if name == "Металлические крепления для трубопроводов" and element in colPipes:
            Number = self.pipe_material(element)
        if name == "Металлические крепления для воздуховодов" and element in colCurves:
            Number = self.duct_material(element)
        if name == "Изоляция для фланцев и стыков" and element in colInsul:
            Number = self.insul_stock(element)
        if name == "Краска антикоррозионная за два раза" and element in colPipes:
            Number = self.colorBT(element)
        if name == "Грунтовка для стальных труб" and element in colPipes:
            Number = self.grunt(element)
        return Number

    def __init__(self, element, collection, parameter, Name, Mark, Maker):

        self.corp = ''
        self.sec = ''
        self.floor = ''
        if element.LookupParameter('ФОП_ВИС_Имя системы'):
            self.system = str(element.LookupParameter('ФОП_ВИС_Имя системы').AsString())
        else:
            try:
                self.system = str(element.LookupParameter('ADSK_Имя системы').AsString())
            except:
                self.system = 'None'
        self.group ='12. Расчетные элементы'
        self.name = Name
        self.mark = Mark
        self.art = ''
        self.maker = Maker
        self.izm = 'None'
        self.number = self.get_number(element, self.name)
        self.mass = ''
        self.comment = ''
        self.EF = str(element.LookupParameter('ФОП_Экономическая функция').AsString())
        self.parentId = element.Id.IntegerValue


        for gen in genList:
            if gen.collection == collection and parameter == gen.method:
                self.izm = gen.unit
                isType = gen.isType




        self.corp = str(element.LookupParameter('ФОП_Блок СМР').AsString())
        self.sec = str(element.LookupParameter('ФОП_Секция СМР').AsString())
        self.floor = str(element.LookupParameter('ФОП_Этаж').AsString())
        if parameter == 'ФОП_ВИС_Совместно с воздуховодом':
            pass

        #self.number = self.get_number(element, self.name)

        elemType = doc.GetElement(element.GetTypeId())
        if element in colInsul and elemType.LookupParameter('ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
            self.name = 'Изоляция ' + get_ADSK_Name(element) + ' для фланцев и стыков'

def is_object_to_generate(element, genCol, collection, parameter, genList = genList):
    if element in genCol:
        for gen in genList:
            if gen.collection == collection and parameter == gen.method:
                try:
                    elemType = doc.GetElement(element.GetTypeId())
                    if elemType.LookupParameter(parameter).AsInteger() == 1:
                        return True
                except Exception:
                    if element.LookupParameter(parameter).AsInteger() == 1:
                        return True


def collapse_list(lists):
    singles = []
    for gen in genList:
        name = gen.name
        isSingle = gen.isType
        if isSingle:
            singles.append(name)


    dict = {}

    for list in lists:
        system = list[3]
        corp = list[0]
        sec = list[1]
        floor = list[2]
        EF = list[13]
        name = list[5]
        number = list[10]

        if name in singles:
            number = 1
        Key = str(corp) + "_" + str(sec) + "_" + str(floor) + "_" + str(EF) + "_" + str(system) + "_" + str(name)

        if Key not in dict:
            dict[Key] = list
        else:
            dict[Key][10] = dict[Key][10] + number
            if name in singles:
                dict[Key][10] = 1


    collapsed_list = []
    for x in dict:
        collapsed_list.append(dict[x])
    return collapsed_list


def script_execute():
    with revit.Transaction("Добавление расчетных элементов"):
        # при каждом повторе расчета удаляем старые версии
        remove_models(colModel, '_Якорный элемен(металл и краска)')

        #список элементов которые будут сгенерированы
        calculation_elements = []

        collections = [colInsul, colPipes, colCurves]

        #тут мы перебираем элементы из коллекций по дурацкому алгоритму соответствия списку параметризации
        for collection in collections:
            for element in collection:
                elemType = doc.GetElement(element.GetTypeId())
                for gen in genList:
                    name = gen.name
                    mark = gen.mark
                    maker = gen.maker

                    parameter = gen.method
                    genCol = gen.collection

                    if is_object_to_generate(element, genCol, collection, parameter):

                        definition = calculation_element(element, collection, parameter, name, mark, maker)
                        definitionList = [definition.corp, definition.sec, definition.floor, definition.system,
                                          definition.group, definition.name, definition.mark, definition.art,
                                          definition.maker, definition.izm, definition.number, definition.mass,
                                          definition.comment, definition.EF]

                        calculation_elements.append(definitionList)


        calculation_elements = collapse_list(calculation_elements)

        newPos = []
        for element in calculation_elements:
            newPos.append(rowOfSpecification(corp=element[0],
                                     sec=element[1],
                                     floor=element[2],
                                     system=element[3],
                                     group=element[4],
                                     name=element[5],
                                     mark=element[6],
                                     art= element[7],
                                     maker=element[8],
                                     unit=element[9],
                                     number=element[10],
                                     mass=element[11],
                                     comment=element[12],
                                     EF=element[13]))


        new_position(newPos, temporary, '_Якорный элемен(металл и краска)')

temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, '_Якорный элемен(металл и краска)')

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

if temporary == None:
    print 'Не обнаружен якорный элемен(металл и краска). Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()



status = paraSpec.check_parameters()
if not status:
    script_execute()

