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
import checkAnchor


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




#Исходные данные
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)
colSystems = make_col(BuiltInCategory.OST_DuctSystem)
colInsul = make_col(BuiltInCategory.OST_DuctInsulations)
nameOfModel = '_Якорный элемент'
description = 'Расчет краски и креплений'


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
    generationElement(group = '12. Расчетные элементы', name = "Грунтовка для стальных труб", mark = 'ГФ-031', art = '', unit = 'кг.', maker = '', method =  'ФОП_ВИС_Расчет краски и грунтовки', collection= colPipes,isType= False),
    generationElement(group = '12. Расчетные элементы', name = "Хомут трубный под шпильку М8", mark = '', art = '', unit = 'шт.', maker = '', method =  'ФОП_ВИС_Расчет хомутов', collection= colPipes,isType= False)
]

class calculation_element:
    def collars(self, element):
        lenght = fromRevitToMeters(element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble())
        D = fromRevitToMilimeters(element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble())
        D = int(D)

        self.name = self.name + ', Ду' + str(D)

        if lenght*1000 < D:
            return 1
        if lenght < 3000:
            return 2
        if lenght < 6000:
            return 3
        if lenght < 9000:
            return 4
        if lenght < 12000:
            return 5

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
        if name == "Хомут трубный под шпильку М8" and element in colPipes:
            Number = self.collars(element)


        return Number

    def __init__(self, element, collection, parameter, Name, Mark, Maker):
        self.corp = str(element.LookupParameter('ФОП_Блок СМР').AsString())
        self.sec = str(element.LookupParameter('ФОП_Секция СМР').AsString())
        self.floor = str(element.LookupParameter('ФОП_Этаж').AsString())

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
        self.unit = 'None'
        self.number = self.get_number(element, self.name)
        self.mass = ''
        self.comment = ''
        self.EF = str(element.LookupParameter('ФОП_Экономическая функция').AsString())
        self.parentId = element.Id.IntegerValue


        for gen in genList:
            if gen.collection == collection and parameter == gen.method:
                self.unit = gen.unit
                isType = gen.isType


        if parameter == 'ФОП_ВИС_Совместно с воздуховодом':
            pass

        #self.number = self.get_number(element, self.name)

        elemType = doc.GetElement(element.GetTypeId())
        if element in colInsul and elemType.LookupParameter('ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
            self.name = 'Изоляция ' + get_ADSK_Name(element) + ' для фланцев и стыков'

        self.key = self.corp + self.sec + self.floor + self.system + \
                   self.group + self.name + self.mark + self.art + \
                   self.maker

def is_object_to_generate(element, genCol, collection, parameter, genList = genList):
    if element in genCol:
        for gen in genList:
            if gen.collection == collection and parameter == gen.method:
                try:
                    elemType = doc.GetElement(element.GetTypeId())
                    if elemType.LookupParameter(parameter).AsInteger() == 1:
                        return True
                except Exception:
                    print parameter
                    if element.LookupParameter(parameter).AsInteger() == 1:
                        return True

def script_execute():
    with revit.Transaction("Добавление расчетных элементов"):
        # при каждом повторе расчета удаляем старые версии
        remove_models(colModel, nameOfModel, description)

        #список элементов которые будут сгенерированы
        calculation_elements = []

        collpasing_objects = []

        collections = [colInsul, colPipes, colCurves]

        elements_to_generate = []
        #перебираем элементы и выясняем какие из них подлежат генерации
        for collection in collections:
            for element in collection:
                elemType = doc.GetElement(element.GetTypeId())
                for gen in genList:
                    binding_name = gen.name
                    binding_mark = gen.mark
                    binding_maker = gen.maker

                    parameter = gen.method
                    genCol = gen.collection
                    if is_object_to_generate(element, genCol, collection, parameter):
                        definition = calculation_element(element, collection, parameter, binding_name, binding_mark, binding_maker)

                        key = definition.corp + definition.sec + definition.floor + definition.system + \
                                          definition.group + definition.name + definition.mark + definition.art + \
                                          definition.maker

                        toAppend = True
                        for element_to_generate in elements_to_generate:
                            if element_to_generate.key == key:
                                toAppend = False
                                element_to_generate.number = element_to_generate.number + definition.number

                        if toAppend:
                            elements_to_generate.append(definition)

        new_position(elements_to_generate, temporary, nameOfModel, description)

temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, nameOfModel)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

if temporary == None:
    print 'Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()



status = paraSpec.check_parameters()
if not status:
    anchor = checkAnchor.check_anchor(showText = False)
    if anchor:
        script_execute()

