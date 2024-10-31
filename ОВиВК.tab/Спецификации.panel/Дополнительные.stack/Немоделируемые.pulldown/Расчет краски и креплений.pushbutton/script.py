#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = "Расчет краски и креплений"
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"

import clr

from Redomine import rowOfSpecification
from paraSpec import shared_parameter

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

import checkAnchor
import math

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *

from collections import defaultdict
from UnmodelingClassLibrary import  *

from dosymep_libs.bim4everyone import *


#Исходные данные
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

def get_elements_by_category(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

def get_elements_types_by_category(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsElementType() \
        .ToElements()
    return col

def get_calculation_elements(types, calculation_name, builtin_category):
    result_list = []

    for type in types:
        if type.GetSharedParamValueOrDefault(calculation_name) == 1:
            for el_id in type.GetDependentElements(None):
                element = doc.GetElement(el_id)
                category = element.Category
                if category and category.IsId(builtin_category) and element.GetTypeId() != ElementId.InvalidElementId:
                    result_list.append(element)

    return result_list

def split_calculation_elements_list(elements):
    # Создаем словарь для группировки элементов по ключу
    grouped_elements = defaultdict(list)

    for element in elements:

        shared_function = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.EconomicFunction.Name, "Нет значения")
        shared_system = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.VISSystemName.Name, "Нет значения")
        function_system_key = shared_function + "_" + shared_system

        # Добавляем элемент в соответствующий список в словаре
        grouped_elements[function_system_key].append(element)

    # Преобразуем значения словаря в список списков
    lists = list(grouped_elements.values())

    print len(lists)
    # for x in lists:
    #     print x

    return lists

col_model = get_elements_by_category(BuiltInCategory.OST_GenericModel)
col_systems = get_elements_by_category(BuiltInCategory.OST_DuctSystem)


name_of_model = "_Якорный элемент"
description = "Расчет краски и креплений"

# VISIsPaintCalculation Расчет краски и грунтовки
# VISIsClampsCalculation ФОП_ВИС_Расчет хомутов
# VISIsFasteningMetalCalculation Расчет металла для креплений

# Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
col_model = \
    [elem for elem in col_model if elem.GetElementType()
    .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == name_of_model]

class CalculationElement:
    pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
    def __init__(self, element, collection, parameter, name, mark, maker):
        self.local_description = description


        self.length = UnitUtils.ConvertFromInternalUnits(element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH),
                                                         UnitTypeId.Meters)

        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            self.pipe_diameter = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM),
                UnitTypeId.Millimeters)

        if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and element.DuctType.Shape == ConnectorProfileType.Round:
            self.duct_diameter = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM),
                UnitTypeId.Millimeters)

        if element.IsExistsParam(SharedParamsConfig.Instance.VISSystemName.Name):
            self.system = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISSystemName)

        # Этот параметр не вызываем с платформы и удаляем из всех шаблонов
        if element.IsExistsParam("ADSK_Имя системы"):
            self.system = element.GetParamValueOrDefault("ADSK_Имя системы")

        self.name = name
        self.mark = mark
        self.art = ""
        self.maker = maker
        self.unit = "None"
        self.number = self.get_number(element, self.name)
        self.mass = ""
        self.comment = ""
        self.EF = element.GetParamValueOrDefault(SharedParamsConfig.Instance.EconomicFunction)
        self.parentId = element.Id.IntegerValue

        for gen in generation_rules_list:
            if gen.collection == collection and parameter == gen.method:
                self.unit = gen.unit
                isType = gen.isType



def get_number(element, name):
    length = 0
    diameter = 0
    width = 0
    height = 0
    area = UnitUtils.ConvertFromInternalUnits(
        element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA),
        UnitTypeId.SquareMeters)

    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        diameter = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM),
            UnitTypeId.Millimeters)

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and element.DuctType.Shape == ConnectorProfileType.Round:
        diameter = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM),
            UnitTypeId.Millimeters)

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and element.DuctType.Shape == ConnectorProfileType.Rectangular:
        width = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM),
            UnitTypeId.Millimeters)

        height = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM),
            UnitTypeId.Millimeters)


    def is_pipe_insulated(pipe):
        pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
        dependent_elements = pipe.GetDependentElements(pipe_insulation_filter)
        return len(dependent_elements) > 0

    def mid_calculation_fix(coef, curve_length):
        number = curve_length / coef
        if number < 1:
            number = 1
        return int(number)

    def get_pins(curve, pipe_length, pipe_diameter):
        dict_var_pins = {15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
                            80: [6, 4], 100: [6, 4.5], 125: [7, 5]}

        # Мы не считаем крепление труб до 0.5 м
        if pipe_length < 0.5:
            return 0

        if is_pipe_insulated(curve):
            if pipe_diameter in dict_var_pins:
                return mid_calculation_fix(dict_var_pins[pipe_diameter][0], pipe_length)
            else:
                return mid_calculation_fix(7, pipe_length)
        else:
            if pipe_diameter in dict_var_pins:
                return mid_calculation_fix(dict_var_pins[pipe_diameter][1], pipe_length)
            else:
                return mid_calculation_fix(5, pipe_length)

    def get_collars(pipe, pipe_diameter, pipe_length):
        # self.name = "{0}, Ду{1}".format(self.name, int(self.pipe_diameter))
        # self.local_description = "{0} {1}".format(self.local_description, self.name)
        dict_var_collars = {15:[2, 1.5], 20:[3, 2], 25:[3.5, 2], 32:[4, 2.5], 40:[4.5, 3], 50:[5, 3], 65:[6, 4],
                            80:[6, 4], 100:[6, 4.5], 125:[7, 5]}

        if pipe_length < 0.5:
            return 0

        if is_pipe_insulated(pipe):
            if pipe_diameter in dict_var_collars:
                return mid_calculation_fix(dict_var_collars[pipe_diameter][0], pipe_length)
            else:
                return mid_calculation_fix(7, pipe_length)
        else:
            if pipe_diameter in dict_var_collars:
                return mid_calculation_fix(dict_var_collars[pipe_diameter][1], pipe_length)
            else:
                return mid_calculation_fix(5, pipe_length)

    def get_duct_material(duct, duct_diameter, duct_width, duct_height, duct_area):
        perimeter = 0
        if duct.DuctType.Shape == ConnectorProfileType.Round:
            perimeter = 3.14 * duct_diameter

        if duct.DuctType.Shape == ConnectorProfileType.Rectangular:
            duct_width = UnitUtils.ConvertFromInternalUnits(
                duct.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            duct_height = UnitUtils.ConvertFromInternalUnits(
                duct.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            perimeter = 2 * (duct_width + duct_height)


        if perimeter < 1001:
            kg = duct_area * 0.65
        elif perimeter < 1801:
            kg = duct_area * 1.22
        else:
            kg = duct_area * 2.25

        return kg

    def get_pipe_material(pipe_length, pipe_diameter):
        dict_var_p_mat = {15: 0.14, 20: 0.12, 25: 0.11, 32: 0.1, 40: 0.11, 50: 0.144, 65: 0.195,
                            80: 0.233, 100: 0.37, 125: 0.53}
        coefficient = 1.7
        # Запас 70% задан по согласованию.

        # Сортируем ключи словаря
        sorted_keys = sorted(dict_var_p_mat.keys())

        # Ищем первый ключ, который больше или равен self.pipe_diameter
        for key in sorted_keys:
            if pipe_diameter <= key:
                key_up = dict_var_p_mat[key] * coefficient
                return key_up * pipe_length
        else:
            return 0.62*coefficient* pipe_length

    def get_grunt(pipe_area):
        number = pipe_area / 10
        return number

    def get_color(pipe):
        area = (pipe.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area * 0.2 * 2
        return number

    if name == "Металлические крепления для трубопроводов" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return get_pipe_material(length, diameter)
    if name == "Металлические крепления для воздуховодов" and element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        return get_duct_material(element, diameter, width, height)
    if name == "Краска антикоррозионная за два раза" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return get_color(element)
    if name == "Грунтовка для стальных труб" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return get_grunt(area)
    if name == "Хомут трубный под шпильку М8" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return get_collars(element, diameter, length)
    if name == "Шпилька М8 1м/1шт" and element in element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return get_pins(element, length, diameter)
    return 0


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with revit.Transaction("Добавление расчетных элементов"):
        generation_rules_list = get_generation_element_list()


        for rule_set in generation_rules_list:
            elem_types = get_elements_types_by_category(rule_set.category)
            elements = get_calculation_elements(elem_types, rule_set.method_name, rule_set.category)

            # поделенные по экономической функции и имени системы листы
            splited_lists = split_calculation_elements_list(elements)

            for elements in splited_lists:
                new_row = RowOfSpecification()

                new_row.local_description = description



        # при каждом повторе расчета удаляем старые версии
        #remove_models(col_model, name_of_model, description)


# if isItFamily():
#     forms.alert(
#         "Надстройка не предназначена для работы с семействами",
#         "Ошибка",
#         exitscript=True
#         )

# if temporary is None:
#     forms.alert(
#         "Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.",
#         "Ошибка",
#         exitscript=True
#         )

script_execute()

