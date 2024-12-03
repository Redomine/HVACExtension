#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = "Расчет краски и креплений"
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"

import clr

from UnmodelingClassLibrary import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep


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
material_calculator = MaterialCalculator()

def get_elements_types_by_category(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsElementType() \
        .ToElements()
    return col

def get_calculation_elements(element_types, calculation_name, builtin_category):
    result_list = []

    for element_type in element_types:
        if element_type.GetSharedParamValueOrDefault(calculation_name) == 1:
            for el_id in element_type.GetDependentElements(None):
                element = doc.GetElement(el_id)
                category = element.Category
                if category and category.IsId(builtin_category) and element.GetTypeId() != ElementId.InvalidElementId:
                    result_list.append(element)

    return result_list

def split_calculation_elements_list(elements):
    # Создаем словарь для группировки элементов по ключу
    grouped_elements = defaultdict(list)

    for element in elements:
        shared_function = element.GetSharedParamValueOrDefault(
            SharedParamsConfig.Instance.EconomicFunction.Name, "Нет значения")
        shared_system = element.GetSharedParamValueOrDefault(
            SharedParamsConfig.Instance.VISSystemName.Name, "Нет значения")
        function_system_key = shared_function + "_" + shared_system

        # Добавляем элемент в соответствующий список в словаре
        grouped_elements[function_system_key].append(element)

    # Преобразуем значения словаря в список списков
    lists = list(grouped_elements.values())

    return lists

def get_number(element, operation_name, unmodeling_factory):
    length = 0
    diameter = 0
    width = 0
    height = 0



    length = UnitUtils.ConvertFromInternalUnits(
        element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH),
        UnitTypeId.Meters)

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

    if operation_name == "Металлические крепления для трубопроводов" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return material_calculator.get_pipe_material(length, diameter)
    if operation_name == "Металлические крепления для воздуховодов" and element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        return material_calculator.get_duct_material(element, diameter, width, height, area)
    if operation_name == "Краска антикоррозионная за два раза" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return material_calculator.get_color(element)
    if operation_name == "Грунтовка для стальных труб" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return material_calculator.get_grunt(area)
    if operation_name == "Хомут трубный под шпильку М8" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return material_calculator.get_collars_and_pins(element, diameter, length)
    if operation_name == "Шпилька М8 1м/1шт" and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return material_calculator.get_collars_and_pins(element, diameter, length)

    return 0

description = "Расчет краски и креплений"

# def check_family():
#     family_names = ["ОбщМд_Отв_Отверстие_Прямоугольное_В стене", "ОбщМд_Отв_Отверстие_Круглое_В стене"]
#
#     param_list = [
#         shared_currency_level_offset_name,
#         shared_currency_from_level_offset_name,
#         shared_currency_absolute_offset_name,
#         shared_level_offset_name,
#         shared_from_level_offset_name,
#         shared_absolute_offset_name
#         ]
#
#     for family_name in family_names:
#         family = find_family_symbol(family_name).Family
#         symbol_params = get_family_shared_parameter_names(family)
#
#         for param in param_list:
#             if param not in symbol_params:
#                 forms.alert("Параметра {} нет в семействах отверстий. Обновите все семейства отверстий из базы семейств.".
#                             format(param), "Ошибка", exitscript=True)

# def get_family_shared_parameter_names(family):
#     # Открываем документ семейства для редактирования
#     family_doc = doc.EditFamily(family)
#
#     shared_parameters = []
#     try:
#         # Получаем менеджер семейства
#         family_manager = family_doc.FamilyManager
#
#         # Получаем все параметры семейства
#         parameters = family_manager.GetParameters()
#
#         # Фильтруем параметры, чтобы оставить только общие
#         shared_parameters = [param.Definition.Name for param in parameters if param.IsShared]
#
#         return shared_parameters
#     finally:
#         # Закрываем документ семейства без сохранения изменений
#         family_doc.Close(False)


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with revit.Transaction("Добавление расчетных элементов"):
        family_name = "_Якорный элемент"

        if doc.IsFamilyDocument:
            forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

        unmodeling_factory = UnmodelingFactory()

        family_symbol = unmodeling_factory.is_family_in(doc, family_name)

        if family_symbol is None:
            forms.alert(
                    "Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.",
                    "Ошибка",
                    exitscript=True)


        generation_rules_list = unmodeling_factory.get_generation_element_list()
        # при каждом повторе расчета удаляем старые версии
        unmodeling_factory.remove_models(doc, description)

        loc = XYZ(0, 0, 0)

        # генерация металла, креплений и краски
        # Для каждого рулсета расчета создаем список сгруппированных по функции-имени системы элементов у которых этот расчет активен
        for rule_set in generation_rules_list:
            elem_types = get_elements_types_by_category(rule_set.category)
            calculation_elements = get_calculation_elements(elem_types, rule_set.method_name, rule_set.category)

            # поделенные по экономической функции и имени системы листы
            split_lists = split_calculation_elements_list(calculation_elements)

            # Проходимся по разделенным спискам элементов и для каждого из них создаем новый якорный элемент
            for elements in split_lists:
                new_row = RowOfSpecification()
                new_row.name = rule_set.name
                new_row.mark = rule_set.mark
                new_row.code = rule_set.code
                new_row.unit = rule_set.unit
                new_row.maker = rule_set.maker

                new_row.group = rule_set.group
                new_row.local_description = description
                # Эти элементы сгруппированы по функции-системы, достаточно забрать у одного
                new_row.system = elements[0].GetParamValueOrDefault(SharedParamsConfig.Instance.VISSystemName, "")
                new_row.function = elements[0].GetParamValueOrDefault(SharedParamsConfig.Instance.EconomicFunction, "")

                for element in elements:
                    new_row.number += get_number(element, rule_set.name, unmodeling_factory)

                # Увеличение координаты X на 1, чтоб элементы не создавались в одном месте
                loc = XYZ(loc.X + 10, loc.Y, loc.Z)

                unmodeling_factory.create_new_position(doc, new_row, family_symbol, family_name, description, loc)

        #генерация расходников изоляции


        insulations = []
        insulations+=(unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_PipeInsulations))
        insulations+=(unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_DuctInsulations))

        split_insulation_lists = split_calculation_elements_list(insulations)
        consumable_description = 'Расходники изоляцияя'

        for insulation_elements in split_insulation_lists:
            consumables = material_calculator.get_insulation_consumables(insulation_elements[0].GetElementType())

            print consumables

            for consumable in consumables:
                loc = XYZ(loc.X - 10, loc.Y, loc.Z)

                new_consumable_row = RowOfSpecification()
                new_consumable_row.name = consumable.name
                new_consumable_row.mark = consumable.mark
                new_consumable_row.unit = consumable.unit
                new_consumable_row.maker = consumable.maker

                for insulation_element in insulation_elements:
                    new_consumable_row.number += 1

                new_consumable_row.group = '12. Расходники изоляции'
                new_consumable_row.system = insulation_elements[0].GetParamValueOrDefault(SharedParamsConfig.Instance.VISSystemName, '')
                new_consumable_row.function = insulation_elements[0].GetParamValueOrDefault(SharedParamsConfig.Instance.EconomicFunction, '')


                unmodeling_factory.create_new_position(doc, new_consumable_row,
                                                       family_symbol, family_name, consumable_description, loc)

script_execute()

