#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = "Расчет краски и креплений"
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"

import clr

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
unmodeling_factory = UnmodelingFactory()

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

def get_number(element, operation_name):
    length = 0
    diameter = 0
    width = 0
    height = 0

    length, area = material_calculator.get_curve_len_area_parameters(element)

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



@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    family_symbol = unmodeling_factory.startup_checks(doc)

    with revit.Transaction("Добавление расчетных элементов"):
        material_description = "Расчет краски и креплений"
        consumable_description = 'Расходники изоляции'

        # при каждом повторе расчета удаляем старые версии и для креплений и для расходников
        unmodeling_factory.remove_models(doc, material_description)
        unmodeling_factory.remove_models(doc, consumable_description)

        # базовые точки для размещения немоделируемых, для материалов растут по Х на 10 для каждого, для расходников - уменьшаются аналогично
        material_location = XYZ(0, 0, 0)
        consumable_location = XYZ(0, 0, 0)

        # генерация металла, креплений и краски
        # Для каждого рулсета расчета создаем список сгруппированных по функции-имени системы элементов у которых этот расчет активен
        generation_rules_list = unmodeling_factory.get_generation_element_list()

        for rule_set in generation_rules_list:
            elem_types = get_elements_types_by_category(rule_set.category)
            calculation_elements = get_calculation_elements(elem_types, rule_set.method_name, rule_set.category)

            # поделенные по экономической функции и имени системы листы
            split_lists = split_calculation_elements_list(calculation_elements)

            # Проходимся по разделенным спискам элементов и для каждого из них создаем новый якорный элемент
            for elements in split_lists:
                # Эти элементы сгруппированы по функции-системы, достаточно забрать у одного
                system, function = unmodeling_factory.get_system_function(elements[0])

                new_row = RowOfSpecification(
                    system,
                    function,
                    rule_set.group,
                    rule_set.name,
                    rule_set.mark,
                    rule_set.code,
                    rule_set.maker,
                    rule_set.unit,
                    material_description
                )

                for element in elements:
                    new_row.number += get_number(element, rule_set.name)

                # Увеличение координаты X на 10, чтоб элементы не создавались в одном месте
                material_location = XYZ(material_location.X + 10, 0, 0)

                unmodeling_factory.create_new_position(doc, new_row, family_symbol, material_description, material_location)

        #генерация расходников изоляции
        insulations = []
        insulations+=(unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_PipeInsulations))
        insulations+=(unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_DuctInsulations))

        split_insulation_lists = split_calculation_elements_list(insulations)

        for insulation_elements in split_insulation_lists:
            # получаем список из расходников для данного типа изоляции в отдельной функции-системе
            consumables = material_calculator.get_insulation_consumables(insulation_elements[0].GetElementType())

            # для каждого расходника генерируем новую строку и вычисляем его количество
            for consumable in consumables:
                # Эти элементы сгруппированы по функции-системы, достаточно забрать у одного
                system, function = unmodeling_factory.get_system_function(
                    insulation_elements[0])

                new_consumable_row = RowOfSpecification(
                    system,
                    function,
                    '12. Расходники изоляции',
                    consumable.name,
                    consumable.mark,
                    '', # У расходников не будет кода изделия
                    consumable.maker,
                    consumable.unit,
                    consumable_description
                )

                for insulation_element in insulation_elements:
                    host_id = insulation_element.HostElementId
                    if host_id is not None:
                        length, area = material_calculator.get_curve_len_area_parameters(doc.GetElement(host_id))
                        if consumable.is_expenditure_by_linear_meter == 0:
                            new_consumable_row.number += consumable.expenditure * area
                        else:
                            new_consumable_row.number += consumable.expenditure * length

                consumable_location = XYZ(consumable_location.X - 10, 0, 0)

                unmodeling_factory.create_new_position(doc,
                                                       new_consumable_row,
                                                       family_symbol,
                                                       consumable_description,
                                                       consumable_location)

script_execute()

