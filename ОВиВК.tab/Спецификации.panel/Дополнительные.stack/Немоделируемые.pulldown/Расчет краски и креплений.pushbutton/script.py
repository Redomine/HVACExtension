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
    diameter = 0
    width = 0
    height = 0

    length, area = material_calculator.get_curve_len_area_parameters(element)

    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        diameter = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER),
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
    if operation_name in ["Хомут трубный под шпильку М8", "Шпилька М8 1м/1шт"] and element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return material_calculator.get_collars_and_pins(element, diameter, length)

    return 0

def remove_old_models(doc, material_description, consumable_description):
    unmodeling_factory.remove_models(doc, material_description)
    unmodeling_factory.remove_models(doc, consumable_description)

def generate_materials(doc, family_symbol, material_description):
    def process_pipe_clamps(doc, elements, system, function, rule_set, material_description, family_symbol):
        pipes = []
        pipe_dict = {}

        for element in elements:
            if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
                pipes.append(element)

        for pipe in pipes:
            full_diameter = UnitUtils.ConvertFromInternalUnits(
                pipe.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER),
                UnitTypeId.Millimeters)
            pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
            dependent_elements = pipe.GetDependentElements(pipe_insulation_filter)

            if len(dependent_elements) > 0:
                insulation = doc.GetElement(dependent_elements[0])
                insulation_thikness = UnitUtils.ConvertFromInternalUnits(
                    insulation.Thickness,
                    UnitTypeId.Millimeters)
                full_diameter += insulation_thikness

            if full_diameter not in pipe_dict:
                pipe_dict[full_diameter] = []
            pipe_dict[full_diameter].append(pipe)

        material_location = XYZ(0, 0, 0)
        for pipe_row in pipe_dict:
            new_row = create_row_of_specification(system, function, rule_set, material_description)
            new_row.name = new_row.name + " D=" + str(pipe_row)

            material_location = XYZ(material_location.X + 10, material_location.Y + 10, 0)

            for element in pipe_dict[pipe_row]:
                new_row.number += get_number(element, rule_set.name)

            unmodeling_factory.create_new_position(doc, new_row, family_symbol, material_description, material_location)

    def process_other_rules(doc, elements, system, function, rule_set, material_description,
                            material_location, family_symbol):
        new_row = create_row_of_specification(system, function, rule_set, material_description)
        for element in elements:
            new_row.number += get_number(element, rule_set.name)

        unmodeling_factory.create_new_position(doc, new_row, family_symbol, material_description, material_location)

    material_location = XYZ(0, 0, 0)
    generation_rules_list = unmodeling_factory.get_generation_element_list()

    for rule_set in generation_rules_list:
        elem_types = get_elements_types_by_category(rule_set.category)
        calculation_elements = get_calculation_elements(elem_types, rule_set.method_name, rule_set.category)

        split_lists = split_calculation_elements_list(calculation_elements)

        for elements in split_lists:
            system, function = unmodeling_factory.get_system_function(elements[0])

            if rule_set.name == "Хомут трубный под шпильку М8":
                process_pipe_clamps(doc, elements, system, function, rule_set, material_description, family_symbol)
            else:
                material_location = XYZ(material_location.X + 10, 0, 0)
                process_other_rules(doc, elements, system, function, rule_set, material_description,
                                    material_location, family_symbol)

def create_row_of_specification(system, function, rule_set, material_description):
    return RowOfSpecification(
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

def generate_insulation_consumables(doc, family_symbol, consumable_description):
    consumable_location = XYZ(0, 0, 0)
    insulations = get_insulation_elements(doc)
    split_insulation_lists = split_calculation_elements_list(insulations)

    # Кэширование результатов get_system_function и get_insulation_consumables
    system_function_cache = {}
    consumables_cache = {}
    curve_params_cache = {}

    for insulation_elements in split_insulation_lists:
        element_type = insulation_elements[0].GetElementType()

        if element_type not in consumables_cache:
            consumables_cache[element_type] = material_calculator.get_insulation_consumables(element_type)

        consumables = consumables_cache[element_type]

        for consumable in consumables:
            if element_type not in system_function_cache:
                system_function_cache[element_type] = unmodeling_factory.get_system_function(insulation_elements[0])

            system, function = system_function_cache[element_type]

            new_consumable_row = create_consumable_row_of_specification(system, function,
                                                                        consumable, consumable_description)

            host_elements = {}
            for insulation_element in insulation_elements:
                host_id = insulation_element.HostElementId
                if host_id is not None:
                    if host_id not in host_elements:
                        host = doc.GetElement(host_id)
                        host_elements[host_id] = host
                        if host.Category.IsId(BuiltInCategory.OST_DuctCurves) or host.Category.IsId(BuiltInCategory.OST_PipeCurves):
                            if host_id not in curve_params_cache:
                                curve_params_cache[host_id] = material_calculator.get_curve_len_area_parameters(host)
                            length, area = curve_params_cache[host_id]
                            if consumable.is_expenditure_by_linear_meter == 0:
                                new_consumable_row.number += consumable.expenditure * area
                            else:
                                new_consumable_row.number += consumable.expenditure * length

            consumable_location = XYZ(consumable_location.X - 10, 0, 0)

            unmodeling_factory.create_new_position(doc, new_consumable_row, family_symbol,
                                                   consumable_description, consumable_location)

def create_consumable_row_of_specification(system, function, consumable, consumable_description):
    return RowOfSpecification(
        system,
        function,
        '12. Расходники изоляции',
        consumable.name,
        consumable.mark,
        '',  # У расходников не будет кода изделия
        consumable.maker,
        consumable.unit,
        consumable_description
    )

def get_insulation_elements(doc):
    insulations = []
    insulations += unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_PipeInsulations)
    insulations += unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_DuctInsulations)
    return insulations

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    family_symbol = unmodeling_factory.startup_checks(doc)

    with revit.Transaction("Добавление расчетных элементов"):
        material_description = "Расчет краски и креплений"
        consumable_description = 'Расходники изоляции'

        remove_old_models(doc, material_description, consumable_description)

        generate_materials(doc, family_symbol, material_description)
        generate_insulation_consumables(doc, family_symbol, consumable_description)

script_execute()