#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = "Расчет краски и креплений"
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"

import clr

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

def get_calculation_elements(types, calculation_name):
    for type in types:
        if type.GetSharedParamValueOrDefault(calculation_name) == 1:
            list = []
            el_ids = type.GetSimilarTypes()
            for el_id in el_ids:
                list.append(doc.GetElement(el_id))




col_pipes = get_elements_by_category(BuiltInCategory.OST_PipeCurves)
col_curves = get_elements_by_category(BuiltInCategory.OST_DuctCurves)
col_model = get_elements_by_category(BuiltInCategory.OST_GenericModel)
col_systems = get_elements_by_category(BuiltInCategory.OST_DuctSystem)
col_insulation = get_elements_by_category(BuiltInCategory.OST_DuctInsulations)

name_of_model = "_Якорный элемент"
description = "Расчет краски и креплений"

# VISIsPaintCalculation Расчет краски и грунтовки
# VISIsClampsCalculation ФОП_ВИС_Расчет хомутов
# VISIsFasteningMetalCalculation Расчет металла для креплений

# Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
col_model = \
    [elem for elem in col_model if elem.GetElementType()
    .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == name_of_model]

generation_rules_list = get_generation_element_list()


class CalculationElement:
    pipe_insulation_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeInsulations)
    def __init__(self, element, collection, parameter, name, mark, maker):
        self.local_description = description

        self.corp = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksBlock)
        self.sec = element.GetParamValueOrDefault(SharedParamsConfig.Instance.BuildingWorksSection)
        self.floor = element.GetParamValueOrDefault(SharedParamsConfig.Instance.Level)

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

        self.group ="12. Расчетные элементы"
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

        parts = [part for part in [self.EF, self.corp, self.sec, self.floor, self.system,
                                   self.group, self.name, self.mark, self.art,
                                   self.maker, self.local_description] if part is not None]
        self.key = ''.join(parts)

    def is_pipe_insulated(self, element):
        dependent_elements = element.GetDependentElements(self.pipe_insulation_filter)
        return len(dependent_elements) > 0

    def mid_calculation_fix(self, coeff):
        number = self.length / coeff
        if number < 1:
            number = 1
        return int(number)

    def pins(self, element):
        self.local_description = "{0} {1}, Ду{2}".format(self.local_description, self.name, self.pipe_diameter)
        dict_var_pins = {15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
                            80: [6, 4], 100: [6, 4.5], 125: [7, 5]}

        # Мы не считаем крепление труб до 0.5 м
        if self.length < 0.5:
            return 0

        if self.is_pipe_insulated(element):
            if self.pipe_diameter in dict_var_pins:
                return self.mid_calculation_fix(dict_var_pins[self.pipe_diameter][0])
            else:
                return self.mid_calculation_fix(7)
        else:
            if self.pipe_diameter in dict_var_pins:
                return self.mid_calculation_fix(dict_var_pins[self.pipe_diameter][1])
            else:
                return self.mid_calculation_fix(5)

    def collars(self, element):
        self.name = "{0}, Ду{1}".format(self.name, int(self.pipe_diameter))
        self.local_description = "{0} {1}".format(self.local_description, self.name)
        dict_var_collars = {15:[2, 1.5], 20:[3, 2], 25:[3.5, 2], 32:[4, 2.5], 40:[4.5, 3], 50:[5, 3], 65:[6, 4],
                            80:[6, 4], 100:[6, 4.5], 125:[7, 5]}

        if self.length < 0.5:
            return 0

        if self.is_pipe_insulated(element):
            if self.pipe_diameter in dict_var_collars:
                return self.mid_calculation_fix(dict_var_collars[self.pipe_diameter][0])
            else:
                return self.mid_calculation_fix(7)
        else:
            if self.pipe_diameter in dict_var_collars:
                return self.mid_calculation_fix(dict_var_collars[self.pipe_diameter][1])
            else:
                return self.mid_calculation_fix(5)

    def duct_material(self, element):
        area = UnitUtils.ConvertFromInternalUnits(
            element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA),
            UnitTypeId.SquareMeters)  # Преобразование в квадратные метры

        if element.DuctType.Shape == ConnectorProfileType.Round:
            diameter = self.duct_diameter
            perimeter = 3.14 * diameter

        if element.DuctType.Shape == ConnectorProfileType.Rectangular:
            width = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            height = UnitUtils.ConvertFromInternalUnits(
                element.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM),
                UnitTypeId.Millimeters)  # Преобразование в метры

            perimeter = 2 * (width + height)


        if perimeter < 1001:
            kg = area * 0.65
        elif perimeter < 1801:
            kg = area * 1.22
        else:
            kg = area * 2.25

        return kg

    def pipe_material(self, element):
        dict_var_p_mat = {15: 0.14, 20: 0.12, 25: 0.11, 32: 0.1, 40: 0.11, 50: 0.144, 65: 0.195,
                            80: 0.233, 100: 0.37, 125: 0.53}
        up_coeff = 1.7
        # Запас 70% задан по согласованию.

        # Сортируем ключи словаря
        sorted_keys = sorted(dict_var_p_mat.keys())

        # Ищем первый ключ, который больше или равен self.pipe_diametr
        for key in sorted_keys:
            if self.pipe_diameter <= key:
                key_up = dict_var_p_mat[key] * up_coeff
                return key_up * self.length
        else:
            return 0.62*up_coeff*self.length

    def grunt(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area / 10
        return number

    def colorBT(self, element):
        area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
        number = area * 0.2 * 2
        return number

    def get_number(self, element, name):
        number = 1
        if name == "Металлические крепления для трубопроводов" and element in col_pipes:
            number = self.pipe_material(element)
        if name == "Металлические крепления для воздуховодов" and element in col_curves:
            number = self.duct_material(element)
        if name == "Краска антикоррозионная за два раза" and element in col_pipes:
            number = self.colorBT(element)
        if name == "Грунтовка для стальных труб" and element in col_pipes:
            number = self.grunt(element)
        if name == "Хомут трубный под шпильку М8" and element in col_pipes:
            number = self.collars(element)
        if name == "Шпилька М8 1м/1шт" and element in col_pipes:
            number = self.pins(element)
        return number


def is_object_to_generate(element, gen_col, collection, parameter):
    if element in gen_col:
        for gen in generation_rules_list:
            if gen.collection == collection and parameter == gen.method:
                elem_type = doc.GetElementType()
                if elem_type.GetParamValueOrDefault(parameter) == 1:
                    return True


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with revit.Transaction("Добавление расчетных элементов"):


        test = get_elements_types_by_category(BuiltInCategory.OST_DuctCurves)
        get_calculation_elements(test, SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name)
        # при каждом повторе расчета удаляем старые версии
        #remove_models(col_model, name_of_model, description)

        # collections = [col_insulation, col_pipes, col_curves]
        #
        # elements_to_generate = []
        #
        # #перебираем элементы и выясняем какие из них подлежат генерации
        # for collection in collections:
        #     for element in collection:
        #         for rule_set in generation_rules_list:
        #             set_name = rule_set.name
        #             set_mark = rule_set.mark
        #             set_maker = rule_set.maker
        #             set_parameter = rule_set.method
        #             set_collection = rule_set.collection
        #             if is_object_to_generate(element, set_collection, collection, set_parameter):
        #                 definition = CalculationElement(element, collection, set_parameter, set_name, set_mark, set_maker)
        #
        #                 parts = [part for part in [definition.EF, definition.corp, definition.sec, definition.floor,
        #                                            definition.system,definition.group, definition.name, definition.mark,
        #                                            definition.art, definition.maker, definition.local_description]
        #                          if part is not None]
        #                 key = ''.join(parts)
        #
        #
        #                 toAppend = True
        #                 for element_to_generate in elements_to_generate:
        #                     if element_to_generate.key == key:
        #                         toAppend = False
        #                         element_to_generate.number = element_to_generate.number + definition.number
        #
        #                 if toAppend:
        #                     elements_to_generate.append(definition)
        #
        # #иначе шпилек получится дробное число, а они в штуках
        # for el in elements_to_generate:
        #     if el.name == "Шпилька М8 1м/1шт":
        #         el.number = int(math.ceil(el.number))
        #
        # new_position(elements_to_generate, temporary, name_of_model, description)

#temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, name_of_model)

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

