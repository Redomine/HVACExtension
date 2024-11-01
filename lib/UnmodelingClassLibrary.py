# -*- coding: utf-8 -*-

import clr

from checkAnchor import famNames

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from dosymep.Bim4Everyone.SharedParams import *
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep_libs.bim4everyone import *
from dosymep.Revit import *

class GenerationElement:
    def __init__(self, group, name, mark, art, maker, unit, method_name, category, isType):
        self.group = group
        self.name = name
        self.mark = mark
        self.maker = maker
        self.unit = unit
        self.category = category
        self.method_name = method_name
        self.isType = isType
        self.art = art

class MaterialVariants:
    def __init__(self, diameter, insulated_rate, not_insulated_rate):
        self.diameter = diameter
        self.insulated_rate = insulated_rate
        self.not_insulated_rate = not_insulated_rate

#класс содержащий все ячейки типовой спецификации
class RowOfSpecification:
    def __init__(self):
        self.system = ""
        self.group = ""
        self.name = ""
        self.mark = ""
        self.code = ""
        self.maker = ""
        self.unit = ""
        self.number = 0
        self.mass = ""
        self.note = ""
        self.function = ""

        self.local_description = ""
        self.diameter = 0
        self.parentId = 0

def get_generation_element_list():
    gen_list = [
        GenerationElement(
            group="12. Расчетные элементы",
            name="Металлические крепления для воздуховодов",
            mark="",
            art="",
            unit="кг.",
            maker="",
            method_name=SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
            category=BuiltInCategory.OST_DuctCurves,
            isType=False),
        GenerationElement(
            group="12. Расчетные элементы",
            name="Металлические крепления для трубопроводов",
            mark="",
            art="",
            unit="кг.",
            maker="",
            method_name=SharedParamsConfig.Instance.VISIsFasteningMetalCalculation.Name,
            category=BuiltInCategory.OST_PipeCurves,
            isType=False),

        GenerationElement(
            group="12. Расчетные элементы",
            name="Краска антикоррозионная за два раза",
            mark="БТ-177",
            art="",
            unit="кг.",
            maker="",
            method_name=SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
            category=BuiltInCategory.OST_PipeCurves,
            isType=False),
        GenerationElement(
            group="12. Расчетные элементы",
            name="Грунтовка для стальных труб",
            mark="ГФ-031",
            art="",
            unit="кг.",
            maker="",
            method_name=SharedParamsConfig.Instance.VISIsPaintCalculation.Name,
            category=BuiltInCategory.OST_PipeCurves,
            isType=False),
        GenerationElement(

            group="12. Расчетные элементы",
            name="Хомут трубный под шпильку М8",
            mark="",
            art="",
            unit="шт.",
            maker="",
            method_name=SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
            category=BuiltInCategory.OST_PipeCurves,
            isType=False),
        GenerationElement(
            group="12. Расчетные элементы",
            name="Шпилька М8 1м/1шт",
            mark="",
            art="",
            unit="шт.",
            maker="",
            method_name=SharedParamsConfig.Instance.VISIsClampsCalculation.Name,
            category=BuiltInCategory.OST_PipeCurves,
            isType=False)
    ]
    return gen_list

def is_element_edited_by(element):
    user_name = __revit__.Application.Username
    edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
    if edited_by is None:
        return None

    if edited_by.lower() != user_name.lower():
        return edited_by
    return None

def get_pipe_material_variants():
    """ Возвращает коллекцию вариантов расхода металла по диаметрам для изолированных труб. Для неизолированных 0 """

    # Ключ словаря - диаметр. Первое значение списка по ключу - значение для изолированной трубы, второе - для неизолированной
    # для труб согласовано использование в расчетах только изолированных трубопроводов, поэтому в качестве неизолированной при создании
    # экземпляра варианта используем 0
    dict_var_p_mat = {15: 0.14, 20: 0.12, 25: 0.11, 32: 0.1, 40: 0.11, 50: 0.144, 65: 0.195,
                      80: 0.233, 100: 0.37, 125: 0.53}

    variants = []
    for diameter, insulated_rate in dict_var_p_mat.items():
        variant = MaterialVariants(diameter, insulated_rate, 0)
        variants.append(variant)

    return variants

def get_collar_material_variants():
    """ Возвращает коллекцию вариантов расхода хомутов по диаметрам для изолированных и неизолированных труб """

    # Ключ словаря - диаметр. Первое значение списка по ключу - значение для изолированной трубы, второе - для неизолированной
    dict_var_collars = {15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
                        80: [6, 4], 100: [6, 4.5], 125: [7, 5]}

    variants = []

    for diameter, insulated_rate, not_insulated_rate in dict_var_collars.items():
        variant = MaterialVariants(diameter, insulated_rate, not_insulated_rate)
        variants.append(variant)

    return variants

# Генерирует пустые элементы в рабочем наборе немоделируемых
def create_new_position(doc, new_row_data, family_symbol, family_name, description, loc):
    family_symbol.Activate()
    col_model = []

    # Находим рабочий набор "99_Немоделируемые элементы"
    fws = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    workset_id = None
    for ws in fws:
        if ws.Name == '99_Немоделируемые элементы':
            workset_id = ws.Id
            break

    if workset_id is None:
        print('Не удалось найти рабочий набор "99_Немоделируемые элементы", проверьте список наборов')
        return

    # Создаем элемент и назначем рабочий набор
    family_inst = doc.Create.NewFamilyInstance(loc, family_symbol, Structure.StructuralType.NonStructural)
    family_inst_workset = family_inst.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    family_inst_workset.Set(workset_id.IntegerValue)

    group = '{}_{}_{}_{}_{}'.format(
        new_row_data.group, new_row_data.name, new_row_data.mark, new_row_data.maker, new_row_data.code)

    family_inst.SetParamValue(SharedParamsConfig.Instance.VISSystemName, new_row_data.system)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISGrouping, group)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISCombinedName, new_row_data.name)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISMarkNumber, new_row_data.mark)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISItemCode, new_row_data.code)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISManufacturer, new_row_data.maker)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISUnit, new_row_data.unit)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISSpecNumbers, new_row_data.number)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISMass, new_row_data.mass)
    family_inst.SetParamValue(SharedParamsConfig.Instance.VISNote, new_row_data.note)
    family_inst.SetParamValue(SharedParamsConfig.Instance.EconomicFunction, new_row_data.function)


#для прогона новых ревизий генерации немоделируемых: стирает элемент с переданным именем модели
def remove_models(doc, description):
    fam_name = "_Якорный элемент"
    # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
    col_model = \
        [elem for elem in get_elements_by_category(doc, BuiltInCategory.OST_GenericModel) if elem.GetElementType()
        .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == fam_name]

    for element in col_model:
        edited_by = is_element_edited_by(element)
        if edited_by:
            forms.alert("Якорные элементы не были обработаны, так как были заняты пользователями:" + edited_by,
                        "Ошибка",
                        exitscript=True)

    for element in col_model:
        if element.LookupParameter('ФОП_ВИС_Назначение'):
            elem_type = doc.GetElement(element.GetTypeId())
            current_name = elem_type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString()
            current_description = element.LookupParameter('ФОП_ВИС_Назначение').AsString()

            if  current_name == fam_name:
                if description in current_description:
                    doc.Delete(element.Id)

# Возвращает FamilySymbol, если семейство есть в проекте, None если нет
def is_family_in(doc, name):
    # Создаем фильтрованный коллектор для категории OST_Mass и класса FamilySymbol
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(FamilySymbol)

    # Итерируемся по элементам коллектора
    for element in collector:
        if element.Family.Name == name:
            return element

    return None

def get_elements_by_category(doc, category):
    return FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()