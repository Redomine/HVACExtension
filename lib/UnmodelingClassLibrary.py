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
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep_libs.bim4everyone import *

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
        self.system = None
        self.group = "12. Расчетные элементы"
        self.name = None
        self.mark = None
        self.code = None
        self.maker = None
        self.unit = None
        self.number = 0
        self.mass = None
        self.note = None
        self.function = None

        self.local_description = None
        self.diameter = None
        self.parentId = None

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
    # для труб согласовано использование в расчетах только изолированных трубопроводов
    dict_var_p_mat = {15: 0.14, 20: 0.12, 25: 0.11, 32: 0.1, 40: 0.11, 50: 0.144, 65: 0.195,
                      80: 0.233, 100: 0.37, 125: 0.53}

    variants = []
    for diameter, insulated_rate in dict_var_p_mat.items():
        variant = MaterialVariants(diameter, insulated_rate, 0)
        variants.append(variant)

    return variants

def get_collar_material_variants():
    """ Возвращает коллекцию вариантов расхода хомутов по диаметрам для изолированных и неизолированных труб """
    dict_var_collars = {15: [2, 1.5], 20: [3, 2], 25: [3.5, 2], 32: [4, 2.5], 40: [4.5, 3], 50: [5, 3], 65: [6, 4],
                        80: [6, 4], 100: [6, 4.5], 125: [7, 5]}

    variants = []

    for diameter, insulated_rate, not_insulated_rate in dict_var_collars.items():
        variant = MaterialVariants(diameter, insulated_rate, not_insulated_rate)
        variants.append(variant)

    return variants

#заполняет ячейки в сгенерированном немоделируемом
def set_element(element, name, setting):

    if setting == 'None':
        setting = ''
    if setting == None:
        setting = ''

    if name == 'ФОП_Экономическая функция':
        if setting == '':
            setting = 'None'
        if setting == None:
            setting = 'None'

    if name == 'ФОП_ВИС_Имя системы':
        if setting == '':
            setting = 'None'
        if setting == None:
            setting = 'None'


    if name == 'ФОП_ВИС_Число' or name == 'ФОП_ВИС_Масса':
        element.LookupParameter(name).Set(setting)
    else:
        try:
            element.LookupParameter(name).Set(str(setting))
        except:
            element.LookupParameter(name).Set(setting)

# Генерирует пустые элементы в рабочем наборе немоделируемых
def create_new_position(doc, calculation_elements, temporary, famName, description):
    # Создаем заглушки по элементам, собранным из таблицы
    loc = XYZ(0, 0, 0)

    temporary.Activate()
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

    # Создаем элементы и добавляем их в colModel
    for _ in calculation_elements:
        family_inst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)
        col_model.append(family_inst)

    # Фильтруем элементы и присваиваем рабочий набор
    for element in col_model:
        try:
            elem_type = doc.GetElement(element.GetTypeId())
            if elem_type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString() == famName:
                if not element.LookupParameter('ФОП_ВИС_Назначение').AsString():
                    ews = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                    ews.Set(workset_id.IntegerValue)
        except Exception as e:
            print 'Ошибка при присвоении рабочего набора'

    index = 1
    # Для первого элемента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for position in calculation_elements:
        group = position.group
        if description != 'Пустая строка':
            pos_group = '{}_{}_{}_{}'.format(position.group, position.name, position.mark, index)
            index += 1
            if description in ['Расходники изоляции', 'Расчет краски и креплений']:
                pos_group = '{}_{}_{}'.format(position.group, position.name, position.mark)
            group = pos_group

        dummy = col_model.pop(0) if description != 'Пустая строка' else family_inst

        set_element(dummy, 'ФОП_Блок СМР', position.corp)
        set_element(dummy, 'ФОП_Секция СМР', position.sec)
        set_element(dummy, 'ФОП_Этаж', position.floor)
        set_element(dummy, 'ФОП_ВИС_Имя системы', position.system)
        set_element(dummy, 'ФОП_ВИС_Группирование', group)
        set_element(dummy, 'ФОП_ВИС_Наименование комбинированное', position.name)
        set_element(dummy, 'ФОП_ВИС_Марка', position.mark)
        set_element(dummy, 'ФОП_ВИС_Код изделия', position.art)
        set_element(dummy, 'ФОП_ВИС_Завод-изготовитель', position.maker)
        set_element(dummy, 'ФОП_ВИС_Единица измерения', position.unit)
        set_element(dummy, 'ФОП_ВИС_Число', position.number)
        set_element(dummy, 'ФОП_ВИС_Масса', position.mass)
        set_element(dummy, 'ФОП_ВИС_Примечание', position.comment)
        set_element(dummy, 'ФОП_Экономическая функция', position.EF)

        # Фильтрация шпилек под разные диаметры
        set_element(dummy, 'ФОП_ВИС_Назначение', description if description != 'Расчет краски и креплений' else position.local_description)

        if description != 'Пустая строка':
            col_model.pop(0)

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
def is_family_in(doc, builtin, name):
    # Создаем фильтрованный коллектор для категории OST_Mass и класса FamilySymbol
    collector = FilteredElementCollector(doc).OfCategory(builtin).OfClass(FamilySymbol)

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