#! /usr/bin/env python
# -*- coding: utf-8 -*-


import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import ctypes
import os
import csv
import codecs

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict

from unmodeling_class_library import *
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uiapp = __revit__.Application
view = doc.ActiveView
material_calculator = MaterialCalculator(doc)
unmodeling_factory = UnmodelingFactory()

class CSVRules:
    name_column = 0
    d_column = 0
    code_column = 0
    len_column = 0
    thickness_column = 0

class GenericPipe:
    def __init__(self, type_comment, name, dn, code, length, thickness):
        self.type_comment = type_comment
        self.name = name
        self.dn = dn
        self.code = code
        self.length = length
        self.thickness = thickness

class TypesCash:
    def __init__(self, dn, id, variants_pool):
        self.dn = dn
        self.id = id
        self.variants_pool = variants_pool

def create_folder_if_not_exist(project_path):
    """
    Создает папку, если она не существует.

    Args:
        project_path (str): Путь к папке проекта.
    """
    if not os.path.exists(project_path):
        os.makedirs(project_path)

def get_document_path():
    """
    Возвращает путь к документу.

    Returns:
        str: Путь к документу.
    """
    path = \
        "W:/Проектный институт/Отд.стандарт.BIM и RD/BIM-Ресурсы/5-Надстройки/Bim4Everyone/A101/MEP/AIEquipment/"

    if not (os.path.exists(path) and os.access(path, os.R_OK) and os.access(path, os.W_OK)):
        version = uiapp.VersionNumber
        documents_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        path = os.path.join(documents_path,
                            'dosymep',
                            str(version),
                            'AIEquipment/')

        create_folder_if_not_exist(path)

        report = ('Нет доступа к сетевому диску. Разместите таблицы выбора по пути: {} \n'
                  'Открыть папку?').format(path)

        # Вызов MessageBox из Windows API
        result = ctypes.windll.user32.MessageBoxW(0, report, "Внимание", 4)

        if result == 6:  # IDYES
            os.startfile(path)

    return path

def filter_elements_ai(elements):
    result = []
    for element in elements:
        if '_B4E_АИ' in element.Name:
            result.append(element)

    return result

def get_column_index(headers, name):
    if name in headers:
        return headers.index(name)
    else:
        forms.alert("Следующие заголовки не были найдены: " + name, "Ошибка", exitscript=True)

def get_float_value(value, column_number):
    try:
        return float(value)
    except:
        forms.alert("Ошибка при попытке получить числовое значение из столбца {}".format(column_number),
                    "Ошибка",
                    exitscript=True)

def get_pipe_variants(file_name):
    path = get_document_path()

    with codecs.open(path + file_name, 'r', encoding='utf-8-sig') as csvfile:
        material_variants = []
        # Создаем объект reader
        csvreader = csv.reader(csvfile, delimiter=";")
        headers = next(csvreader)

        rules = CSVRules()

        rules.type_comment_column = get_column_index(headers,'Комментарий к типоразмеру')
        rules.name_column = get_column_index(headers,'Наименование')
        rules.d_column = get_column_index(headers,'Диаметр')
        rules.code_column = get_column_index(headers,'Артикул')
        rules.len_column = get_column_index(headers,'Длина трубы')
        rules.thickness_column = get_column_index(headers,'Толщина трубы')

        # Итерируемся по строкам в файле
        for row in csvreader:
            material_variants.append(
                GenericPipe(
                    row[rules.type_comment_column],
                    row[rules.name_column],
                    get_float_value(row[rules.d_column], rules.d_column),
                    row[rules.code_column],
                    get_float_value(row[rules.len_column], rules.len_column),
                    row[rules.thickness_column]
                )
            )

    return material_variants

def convert_to_mms(value):
    result = UnitUtils.ConvertFromInternalUnits(value,
        UnitTypeId.Millimeters)
    return result

def get_variants_pool(element, variants, type_comment, dn):

    result = []

    is_pipe = element.Category.IsId(BuiltInCategory.OST_PipeCurves)
    #is_insulation = element.Category.IsId(BuiltInCategory.OST_PipeInsulation)

    # Проверяем, есть ли совпадение по комментарию типоразмера
    for variant in variants:
        if is_pipe:
            if type_comment == variant.type_comment and dn == variant.dn:
                result.append(variant)

    if len(result) == 0:
        forms.alert("Часть элементов в модели, помеченных как B4E_АИ, не обнаружена в согласованных каталогах, "
                    "что может привести к не полному формированию спецификации. "
                    "Устраните расхождения перед продолжением работы. \n"
                    "Пример элемента - ID:{}".format(element.Id),
                    "Ошибка",
                    exitscript=True)

    return result

def create_new_row(element, variant, number):
    shared_function = element.GetParamValueOrDefault(
        SharedParamsConfig.Instance.EconomicFunction, unmodeling_factory.out_of_function_value)
    shared_system = element.GetParamValueOrDefault(
        SharedParamsConfig.Instance.VISSystemName, unmodeling_factory.out_of_system_value)
    mark = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISMarkNumber, '')
    maker = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISManufacturer, '')
    unit = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISUnit, '')
    note = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISNote, '')
    group = '8. Трубопроводы'

    new_row = RowOfSpecification(
        shared_system,
        shared_function,
        group,
        name=variant.name,
        mark=mark,
        code=variant.code,
        maker=maker,
        unit=unit,
        local_description=unmodeling_factory.ai_description,
        number=number,
        mass='',
        note=note
    )

    return new_row

def separate_element(ai_pipe, variants_pool, family_symbol):
    ai_pipe_len = convert_to_mms(ai_pipe.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH))
    sorted_variants_pool = sorted(variants_pool, key=lambda x: x.length, reverse=True)

    result = []
    for variant in sorted_variants_pool:
        if ai_pipe_len <= 0:
            break

        number = ai_pipe_len // variant.length

        if number >= 1:
            ai_pipe_len = ai_pipe_len % variant.length

        len_not_minimal = ai_pipe_len > 50  # Принято что не считаем обрезки труб меньше 50мм
        last_variant = variant == sorted_variants_pool[-1]  # Проверка есть ли еще вариаты

        if last_variant and len_not_minimal:
            number += 1
            ai_pipe_len -= variant.length

        if number > 0:
            new_row = create_new_row(ai_pipe, variant, number)

            result.append(
                new_row
            )
    return result

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):

    pipe_variants = get_pipe_variants('/Трубопроводы АИ.csv')

    family_symbol = unmodeling_factory.startup_checks(doc)

    pipes = unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_PipeCurves)

    ai_pipes = filter_elements_ai(pipes)



    cash = [] # сохранение типоразмеров, чтоб не перебирать для каждой трубы каталог
    elements_to_generation = []
    for ai_pipe in ai_pipes:
        ai_pipe_type = ai_pipe.GetElementType()
        id = ai_pipe_type.Id

        type_comment = ai_pipe_type.GetParamValue(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)
        ai_pipe_dn = convert_to_mms(ai_pipe.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))

        # Проверяем наличие объекта в кэше
        cached_item = next((item for item in cash if item.dn == ai_pipe_dn and item.id == id), None)

        if cached_item:
            # Если объект найден в кэше, используем его данные
            variants_pool = cached_item.variants_pool
        else:
            # Если объект не найден в кэше, создаем новый объект и добавляем его в кэш
            variants_pool = get_variants_pool(ai_pipe, pipe_variants, type_comment, ai_pipe_dn)
            new_item = TypesCash(ai_pipe_dn, id, variants_pool)
            cash.append(new_item)

        # Если нет совпадений по комментарию типоразмера - продолжаем перебор
        # Если заявленная длина в каталоге 0 - элемент не бьется на части, можно пропускать
        if len(variants_pool) == 0 or variants_pool[0].length == 0:
            continue

        generic_elements = separate_element(ai_pipe, variants_pool, family_symbol)
        elements_to_generation.extend(generic_elements)

    with revit.Transaction("BIM: Добавление расчетных элементов"):
        # При каждом запуске затираем расходники с соответствующим описанием и генерируем заново
        unmodeling_factory.remove_models(doc, unmodeling_factory.ai_description)

        family_symbol.Activate()
        material_location = unmodeling_factory.get_base_location(doc)

        for element in elements_to_generation:
            material_location = unmodeling_factory.update_location(material_location)

            unmodeling_factory.create_new_position(doc, element, family_symbol,
                                                   unmodeling_factory.ai_description,
                                                   material_location)

script_execute()
