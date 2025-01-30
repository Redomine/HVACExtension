#! /usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import os
import csv
import codecs

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *

from itertools import chain
from System import Environment
from Autodesk.Revit.UI import TaskDialog
from unmodeling_class_library import *

doc = __revit__.ActiveUIDocument.Document
uiapp = __revit__.Application
view = doc.ActiveView
material_calculator = MaterialCalculator(doc)
unmodeling_factory = UnmodelingFactory()

class CSVRules:
    """Класс-правило для нумерации столбцов в CSV"""
    name_column = 0
    d_column = 0
    code_column = 0
    mark_column = 0
    maker_column = 0
    len_column = 0

class AICatalogElement:
    def __init__(self, type_comment, name, dn, code, length, mark, maker):
        self.type_comment = type_comment
        self.name = name
        self.dn = dn
        self.code = code
        self.length = length
        self.mark = mark
        self.maker = maker

class UpdateElement:
    def __init__(self, element, data):
        self.element = element
        self.data = data

class TypesCash:
    def __init__(self, dn, id, variants_pool):
        self.dn = dn
        self.id = id
        self.variants_pool = variants_pool

def get_document_path():
    """
    Возвращает путь к документу.

    Returns:
        str: Путь к документу.
    """
    # Исходный путь к файлу
    network_path = "W:/Проектный институт/Отд.стандарт.BIM и RD/BIM-Ресурсы/5-Надстройки/Bim4Everyone/A101/MEP/AIEquipment/"
    file_name = 'Элементы АИ.csv'
    full_network_path = os.path.join(network_path, file_name)

    # Проверка доступности сетевого пути
    if os.path.exists(network_path) and os.access(network_path, os.R_OK | os.W_OK):
        return full_network_path

    # Если сетевой путь недоступен, используем локальный путь
    version = uiapp.VersionNumber
    documents_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    local_path = os.path.join(documents_path, 'dosymep', str(version), 'AIEquipment')
    full_local_path = os.path.join(local_path, file_name)

    # Создаем локальную директорию, если она не существует
    if not os.path.exists(local_path):
        os.makedirs(local_path)

    # Если файл отсутствует, уведомляем пользователя
    if not os.path.exists(full_local_path):
        report = ('Нет доступа к сетевому диску. Разместите таблицы выбора по пути: {} \n'
                  'Открыть папку?').format(local_path)

        if show_dialog(report):
            os.startfile(local_path)

        sys.exit()

    return full_local_path

def filter_elements_ai(elements):
    """Оставляет от исходного списка только те элементы, у которых есть _B4E_АИ в назвнии типа

    Args:
        elements: Список Element из модели
    """
    result = []
    for element in elements:
        if '_B4E_АИ' in element.Name:
            result.append(element)

    return result

def get_column_index(headers, name):
    """Возвращает индекс столбца, если нужного столбца нет - останавливает работу

    Args:
        headers: Заголовки из CSV файла
        name: Имя искомого столбца
    """
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

def get_ai_catalog():
    """Возвращает АИ каталог в виде списка из экземпляров AICatalogElement.
    Изначально ищет в сетевых папках, если не получается - проверяем мои документы
    """

    with codecs.open(get_document_path(), 'r', encoding='utf-8-sig') as csvfile:
        material_variants = []
        csvreader = csv.reader(csvfile, delimiter=";")
        headers = next(csvreader)

        rules = CSVRules()

        rules.type_comment_column = get_column_index(headers, 'Комментарий к типоразмеру')
        rules.d_column = get_column_index(headers, 'Диаметр')
        rules.name_column = get_column_index(headers, 'Наименование')
        rules.mark_column = get_column_index(headers, 'Марка')
        rules.code_column = get_column_index(headers, 'Артикул')
        rules.maker_column = get_column_index(headers, 'Завод-изготовитель')
        rules.len_column = get_column_index(headers, 'Длина трубы')

        # Итерируемся по строкам в файле
        for row in csvreader:
            type_comment = row[rules.type_comment_column]

            # если комментария к типоразмеру нет - скорее всего пустая строка или ошибка заполнения. Пропускаем
            if type_comment is None or type_comment == '':
                continue

            material_variants.append(
                AICatalogElement(
                    row[rules.type_comment_column],
                    row[rules.name_column],
                    get_float_value(row[rules.d_column], rules.d_column),
                    row[rules.code_column],
                    get_float_value(row[rules.len_column], rules.len_column),
                    row[rules.mark_column],
                    row[rules.maker_column]
                )
            )

    return material_variants

def convert_to_mms(value):
    """Конвертирует из внутренних значений ревита в миллиметры"""
    result = UnitUtils.ConvertFromInternalUnits(value,
                                               UnitTypeId.Millimeters)
    return result

def get_variants_pool(element, catalog, type_comment, dn):
    """Получаем пул вариантов каталожных значений для элементов. Если варианты не были найдены - останавливаем работу

    Args:
        element: Element для которого ищется пул вариантов
        catalog: Каталог элементов АИ
        type_comment: Комментарий к типоразмеру, по нему сверяемся с каталогом
        dn: Диаметр, по нему сверяемся с каталогом
    """

    result = []

    is_pipe = element.Category.IsId(BuiltInCategory.OST_PipeCurves)
    is_insulation = element.Category.IsId(BuiltInCategory.OST_PipeInsulations)

    # Проверяем, есть ли совпадение по комментарию типоразмера
    if is_pipe:
        for variant in catalog:
            if type_comment == variant.type_comment and dn == variant.dn:
                result.append(variant)

    # Для изоляции очень редок случай точного совпадения диаметров, проверяем с погрешностью
    if is_insulation:
        sorted_catalog = sorted(catalog, key=lambda x: x.dn, reverse=True)
        for variant in sorted_catalog:
            if type_comment == variant.type_comment and variant.dn - 5 <= dn <= variant.dn:
                result.append(variant)

    if len(result) == 0:
        forms.alert("Часть элементов в модели, помеченных как B4E_АИ, не обнаружена в согласованных каталогах, "
                    "что может привести к не полному формированию спецификации. "
                    "Устраните расхождения перед продолжением работы. \n"
                    "Пример элемента - ID:{}".format(element.Id),
                    "Ошибка",
                    exitscript=True)

    return result

def get_dn(ai_element):
    """Получение диаметра элемента"""
    if ai_element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return convert_to_mms(ai_element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
    if ai_element.Category.IsId(BuiltInCategory.OST_PipeInsulations) and ai_element.HostElementId is not None:
        pipe = doc.GetElement(ai_element.HostElementId)
        return convert_to_mms(pipe.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER))

def create_new_row(element, variant, number):
    """Создание нового элемента RowOfSpecification на базе Element из модели для последующей генерации якоря"""

    shared_function = element.GetParamValueOrDefault(
        SharedParamsConfig.Instance.EconomicFunction, unmodeling_factory.out_of_function_value)
    shared_system = element.GetParamValueOrDefault(
        SharedParamsConfig.Instance.VISSystemName, unmodeling_factory.out_of_system_value)
    mark = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISMarkNumber, '')
    maker = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISManufacturer, '')
    unit = 'шт.'  # В этом плагине мы бьем элементы поштучно, поэтому блокируем это значение
    note = element.GetParamValueOrDefault(SharedParamsConfig.Instance.VISNote, '')
    group = '8. Трубопроводы'

    new_row = RowOfSpecification(
        shared_system,
        shared_function,
        group,
        name=variant.name,
        mark=variant.mark,
        code=variant.code,
        maker=variant.maker,
        unit=unit,
        local_description=unmodeling_factory.ai_description,
        number=number,
        mass='',
        note=note
    )

    return new_row

def separate_element(ai_element, variants_pool, pipe_insulation_stock, pipe_stock):
    """
    Делим элемент по длине между вариантами в его пуле,
    и на базе этих отрезков создаем новые RowOfSpecification
    """

    ai_element_len = convert_to_mms(ai_element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH))
    if ai_element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        ai_element_len = ai_element_len * pipe_stock
    if ai_element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
        ai_element_len = ai_element_len * pipe_insulation_stock

    sorted_variants_pool = sorted(variants_pool, key=lambda x: x.length, reverse=True)

    result = []
    for variant in sorted_variants_pool:
        if ai_element_len <= 0:
            break

        number = ai_element_len // variant.length

        if number >= 1:
            ai_element_len = ai_element_len % variant.length

        len_not_minimal = ai_element_len > 50  # Принято что не считаем обрезки труб меньше 50мм
        last_variant = variant == sorted_variants_pool[-1]  # Проверка есть ли еще вариаты

        if last_variant and len_not_minimal:
            number += 1
            ai_element_len -= variant.length

        if number > 0:
            new_row = create_new_row(ai_element, variant, number)

            result.append(
                new_row
            )
    return result

def process_ai_element(ai_element, cash, elements_to_generation, elements_to_update, catalog,
                       pipe_insulation_stock, pipe_stock):
    """Обработка элемента помеченного как B4E_АИ, наполняем список элементов для генерации и кэш типоразмеров для оптимизации

    Args:
        ai_element: Element помеченный как АИ
        elements_to_generation: Список на базе которого будут созданы Якоря, должен обновляться
        elements_to_update: Элементы которые будут обновлены без создания якорей, должен обновляться
        catalog: Каталог АИ
        pipe_insulation_stock: Запас изоляции
        pipe_stock: Запас трубопроводов
    """

    ai_element_type = ai_element.GetElementType()
    id = ai_element_type.Id

    type_comment = ai_element_type.GetParamValue(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)
    ai_element_dn = get_dn(ai_element)

    # Проверяем наличие объекта в кэше
    cached_item = next((item for item in cash if item.dn == ai_element_dn and item.id == id), None)

    if cached_item:
        # Если объект найден в кэше, используем его данные
        variants_pool = cached_item.variants_pool
    else:
        # Если объект не найден в кэше, создаем новый объект и добавляем его в кэш
        variants_pool = get_variants_pool(ai_element, catalog, type_comment, ai_element_dn)
        new_item = TypesCash(ai_element_dn, id, variants_pool)
        cash.append(new_item)

    # Если нет совпадений по комментарию типоразмера - продолжаем перебор
    if len(variants_pool) == 0:
        return cash, elements_to_generation, elements_to_update

    # Если заявленная длина в каталоге 0 - элемент не бьется на части, можно пропускать
    if variants_pool[0].length == 0:
        elements_to_update.append(
            UpdateElement(ai_element, variants_pool[0])
            # Добавляем только 0-ой индекс, т.к.
            # для элементов без длины нет вариативности по диаметрам
        )
        return cash, elements_to_generation, elements_to_update

    generic_elements = separate_element(
        ai_element,
        variants_pool,
        pipe_insulation_stock,
        pipe_stock)
    elements_to_generation.extend(generic_elements)

    return cash, elements_to_generation, elements_to_update

def get_stocks():
    def get_percent_value(info, param):
        percent = info.GetParamValueOrDefault(param)
        if percent is None:
            percent = 0
        return percent

    info = doc.ProjectInformation
    insulation_percent = get_percent_value(info, SharedParamsConfig.Instance.VISPipeInsulationReserve)
    duct_and_pipe_percent = get_percent_value(info, SharedParamsConfig.Instance.VISPipeDuctReserve)

    pipe_insulation_stock = (1 + insulation_percent / 100)
    duct_and_pipe_stock = (1 + duct_and_pipe_percent / 100)

    return pipe_insulation_stock, duct_and_pipe_stock

def optimize_generation_list(new_rows):
    result = []
    unique_rows = {}

    for new_row in new_rows:
        key = (new_row.system, new_row.function, new_row.group, new_row.name, new_row.mark,
               new_row.code, new_row.maker, new_row.unit, new_row.local_description, new_row.mass, new_row.note)
        if key in unique_rows:
            unique_rows[key].number += new_row.number
        else:
            unique_rows[key] = RowOfSpecification(
                system=new_row.system,
                function=new_row.function,
                group=new_row.group,
                name=new_row.name,
                mark=new_row.mark,
                code=new_row.code,
                maker=new_row.maker,
                unit=new_row.unit,
                local_description=new_row.local_description,
                number=new_row.number,
                mass=new_row.mass,
                note=new_row.note
            )

    result.extend(unique_rows.values())
    return result

def update_element(element, variant):
    element.SetParamValue(SharedParamsConfig.Instance.VISCombinedName, variant.name)
    element.SetParamValue(SharedParamsConfig.Instance.VISMarkNumber, variant.mark)
    element.SetParamValue(SharedParamsConfig.Instance.VISItemCode, variant.code)
    element.SetParamValue(SharedParamsConfig.Instance.VISManufacturer, variant.maker)

def show_dialog(instr, content=''):
    dialog = TaskDialog("Внимание")
    dialog.MainInstruction = instr
    dialog.MainContent = content
    dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No

    result = dialog.Show()

    if result == TaskDialogResult.Yes:
        return True
    elif result == TaskDialogResult.No:
        return False

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    # Стартовые проверки и поиск семейства якоря
    family_symbol = unmodeling_factory.startup_checks(doc)

    # Получаем запасы из сведений. Если параметра нет - запасов тоже нет(1)
    pipe_insulation_stock, pipe_stock = get_stocks()

    # Получаем согласованный каталог из сетевой папки или из моих документов при отсутствии доступа
    ai_catalog = get_ai_catalog()

    elements = list(chain(
        unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_PipeCurves),
        unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_PipeInsulations)
    ))

    # Фильтруем те элементы у которых в имени типа есть "_B4E_АИ"
    ai_elements = filter_elements_ai(elements)

    cash = []  # сохранение типоразмеров, чтоб не перебирать для каждой трубы каталог
    elements_to_generation = []
    elements_to_update = []
    for ai_element in ai_elements:
        # Собираем элементы для генерации через сопоставление комментария к типоразмеру и диаметра элемента
        cash, elements_to_generation, elements_to_update = process_ai_element(ai_element,
                                                                             cash,
                                                                             elements_to_generation,
                                                                             elements_to_update,
                                                                             ai_catalog,
                                                                             pipe_insulation_stock,
                                                                             pipe_stock)

    with revit.Transaction("BIM: Добавление расчетных элементов"):
        # При каждом запуске затираем расходники с соответствующим описанием и генерируем заново
        unmodeling_factory.remove_models(doc, unmodeling_factory.ai_description)

        family_symbol.Activate()

        material_location = unmodeling_factory.get_base_location(doc)

        # На данном этапе элементы созданы для каждого прямого участка трубы.
        # Для оптимизации работы превращаем одинаковые элементы в один, складывая их числа
        elements_to_generation = optimize_generation_list(elements_to_generation)

        for element in elements_to_generation:
            material_location = unmodeling_factory.update_location(material_location)

            unmodeling_factory.create_new_position(doc, element, family_symbol,
                                                  unmodeling_factory.ai_description,
                                                  material_location)

        for element in elements_to_update:
            update_element(element.element, element.data)

script_execute()