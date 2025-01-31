#! /usr/bin/env python
# -*- coding: utf-8 -*-
from tarfile import TUEXEC

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import glob
import re
import sys
import json
import os
import ctypes
import codecs
from low_voltage_task_class_lib import JsonOperator, EditedReport, LowVoltageSystemData

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict

from dosymep_libs.bim4everyone import *
from System.Collections.Generic import List

from datetime import datetime, timedelta
from System import Environment

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

def split_valves_by_floors(valves):
    ''' Делим список элементов на словарь, ключ - номер этажа, значение - экземпляр LowVoltageSystemData '''
    valves_by_floors = {}

    for valve in valves:
        floor_name = valve.GetParamValueOrDefault(FLOOR_PARAM)
        system_name = valve.GetParamValueOrDefault(SYSTEM_PARAM)

        is_floor_name_exists = floor_name is None or floor_name == ""
        is_system_name_exists = system_name is None or system_name == "!Нет системы" or system_name == ""

        if is_floor_name_exists or is_system_name_exists:
            forms.alert("У части отмеченного для задания оборудования не заполнено ФОП_ВИС_Имя системы или ФОП_Этаж. "
                        "Устраните проблему и повторите запуск скрипта.", "Ошибка", exitscript=True)


        if floor_name not in valves_by_floors:
            valves_by_floors[floor_name] = []


        valve_base_name = system_name + "-" + floor_name

        valves_by_floors[floor_name].append(LowVoltageSystemData
            (
                valve.Id,
                creation_date=operator.get_moscow_date(),
                valve_base_name= valve_base_name,
                element=valve
            )
        )

    return valves_by_floors

def use_open_algorithm(valves, max_numbers):
    ''' Создаем имена для списка элементов используя алгоритм для открытых клапанов от СППЗ '''
    valves_by_floors = split_valves_by_floors(valves)

    result = []
    for floor_name, valve_data_instances in valves_by_floors.items():
        count = 0
        for valve_data in valve_data_instances:
            if valve_data.element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
                key = valve_data.valve_base_name
            else:
                key = "НО-" + valve_data.valve_base_name



            # Проверяем, существует ли ключ в словаре max_numbers
            if key in max_numbers and count < max_numbers[key]:
                count = max_numbers[key] + 1
            else:
                count += 1

            valve_data.json_name = key + "." + str(count)

            result.append(valve_data)

    return result

def use_closed_algorithm(valves):
    ''' Создаем имена для списка элементов используя алгоритм для закрытых клапанов от СППЗ '''
    valves_by_floors = split_valves_by_floors(valves)
    result = []

    for floor_name, valve_data_instances in valves_by_floors.items():
        # Группировка valve_data_instances по valve_base_name
        valve_groups = {}
        for valve_data in valve_data_instances:
            if valve_data.valve_base_name not in valve_groups:
                valve_groups[valve_data.valve_base_name] = []
            valve_groups[valve_data.valve_base_name].append(valve_data)

        # проходим по экземплярам НЗ сгруппированных по базовому имени(одинаковая система и этаж)
        for valve_base_name, valve_data_instances_group in valve_groups.items():
            count = 0
            for valve_data in valve_data_instances_group:
                count += 1
                valve_data.json_name = "НЗ-" + valve_data.valve_base_name + "-" + str(count)

                result.append(valve_data)

    return result

def split_collection(equipment_collection):
    ''' Делим список элементов на открытые клапана, закрытые и оборудование. Если один из элементов '''
    open_valves = []
    closed_valves = []
    equipment_elements = []

    for element in equipment_collection:

        if element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
            equipment_elements.append(element)

        if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            mark = element.GetSharedParamValueOrDefault(MARK_PARAM)
            if mark is not None and mark != "":
                if "НО" in mark:
                    open_valves.append(element)
                if "НЗ" in mark:
                    closed_valves.append(element)


    return open_valves, closed_valves, equipment_elements

def get_elements():
    ''' Забираем список элементов арматуры и оборудования '''
    categories = [
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PipeAccessory
    ]

    category_ids = List[ElementId]([ElementId(int(category)) for category in categories])

    multicategory_filter = ElementMulticategoryFilter(category_ids)

    elements = FilteredElementCollector(doc) \
        .WherePasses(multicategory_filter) \
        .WhereElementIsNotElementType() \
        .ToElements()

    return elements

def get_elements_to_objective(elements):
    ''' Фильтруем, у каких элементов модели стоит галочка для добавления в задание '''
    filtered_elements = []
    for element in elements:

        if element.GetParamValueOrDefault(CREATE_TASK_SS_PARAM) == 1:
            filtered_elements.append(element)
        else:
            element_type = element.GetElementType()
            if element_type.GetParamValueOrDefault(CREATE_TASK_SS_PARAM) == 1:
                filtered_elements.append(element)

    return  filtered_elements

def math_elements_to_old_data(elements, old_data):
    ''' Сравниваем список элементов с данными из json-файла, возвращая только те которых в нем нет '''
    new_elements = []

    for element in elements:
        if not any(element.Id == data.id for data in old_data):
            new_elements.append(element)

    return new_elements

def get_max_numbers(old_data):
    ''' Вычисляем для каждого базового имени в json-файле максимальный номер, с которого будет продолжаться нумерация '''
    # Словарь для хранения максимальных порядковых номеров
    max_numbers = defaultdict(int)

    # Регулярное выражение для извлечения числовой части
    pattern = re.compile(r"(.+)\.(\d+)")
    for data in old_data:
        match = re.search(r'(.*)\.(\d+)$', data.json_name)

        if match:
            base_name = match.group(1)
            number = int(match.group(2))
            if number > max_numbers[base_name]:
                max_numbers[base_name] = number

    return max_numbers

def check_edited_elements(elements):
    ''' Проверяем занят ли элемент '''
    edited_report = EditedReport(doc)

    for element in elements:
        edited_report.is_elemet_edited(element)

    edited_report.show_report()

def clear_param_false_values(elements, json_data):
    ''' Проверяем что в данных которые мы получили/будем загружать в json существуют айди оборудования.
    Если их не существует - обнуляем значения марок, считается что это ошибка '''
    for element in elements:
        if not any(element.Id == data.id for data in json_data):
            element.SetParamValue(TASK_SS_PARAM, "")
            element.SetParamValue(DATE_SS_PARAM, "")

def setup_params():
    revit_params = [MARK_PARAM,
                    FLOOR_PARAM,
                    SYSTEM_PARAM
                    ]

    project_parameters = ProjectParameters.Create(self.doc.Application)
    project_parameters.SetupRevitParams(self.doc, revit_params)


MARK_PARAM = SharedParamsConfig.Instance.VISMarkNumber
FLOOR_PARAM = SharedParamsConfig.Instance.Level
SYSTEM_PARAM = SharedParamsConfig.Instance.VISSystemName
TASK_SS_PARAM = 'ФОП_ВИС_СС Марка задания'
DATE_SS_PARAM = 'ФОП_ВИС_СС Дата задания'
CREATE_TASK_SS_PARAM = 'ФОП_ВИС_СС Сформировать марку'


operator = JsonOperator(doc, uiapp)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

    file_folder_path = operator.get_document_path()

    # Получаем данные из последнего по дате редактирования файла
    old_data = operator.get_json_data(file_folder_path)

    # Получаем максимальные номера из старых данных для продолжения нумерации. Если данных нет то пустой словарь
    max_numbers = get_max_numbers(old_data)

    # Получаем список элементов, без фильтрации
    raw_collection = get_elements()

    # Если часть элементов занята, то задание может выйти не актуальным. Сразу останавливаемся,  если так
    check_edited_elements(raw_collection)

    # Матчим с старыми элементами для фильтрации тех что являются новыми
    new_elements = math_elements_to_old_data(raw_collection, old_data)

    # Оставляем только те новые элементы, у которых стоит галочка о формировании задания
    elements_to_objective = get_elements_to_objective(new_elements)

    # если новых элементов для задания нет то на этом можно остановиться
    if len(elements_to_objective) == 0:
        with revit.Transaction("BIM: Обновление задания"):
            for data in old_data:
                data.insert(doc, operator.get_moscow_date())

            # Проверяем все элементы на предмет ложно-заполненности(У элемента есть марка, но его не было в старых заданиях)
            clear_param_false_values(raw_collection, old_data)

            # Записываем в json-файл
            operator.send_json_data(old_data, file_folder_path)
            return


    # Делим новые элементы по спискам для применения отдельных алгоритмов
    open_valves, closed_valves, equipment_elements = split_collection(elements_to_objective)

    open_valves_data = use_open_algorithm(open_valves, max_numbers)
    closed_valves_data = use_closed_algorithm(closed_valves)
    equipment_elements_data = use_open_algorithm(equipment_elements, max_numbers)

    with revit.Transaction("BIM: Обновление задания"):
        if open_valves_data or closed_valves_data or equipment_elements_data:
            json_data = []
            json_data.extend(old_data)
            json_data.extend(open_valves_data)
            json_data.extend(closed_valves_data)
            json_data.extend(equipment_elements_data)

            # вставляем новые данные в проект. Если элемента нет в проекте - указываем для него дату удаления. Если дата удаления совпадает с сегодняшней - игнорируем, это просто рабочие правки
            for low_voltage_system_data in json_data:
                low_voltage_system_data.insert(doc, operator.get_moscow_date())

            # Записываем в json-файл
            operator.send_json_data(json_data, file_folder_path)

        # Проверяем все элементы на предмет ложно-заполненности(У элемента есть марка, но его не было в старых заданиях или в новых данных для json)
        clear_param_false_values(raw_collection, json_data)

script_execute()
