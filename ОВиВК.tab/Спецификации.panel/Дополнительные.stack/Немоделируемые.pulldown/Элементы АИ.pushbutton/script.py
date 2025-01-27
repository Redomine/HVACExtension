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

class CSVRules:
    name_column = 0
    d_column = 0
    code_column = 0
    len_column = 0
    thickness_column = 0

class GenericPipe:
    def __init__(self, name, dn, code, length, thickness):
        self.name = name
        self.dn = dn
        self.code = code
        self.length = length
        self.thickness = thickness

def filter_elements_ai(elements):
    result = []
    for element in elements:
        if '_B4E_АИ' in element.Name:
            result.append(element)

    return result

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    path = get_document_path()

    with codecs.open(path + '/Трубопроводы АИ.csv', 'r', encoding='cp1251') as csvfile:
        pipe_variants = []
        # Создаем объект reader
        csvreader = csv.reader(csvfile, delimiter=";")
        headers = next(csvreader)

        rules = CSVRules()

        rules.name_column = headers.index('Наименование')
        rules.d_column = headers.index('Диаметр условный')
        rules.code_column = headers.index('Артикул')
        rules.len_column = headers.index('Длина трубы')
        rules.thickness_column = headers.index('Толщина трубы')

        # Итерируемся по строкам в файле
        for row in csvreader:
            pipe_variants.append(
                GenericPipe(
                    row[rules.name_column],
                    row[rules.d_column],
                    row[rules.code_column],
                    row[rules.len_column],
                    row[rules.thickness_column]
                )
            )

    family_symbol = unmodeling_factory.startup_checks(doc)

    pipes = unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_PipeCurves)

    ai_pipes = filter_elements_ai(pipes)

    print(ai_pipes)
    with revit.Transaction("BIM: Добавление расчетных элементов"):
        family_symbol.Activate()

        # При каждом запуске затираем расходники с соответствующим описанием и генерируем заново
        unmodeling_factory.remove_models(doc, unmodeling_factory.ai_description)




script_execute()
