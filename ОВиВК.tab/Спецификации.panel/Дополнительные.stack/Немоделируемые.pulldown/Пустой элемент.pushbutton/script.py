#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пустой элемент'
__doc__ = "Генерирует в модели пустой якорный элемент"

from itertools import count

import clr

from UnmodelingClassLibrary import UnmodelingFactory, MaterialCalculator, RowOfSpecification

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from pyrevit import forms
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict
from UnmodelingClassLibrary import  *
from dosymep_libs.bim4everyone import *


#Исходные данные
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
selected_ids = uidoc.Selection.GetElementIds()
unmodeling_factory = UnmodelingFactory()
description = 'Пустая строка'


def get_new_position(location, family_symbol, rows_number):
    element = doc.GetElement(selected_ids[0])

    parent_system, parent_function = unmodeling_factory.get_system_function(element)
    parent_group = element.GetSharedParamValueOrDefault(SharedParamsConfig.Instance.VISGrouping.Name, '')

    for count in range(1, rows_number + 1):
        new_group = "{}{}".format(parent_group, '_' + str(count))

        new_position = RowOfSpecification(
            parent_system,
            parent_function,
            new_group
        )

        location = XYZ(0, location.Y + 10, 0)
        unmodeling_factory.create_new_position(doc, new_position, family_symbol, description, location)

def get_location(generic_models):
    # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
    filtered_generics = \
        [elem for elem in generic_models if elem.GetElementType()
        .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == '_Якорный элемент']

    count = 0
    for generic in filtered_generics:
        if generic.GetSharedParamValue('ФОП_ВИС_Назначение') == description:
            count+=1

    return XYZ(0, 10+10*count, 0)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):


    family_symbol = unmodeling_factory.startup_checks(doc)

    if view.Category == None or not view.Category.IsId(BuiltInCategory.OST_Schedules):
        forms.alert(
            "Добавление пустого элемента возможно только на целевой спецификации.",
            "Ошибка",
            exitscript=True)

    if 0 == selected_ids.Count:
        forms.alert(
            "Выделите целевой элемент",
            "Ошибка",
            exitscript=True)

    rows_number = forms.ask_for_string(
        default='1',
        prompt='Введите количество пустых строк:',
        title=__title__
    )

    try:
        rows_number = int(rows_number)
    except ValueError:
        forms.alert(
            "Нужно ввести число.",
            "Ошибка",
            exitscript=True)

    with revit.Transaction("Добавление пустого элемента"):
        generic_models = unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_GenericModel)
        location = get_location(generic_models)

        get_new_position(location, family_symbol, rows_number)



script_execute()