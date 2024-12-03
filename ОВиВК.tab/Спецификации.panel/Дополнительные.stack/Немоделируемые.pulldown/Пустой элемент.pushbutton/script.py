#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пустой элемент'
__doc__ = "Генерирует в модели пустой якорный элемент"


import clr

from UnmodelingClassLibrary import UnmodelingFactory, MaterialCalculator, RowOfSpecification

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
uidoc = __revit__.ActiveUIDocument
selectedIds = uidoc.Selection.GetElementIds()
nameOfModel = '_Якорный элемент'
description = 'Пустая строка'


def get_new_position():
    element = doc.GetElement(selectedIds[0])

    system = element.LookupParameter('ФОП_ВИС_Имя системы').AsString()
    parent_group = element.LookupParameter('ФОП_ВИС_Группирование').AsString()
    parent_function = element.LookupParameter('ФОП_Экономическая функция').AsString()
    group = parent_group + '_1'

    new_position = RowOfSpecification()

    new_position.system = system
    new_position.group = group
    new_position.function = parent_function

    return new_position

def get_location(family_name, generic_models):
    # Фильтруем элементы, чтобы получить только те, у которых имя семейства равно "_Якорный элемент"
    filtered_generics = \
        [elem for elem in generic_models if elem.GetElementType()
        .GetParamValue(BuiltInParameter.ALL_MODEL_FAMILY_NAME) == family_name]

    count = 0
    for generic in filtered_generics:
        if generic.GetSharedParamValue('ФОП_ВИС_Назначение') == description:
            count+=1

    return XYZ(0, 10+10*count, 0)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with revit.Transaction("Добавление пустого элемента"):
        family_name = "_Якорный элемент"

        if doc.IsFamilyDocument:
            forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

        unmodeling_factory = UnmodelingFactory()

        family_symbol = unmodeling_factory.is_family_in(doc, family_name)

        if family_symbol is None:
            forms.alert(
                "Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.",
                "Ошибка",
                exitscript=True)

        if not view.Category.IsId(BuiltInCategory.OST_Schedules):
            forms.alert(
                "Добавление пустого элемента возможно только на целевой спецификации.",
                "Ошибка",
                exitscript=True)

        unmodeling_factory = UnmodelingFactory()

        generic_models = unmodeling_factory.get_elements_by_category(doc, BuiltInCategory.OST_GenericModel)
        location = get_location(family_name, generic_models)

        unmodeling_factory.create_new_position(doc, get_new_position(), family_symbol, family_name, description, location)


script_execute()