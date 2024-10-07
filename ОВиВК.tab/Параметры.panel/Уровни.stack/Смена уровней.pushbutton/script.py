# -*- coding: utf-8 -*-
import clr
import sys
import System
from System.Collections.Generic import *
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from pyrevit import revit
from pyrevit import forms
from rpw.ui.forms import SelectFromList
from Redomine import *
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
import sys
import paraSpec
from Autodesk.Revit.DB import *
from System import Guid
from pyrevit import revit
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

doc = __revit__.ActiveUIDocument.Document  # type: Document
uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application
uidoc = __revit__.ActiveUIDocument

# типы параметров отвечающих за уровень
built_in_level_params = [
                         BuiltInParameter.RBS_START_LEVEL_PARAM,
                         BuiltInParameter.FAMILY_LEVEL_PARAM,
                         BuiltInParameter.GROUP_LEVEL]
# BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,

# типы параметров отвечающих за смещение от уровня
built_in_offset_params = [BuiltInParameter.INSTANCE_ELEVATION_PARAM,
                          BuiltInParameter.RBS_OFFSET_PARAM,
                          BuiltInParameter.GROUP_OFFSET_FROM_LEVEL,
                          BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM]

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

def get_collection(category):
    """ Возвращает коллекцию элементов по категории """
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

def convert(value):
    """ Преобразует дабл в миллиметры """
    unit_type = DisplayUnitType.DUT_MILLIMETERS
    new_v = UnitUtils.ConvertFromInternalUnits(value, unit_type)
    return new_v

def pick_elements(uidoc):
    """ Возвращает выбранные элементы """
    result = []
    message = "Выберите элементы"
    ob_type = Selection.ObjectType.Element
    for ref in uidoc.Selection.PickObjects(ob_type, message):
        result.append(doc.GetElement(ref))
    return result

def check_is_nested(element):
    """ Проверяет, является ли вложением """
    if hasattr(element, "SuperComponent"):
        if not element.SuperComponent:
            return False
    if hasattr(element, "HostRailingId"):
        return True
    if hasattr(element, "GetStairs"):
        return True
    return False

def get_parameter_if_exist_not_ro(element, built_in_parameters):
    """ Получает параметр, если он существует и если он не ReadOnly """
    for built_in_parameter in built_in_parameters:
        parameter = element.get_Parameter(built_in_parameter)
        if parameter is not None and not parameter.IsReadOnly:
            return built_in_parameter

    return None

def filter_elements(elements):
    """Возвращает фильтрованный от вложений и от свободных от групп список элементов"""
    result = []
    for element in elements:
        if element.GroupId == ElementId.InvalidElementId:
            builtin_level_param = get_parameter_if_exist_not_ro(element, built_in_level_params)
            builtin_offset_param = get_parameter_if_exist_not_ro(element, built_in_offset_params)

            if builtin_level_param is None or builtin_offset_param is None:
                continue

            # Даже если у элемента нашелся builtin - все равно просто параметра может и не быть.
            if not element.IsExistsParam(LabelUtils.GetLabelFor(builtin_level_param)):
                continue

            if not element.IsExistsParam(LabelUtils.GetLabelFor(builtin_offset_param)):
                continue

            # проверяем вложение или нет
            if not check_is_nested(element):
                result.append(element)

    return result

def get_real_height(doc, element, level_param_name, offset_param_name):
    """ Возвращает реальную абсолютную отметку элемента """
    level_id = element.GetParam(level_param_name).AsElementId()
    level = doc.GetElement(level_id)
    height_value = level.Elevation
    height_offset_value = element.GetParam(offset_param_name).AsDouble()
    real_height = height_value + height_offset_value
    return real_height

def find_parameter(element, parameter_name):
    for parameter in element.GetParameters(parameter_name):
        if not parameter.IsReadOnly:
            return parameter

def get_height_by_element(doc, element):
    """ Возвращает абсолютную отметку, параметр смещения и параметр уровня """

    level_builtin_param = get_parameter_if_exist_not_ro(element, built_in_level_params)
    offset_builtin_param = get_parameter_if_exist_not_ro(element, built_in_offset_params)

    #print element.Id
    real_height = get_real_height(doc, element, level_builtin_param, offset_builtin_param)
    level_param = element.GetParam(level_builtin_param)
    offset_param = element.GetParam(offset_builtin_param)

    return [real_height, offset_param, level_param]

def find_new_level(height):
    """ Ищем новый уровень. Здесь мы собираем лист из всех уровней, вычисляем у какого из них минимальное неотрцицательное(при наличии) смещение
    от нашей точки. Он и будет целевым """
    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

    sorted_levels = sorted(all_levels, key=lambda level: level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())

    offsets = []
    for level in sorted_levels:
        offsets.append(height - level.Elevation)

    target_ind = -1
    # мы проходим по всем смещениям и ищем первое не отрицательное, т.к. просто минимальное смещение будет для уровней которые сильно выше реальной отметки(отрицательное)
    for offset in offsets:
        if offset > 0:
            target_ind = offsets.index(offset)

    # если целевой индекс остался -1 - значит мы не нашли нормальных уровней с верным смещением. Берем просто минимальный
    if target_ind == -1:
        target_ind = offsets.index(min(offsets))

    level = sorted_levels[target_ind]
    offset_from_new_level = offsets[target_ind]

    return level, offset_from_new_level

def change_level(element, new_level, new_offset, offset_param, height_param):
    height_param.Set(new_level.Id)
    offset_param.Set(new_offset)
    return element

def get_selected_mode():
    method = forms.SelectFromList.show(["Все элементы на активном виде к ближайшим уровням",
                                        "Все элементы на активном виде к выбранному уровню",
                                        "Выбранные элементы к выбранному уровню"],
                                       title="Выберите метод привязки",
                                       button_name="Применить")
    return method

def get_selected_level(method):
    if method != 'Все элементы на активном виде к ближайшим уровням':
        selected_view = True

        levelCol = get_collection(BuiltInCategory.OST_Levels)

        levels = []

        for levelEl in levelCol:
            levels.append(levelEl.Name)

        level_name = forms.SelectFromList.show(levels,
                                               title="Выберите уровень",
                                               button_name="Применить")

        for levelEl in levelCol:
            if levelEl.Name == level_name:
                level = levelEl
                return level

    return False

def get_list_of_elements(method):
    if method == 'Выбранные элементы к выбранному уровню':
        elements = pick_elements(uidoc)
    if (method == 'Все элементы на активном виде к выбранному уровню'
            or method == 'Все элементы на активном виде к ближайшим уровням'):
        elements = FilteredElementCollector(doc, doc.ActiveView.Id)

    filtered = filter_elements(elements)

    if len(filtered) == 0:
        print "Элементы не выбраны"
        sys.exit()

    return filtered

def main():
    result = []
    result_error = []
    result_ok = []

    method = get_selected_mode()
    elements = get_list_of_elements(method)
    level = get_selected_level(method)

    with revit.Transaction("Смена уровней"):
            for element in elements:
                height_result = get_height_by_element(doc, element)

                if height_result:
                    real_height = height_result[0]
                    offset_param = height_result[1]
                    height_param = height_result[2]

                    if level:
                        new_offset = real_height - level.Elevation
                        change_level(element, level, new_offset, offset_param, height_param)
                    else:
                        new_level, new_offset = find_new_level(real_height)
                        change_level(element, new_level, new_offset, offset_param, height_param)
                    result_ok.append(element)
                else:
                    result_error.append(element)

main()