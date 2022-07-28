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
from rpw.ui.forms import SelectFromList

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *


doc = __revit__.ActiveUIDocument.Document  # type: Document
uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application
uidoc = __revit__.ActiveUIDocument


def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

def convert(value):
    unit_type = DisplayUnitType.DUT_MILLIMETERS
    new_v = UnitUtils.ConvertFromInternalUnits(value, unit_type)
    return new_v


def pick_elements(uidoc):
    result = []
    message = "Выберите элементы"
    ob_type = Selection.ObjectType.Element
    for ref in uidoc.Selection.PickObjects(ob_type, message):
        result.append(doc.GetElement(ref))
    return result


def check_is_nested(element):
    if hasattr(element, "SuperComponent"):
        if not element.SuperComponent:
            return False
    if hasattr(element, "HostRailingId"):
        return True
    if hasattr(element, "GetStairs"):
        return True
    return False


def filter_elements(elements):
    result = []
    for element in elements:
        if str(element.GroupId) == "-1":
            if not check_is_nested(element):
                result.append(element)
    return result


def get_real_height(doc, element, height, height_offset):
    level_id = element.LookupParameter(height).AsElementId()
    level = doc.GetElement(level_id)
    height_value = level.Elevation
    height_offset_value = element.LookupParameter(height_offset).AsDouble()
    real_height = height_value + height_offset_value
    return real_height


def find_parameter(element, parameter_name):
    for parameter in element.GetParameters(parameter_name):
        if not parameter.IsReadOnly:
            return parameter


def get_height_by_element(doc, element):
    parameters = {
        'Смещение начала от уровня': "Базовый уровень",
        'Отметка от уровня': ["Уровень спецификации", "Уровень"],
        'Отметка посередине': "Базовый уровень",
        }
    for offset_param_name in parameters.keys():
        if element.LookupParameter(offset_param_name):
            height_param_name = parameters[offset_param_name]
            offset_param = find_parameter(element, offset_param_name)
            if isinstance(height_param_name, list):
                for var_param_height_name in height_param_name:
                    var_param_height = find_parameter(element, var_param_height_name)
                    if var_param_height:
                            real_height = get_real_height(doc, element, var_param_height_name, offset_param_name)
                            return [real_height, offset_param, var_param_height]
                return False
            else:
                param_height = find_parameter(element, height_param_name)
                real_height = get_real_height(doc, element, height_param_name, offset_param_name)
                return [real_height, offset_param, param_height]


def find_new_level(height):
    all_levels = FilteredElementCollector(doc).OfClass(Level)
    new_offset = 10000
    for level in all_levels:
        level_height = level.Elevation
        offset = height - level_height
        if offset < new_offset and offset >= 0:
            new_level = level
            new_offset = offset
    return new_level, new_offset


def change_level(element, new_level, new_offset, offset_param, height_param):
    height_param.Set(new_level.Id)
    offset_param.Set(new_offset)
    return element

method = SelectFromList('Выберите метод привязки', ['Все элементы на активном виде к ближайшим уровням', 'Все элементы на активном виде к выбранному уровню', 'Выбранные элементы к выбранному уровню'])

if method != 'Все элементы на активном виде к ближайшим уровням':
    selected_view = True

    levelCol = make_col(BuiltInCategory.OST_Levels)

    levels = []

    for levelEl in levelCol:
        levels.append(levelEl.Name)

    level_name = SelectFromList('Выберите уровень', levels)

    for levelEl in levelCol:
        if levelEl.Name == level_name:
            level = levelEl

else:
    selected_view = False




try:
    if method == 'Выбранные элементы к выбранному уровню':
        elements = pick_elements(uidoc)
    if method == 'Все элементы на активном виде к выбранному уровню' or method == 'Все элементы на активном виде к ближайшим уровням':
        elements = FilteredElementCollector(doc, doc.ActiveView.Id)
except Exception:
    print "Элементы не выбраны"
    sys.exit()

filtered = filter_elements(elements)

result = []
result_error = []
result_ok = []



# with Transaction(doc, "Change level") as t:
#     t.Start()
with revit.Transaction("Смена уровней"):
    try:
        for element in filtered:
            height_result = get_height_by_element(doc, element)
            if height_result:
                real_height = height_result[0]
                offset_param = height_result[1]
                height_param = height_result[2]
                new_level, new_offset = find_new_level(real_height)
                if selected_view:
                    new_offset = real_height - level.Elevation
                    change_level(element, level, new_offset, offset_param, height_param)
                else:
                    change_level(element, new_level, new_offset, offset_param, height_param)
                result_ok.append(element)
                # result_ok.append(convert(real_height)
                # result_ok.append(new_level)
                # result_ok.append(convert(new_offset))
            else:
                # result_error.append("Error with:")
                result_error.append(element)
            # t.Commit()
    except Exception:
        print 'Ошибка переноса для элементов с айди'
        print element.Id


#
# result.append(result_ok)
# result.append(result_error)

