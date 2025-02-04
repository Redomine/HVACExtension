# -*- coding: utf-8 -*-
import sys
import clr


clr.AddReference('ProtoGeometry')
clr.AddReference("RevitNodes")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import Revit
import dosymep
import codecs
import math

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

import System
from System.Collections.Generic import *


from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import Selection
from Autodesk.DesignScript.Geometry import *


import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from rpw.ui.forms import select_file


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *


doc = __revit__.ActiveUIDocument.Document  # type: Document
uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application
uidoc = __revit__.ActiveUIDocument


class AuditorEquipment:
    processed = False

    def __init__(self,
                 connection_type,
                 x,
                 y,
                 z,
                 len,
                 code,
                 real_power,
                 nominal_power,
                 setting,
                 maker,
                 full_name):
        self.connection_type = connection_type
        self.x = x
        self.y = y
        self.z = z
        self.len = len
        self.code = code
        self.real_power = real_power
        self.nominal_power = nominal_power
        self.setting = setting
        self.maker = maker
        self.full_name = full_name


class ReadingRules:
    connection_type_index = 2
    x_index = 3
    y_index = 4
    z_index = 5
    len_index = 12
    code_index = 16
    real_power_index = 20
    nominal_power_index = 22
    setting_index = 28
    maker_index = 30
    full_name_index = 31

class RevitXYZmms:
    def __init__(self, x, y, z):
        self.x = x
        self.y = y
        self.z = z

def convert_to_mms(value):
    """Конвертирует из внутренних значений ревита в миллиметры"""
    result = UnitUtils.ConvertFromInternalUnits(value,
                                               UnitTypeId.Millimeters)
    return result

def distance(point1, point2):
    """ Вычисляет расстояние между двумя точками в 3D пространстве """
    return math.sqrt((point1.x - point2.x)**2 + (point1.y - point2.y)**2 + (point1.z - point2.z)**2)

def is_within_sphere(auditor_equipment, revit_coords, radius=1500):
    """ Проверяет, входит ли точка auditor_equipment в сферу радиуса radius вокруг revit_coords """
    return distance(auditor_equipment, revit_coords) <= radius

def extract_heating_device_description(file_path, reading_rules):
    with codecs.open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    equipment = []
    i = 0

    while i < len(lines):
        if "Отопительные приборы CO на плане" in lines[i]:
            description_start_index = i + 3
            i = description_start_index

            while i < len(lines) and lines[i].strip() != "":
                data = lines[i].strip().split(';')
                equipment.append(AuditorEquipment(
                    data[reading_rules.connection_type_index],
                    float(data[reading_rules.x_index].replace(',', '.')) * 1000,
                    float(data[reading_rules.y_index].replace(',', '.')) * 1000,
                    float(data[reading_rules.z_index].replace(',', '.')) * 1000,
                    float(data[reading_rules.len_index].replace(',', '.'))/304.8,
                    data[reading_rules.code_index],
                    float(data[reading_rules.real_power_index])  * 10.764,
                    float(data[reading_rules.nominal_power_index] ) * 10.764,
                    float(data[reading_rules.setting_index]),
                    data[reading_rules.maker_index],
                    data[reading_rules.full_name_index]
                ))
                i += 1

        i += 1

    if not equipment:
        forms.alert("Строка 'Отопительные приборы CO на плане' не найдена в файле.", "Ошибка", exitscript=True)

    return equipment

def get_elements_by_category(category):
    """ Возвращает коллекцию элементов по категории """
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

def insert_data(element, auditor_data):
    element.SetParamValue('ADSK_Размер_Длина', auditor_data.len)
    element.SetParamValue('ADSK_Код изделия', auditor_data.code)
    element.SetParamValue('ADSK_Настройка', auditor_data.setting)
    element.SetParamValue('ADSK_Тепловая мощность', auditor_data.real_power)


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True )

    filepath = select_file('Файл расчетов (*.txt)|*.txt')

    reading_rules = ReadingRules()

    ayditror_equipment = extract_heating_device_description(filepath, reading_rules)

    equipment = get_elements_by_category(BuiltInCategory.OST_MechanicalEquipment)

    not_found_revit = []
    not_found_ayditor = []

    with revit.Transaction("BIM: Импорт расчетов"):
        for element in equipment:
            was_found = False
            xyz = element.Location.Point

            revit_coords = RevitXYZmms(
                convert_to_mms(xyz.X),
                convert_to_mms(xyz.Y),
                convert_to_mms(xyz.Z - element.GetParamValue(BuiltInParameter.INSTANCE_ELEVATION_PARAM))
            )

            for auditor_data in ayditror_equipment:
                if is_within_sphere(auditor_data, revit_coords):
                    family_name = element.Symbol.Family.Name
                    if 'Обр_ОП_Универсальный' in family_name:
                        was_found = True
                        insert_data(element, auditor_data)
            if not was_found:
                not_found_revit.append(element.Id.IntegerValue)

    if len(not_found_revit) > 0:
        print('ID Элементов Revit, которые не были обработаны:')
        print(not_found_revit)


script_execute()