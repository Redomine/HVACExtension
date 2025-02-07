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

def is_within_sphere(auditor_equipment, revit_coords, radius=1000):

    """ Проверяет, входит ли точка auditor_equipment в сферу радиуса radius вокруг revit_coords """
    # if(distance(auditor_equipment, revit_coords) <= radius):
    #     print(distance(auditor_equipment, revit_coords))

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
                    float(data[reading_rules.len_index].replace(',', '.')),
                    data[reading_rules.code_index],
                    float(data[reading_rules.real_power_index]),
                    float(data[reading_rules.nominal_power_index]),
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
    real_power_watts = UnitUtils.ConvertToInternalUnits(auditor_data.real_power, UnitTypeId.Watts)
    len_meters = UnitUtils.ConvertToInternalUnits(auditor_data.len, UnitTypeId.Millimeters)

    element.SetParamValue('ADSK_Размер_Длина', len_meters)
    element.SetParamValue('ADSK_Код изделия', auditor_data.code)
    element.SetParamValue('ADSK_Настройка', auditor_data.setting)
    element.SetParamValue('ADSK_Тепловая мощность', real_power_watts)

def get_bb_center(bb):
    minPoint = bb.Min
    maxPoint = bb.Max

    centroid = XYZ(
        (minPoint.X + maxPoint.X) / 2,
        (minPoint.Y + maxPoint.Y) / 2,
        (minPoint.Z + maxPoint.Z) / 2
    )
    return centroid

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
            xyz = element.Location.Point
            bb = element.GetBoundingBox()
            bb_center = get_bb_center(bb)
            elem_height = 0
            if element.IsExistsParam('ADSK_Размер_Высота'):
                elem_height = element.GetParamValue('ADSK_Размер_Высота')

            revit_coords = RevitXYZmms(
                convert_to_mms(bb_center.X),
                convert_to_mms(bb_center.Y),
                convert_to_mms(bb.Min.Z - element.GetParamValue(BuiltInParameter.INSTANCE_ELEVATION_PARAM))
            )

            # if element.Id.IntegerValue == 2708661:
            #     print('{} {} {}'.format(bb_center.X, bb_center.Y, bb.Min.Z))
            #     print(elem_height)
            for auditor_data in ayditror_equipment:
                if is_within_sphere(auditor_data, revit_coords):
                    family_name = element.Symbol.Family.Name
                    if 'Обр_ОП_Универсальный' in family_name:
                        auditor_data.processed = True
                        insert_data(element, auditor_data)
                        break  # Переходим к следующему элементу в equipment
            else:
                not_found_revit.append(element.Id.IntegerValue)

        for ayditror_data in ayditror_equipment:
            if not ayditror_data.processed:
                not_found_ayditor.append(ayditror_data)


    # if len(not_found_revit) > 0:
    #     print('ID Элементов Revit, которые не были обработаны:')
    #     print(not_found_revit)
    #
    # if len(not_found_ayditor) > 0:
    #     print('Оборудование аудитор, которое не было найдено в модели:')
    #     for ayditror_data in not_found_ayditor:
    #         print('Прибор х: {}, y: {}, z: {}, артикул: {}, мощность: {}'.format(
    #             ayditror_data.x,
    #             auditor_data.y,
    #             ayditror_data.z,
    #             auditor_data.code,
    #             auditor_data.real_power))


script_execute()
