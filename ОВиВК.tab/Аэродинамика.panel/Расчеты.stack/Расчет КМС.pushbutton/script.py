#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пересчет КМС'
__doc__ = "Пересчитывает КМС соединительных деталей воздуховодов"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import sys
import System
import math
import CoefficientCalculator
from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.DB.Mechanical import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, Selection
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from collections import namedtuple

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep_libs.bim4everyone import *

class CalculationMethod:
    name = None
    server_id = None
    server = None

    def __init__(self, name, server, server_id):
        self.name = name
        self.server = server
        self.server_id = server_id

class EditorReport:
    edited_reports = []
    status_report = ''
    edited_report = ''

    def __get_element_editor_name(self, element):
        """
        Возвращает имя пользователя, занявшего элемент, или None.

        Args:
            element (Element): Элемент для проверки.

        Returns:
            str или None: Имя пользователя или None, если элемент не занят.
        """
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None

        if edited_by.lower() in user_name.lower():
            return None
        return edited_by

    def is_element_edited(self, element):
        """
        Проверяет, заняты ли элементы другими пользователями.

        Args:
            element: Элемент для проверки.
        """

        self.update_status = WorksharingUtils.GetModelUpdatesStatus(doc, element.Id)

        if self.update_status == ModelUpdatesStatus.UpdatedInCentral:
            self.status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию. "

        name = self.__get_element_editor_name(element)
        if name is not None and name not in edited_reports:
            self.edited_reports.append(name)
            return True

    def show_report(self):
        if len(self.edited_reports) > 0:
            self.edited_report = ("Часть элементов спецификации занята пользователями: {}"
                                  .format(", ".join(self.edited_reports)))
        if self.edited_report != '' or self.status_report != '':
            report_message = status_report + ('\n' if (edited_report and status_report) else '') + self.edited_report
            forms.alert(report_message, "Ошибка", exitscript=True)

class SelectedSystem:
    name = None
    elements = None
    system = None

    def __init__(self, name, elements, system):
        self.name = name
        self.system = system
        self.elements = elements

def get_system_elements():
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count != 1:
        forms.alert(
            "Должна быть выделена одна система воздуховодов.",
            "Ошибка",
            exitscript=True)

    system = doc.GetElement(selected_ids[0])

    if system.Category.IsId(BuiltInCategory.OST_DuctSystem) == False:
        forms.alert(
            "Должна быть выделена одна система воздуховодов.",
            "Ошибка",
            exitscript=True)

    duct_elements = system.DuctNetwork
    system_name = system.GetParamValue(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)

    selected_system = SelectedSystem(system_name, duct_elements, system)

    return selected_system

def get_loss_methods():
    service_id = ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService

    service = ExternalServiceRegistry.GetService(service_id)
    server_ids = service.GetRegisteredServerIds()

    for server_id in server_ids:
        server = get_server_by_id(server_id, service_id)
        name = server.GetName()
        if str(server_id) == calculator.COEFF_GUID_CONST:
            calculation_method = CalculationMethod(name, server, server_id)
            return calculation_method

def get_server_by_id(server_guid, service_id):
    service = ExternalServiceRegistry.GetService(service_id)
    if service is not None and server_guid is not None:
        server = service.GetServer(server_guid)
        if server is not None:
            return server
    return None

def set_method(element, value = 0):
    param = element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)
    current_guid = param.AsString()

    method = get_loss_methods()

    if current_guid != calculator.LOSS_GUID_CONST and current_guid != calculator.COEFF_GUID_CONST:
        param.Set(method.server_id.ToString())

    if value != 0 and current_guid != calculator.LOSS_GUID_CONST:
        schema = method.server.GetDataSchema()
        entity = element.GetEntity(schema)
        coefficient_field = schema.GetField("Coefficient")
        entity.Set[coefficient_field.ValueType](coefficient_field, str(value))
        element.SetEntity(entity)

def split_elements(system_elements):
    fittings = []
    accessories = []
    for element in system_elements:
        editor_report.is_element_edited(element)
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            fittings.append(element)
        if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            accessories.append(element)
    return fittings, accessories

def get_local_coefficient(fitting):
    part_type = fitting.MEPModel.PartType

    if part_type == fitting.MEPModel.PartType.Elbow:
        local_section_coefficient = calculator.get_coef_elbow(fitting)
    elif part_type == fitting.MEPModel.PartType.Transition:
        local_section_coefficient = calculator.get_coef_transition(fitting)
    elif part_type == fitting.MEPModel.PartType.Tee:
        local_section_coefficient = calculator.get_coef_tee(fitting)
    elif part_type == fitting.MEPModel.PartType.TapAdjustable:
        local_section_coefficient = calculator.get_coef_tap_adjustable(fitting)
    else:
        local_section_coefficient = 0

    return local_section_coefficient

def get_network_element_name(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        name = 'Воздуховод'
    elif element.Category.IsId(BuiltInCategory.OST_DuctTerminal):
        name = 'Воздухораспределитель'
    elif element.Category.IsId(BuiltInCategory.OST_MechanicalEquipment):
        name = 'Оборудование'
    elif element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        name = 'Фасонный элемент воздуховода'
        if str(element.MEPModel.PartType) == 'Elbow':
            name = 'Отвод воздуховода'
        if str(element.MEPModel.PartType) == 'Transition':
            name = 'Переход между сечениями'
        if str(element.MEPModel.PartType) == 'Tee':
            name = 'Тройник'
        if str(element.MEPModel.PartType) == 'TapAdjustable':
            name = 'Врезка'
    else:
        name = 'Арматура'

    return name

def get_network_element_length(section, element_id):
    length = '-'
    try:
        length = section.GetSegmentLength(element_id) * 304.8 / 1000
        length = float('{:.2f}'.format(length))
    except Exception:
        pass
    return length

def get_network_element_coefficient(section, element):
    coefficient = element.GetParamValueOrDefault('ФОП_ВИС_КМС')

    if coefficient is None:
        coefficient = section.GetCoefficient(element.Id)

    return coefficient

def get_network_element_real_size(element, element_type):
    size = element.GetParamValueOrDefault('ФОП_ВИС_Живое сечение, м2', 0)
    if size == 0:
        element_type.GetParamValueOrDefault('ФОП_ВИС_Живое сечение, м2', 0)

    if size == 0:
        size = element.GetParamValueOrDefault(BuiltInParameter.RBS_CALCULATED_SIZE, "-")

    return size

def get_network_element_pressure_drop(section, element, size):
    pressure_drop = 1

    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        pressure_drop = section.GetPressureDrop(element.Id) * 3.280839895

    if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
        pass

    return pressure_drop

doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

calculator = CoefficientCalculator.Aerodinamiccoefficientcalculator(doc, uidoc, view)
editor_report = EditorReport()

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with revit.Transaction("BIM: Пересчет потерь напора"):
        selected_system = get_system_elements()

        if selected_system.elements is None:
            forms.alert(
                "Не найдены элементы в системе.",
                "Ошибка",
                exitscript=True)

        fittings, accessories = split_elements(selected_system.elements)

        for fitting in fittings:
            local_section_coefficient = get_local_coefficient(fitting)
            set_method(fitting, local_section_coefficient)

        for accessory in accessories:
            set_method(accessory)

        editor_report.show_report()

    with revit.Transaction("BIM: Вывод отчета"):
        # заново забираем систему  через ID, мы в прошлой транзакции обновили потери напора на элементах, поэтому данные
        # на системе могли измениться
        system = doc.GetElement(selected_system.system.Id)
        path_numbers = system.GetCriticalPathSectionNumbers()

        path = []
        for number in path_numbers:
            path.append(number)
        if system.SystemType == DuctSystemType.SupplyAir:
            path.reverse()

        data = []
        count = 0
        passed_taps = []
        output = script.get_output()
        flow = 0
        velocity = 0
        old_flow = 0


        settings = DuctSettings.GetDuctSettings(doc)
        density = settings.AirDensity * 35.3146667215
        print 'Плотность воздушной среды: ' + str(density) + ' кг/м3'

        pressure_total = 0
        for number in path:
            section = system.GetSectionByNumber(number)
            elements_ids = section.GetElementIds()
            for element_id in elements_ids:
                element = doc.GetElement(element_id)
                element_type = element.GetElementType()

                name = get_network_element_name(element)

                size = get_network_element_real_size(element, element_type)

                length = get_network_element_length(section, element_id)

                coefficient = get_network_element_coefficient(section, element)

                real_size = get_network_element_real_size(element, element_type)

                pressure_drop = get_network_element_pressure_drop(section, element, size)

                #pressure_drop = float('{:.2f}'.format(pressure_drop))

                pressure_total += pressure_drop

                if pressure_drop == 0:
                    continue
                else:
                    data.append([
                        count,
                        name,
                        length,
                        real_size,
                        flow,
                        velocity,
                        coefficient,
                        pressure_drop,
                        pressure_total,
                        output.linkify(element_id)])


    output.print_table(table_data=data,
                       title=("Отчет о расчете аэродинамики системы " + selected_system.name),
                       columns=[
                           "Номер участка",
                           "Наименование элемента",
                           "Длина, м.п.",
                           "Размер",
                           "Расход, м3/ч",
                           "Скорость, м/с",
                           "КМС",
                           "Потери напора элемента, Па",
                           "Суммарные потери напора, Па",
                           "Id элемента"],
                       formats=['', '', ''],
                       )


script_execute()
