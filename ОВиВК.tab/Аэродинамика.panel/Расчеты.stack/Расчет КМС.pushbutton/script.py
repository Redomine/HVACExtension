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
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB.ExternalService import *
from Autodesk.Revit.DB.ExtensibleStorage import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, Selection
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit
from collections import namedtuple

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig

class CalculationMethod:
    name = None
    server_id = None
    server = None

    def __init__(self, name, server, server_id):
        self.name = name
        self.server = server
        self.server_id = server_id

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

    return duct_elements

def get_loss_methods():
    service_id = ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService

    service = ExternalServiceRegistry.GetService(service_id)
    server_ids = service.GetRegisteredServerIds()

    for server_id in server_ids:
        server = get_server_by_id(server_id, service_id)
        name = server.GetName()

        if str(server_id) == "5a598293-1504-46cc-a9c0-de55c82848b9": # Это - Гуид "Определенный коэффициент". Вроде бы одинаков всегда
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

    method = get_loss_methods()

    param.Set(method.server_id.ToString())

    if value != 0:
        schema = method.server.GetDataSchema()
        entity = element.GetEntity(schema)
        coefficient_field = schema.GetField("Coefficient")
        entity.Set[coefficient_field.ValueType](coefficient_field, str(value))
        element.SetEntity(entity)

def split_elements(system_elements):
    fittings = []
    accessories = []
    for element in system_elements:
        if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
            fittings.append(element)
        if element.Category.IsId(BuiltInCategory.OST_DuctAccessory):
            accessories.append(element)
    return fittings, accessories

def get_local_coefficient(fitting):
    local_section_coefficient = 0

    if str(fitting.MEPModel.PartType) == 'Elbow':
        local_section_coefficient = calculator.get_coef_elbow(fitting)

    if str(fitting.MEPModel.PartType) == 'Transition':
        local_section_coefficient = calculator.get_coef_transition(fitting)

    if str(fitting.MEPModel.PartType) == 'Tee':
        local_section_coefficient = calculator.get_coef_tee(fitting)

    if str(fitting.MEPModel.PartType) == 'TapAdjustable':
        local_section_coefficient = calculator.get_coef_tap_adjustable(fitting)

    return local_section_coefficient

doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

calculator = CoefficientCalculator.Aerodinamiccoefficientcalculator(doc, uidoc, view)

def script_execute():
    with revit.Transaction("BIM: Пересчет потерь напора"):
        system_elements = get_system_elements()

        if system_elements is None:
            forms.alert(
                "Не найдены элементы в системе.",
                "Ошибка",
                exitscript=True)

        fittings, accessories = split_elements(system_elements)

        for fitting in fittings:
            local_section_coefficient = get_local_coefficient(fitting)
            set_method(fitting, local_section_coefficient)

        for accessory in accessories:
            set_method(accessory)

    # with revit.Transaction("Выключение систем"):
    #     col_systems = make_col(BuiltInCategory.OST_DuctSystem)
    #     for el in col_systems:
    #         sys_type = doc.GetElement(el.GetTypeId())
    #         sys_type.CalculationLevel = sys_type.CalculationLevel.None
    #
    # with revit.Transaction("Включение систем"):
    #     col_systems = make_col(BuiltInCategory.OST_DuctSystem)
    #     for el in col_systems:
    #         sys_type = doc.GetElement(el.GetTypeId())
    #         sys_type.CalculationLevel = sys_type.CalculationLevel.All

script_execute()
