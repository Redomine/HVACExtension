# -*- coding: utf-8 -*-
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from System.Collections.Generic import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import *

from pyrevit.script import output

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from System import Guid

document = __revit__.ActiveUIDocument.Document
sub_element_ids = []


def make_col(category):
    col = FilteredElementCollector(document)\
                            .OfCategory(category)\
                            .WhereElementIsElementType()\
                            .ToElements()
    return col

def make_eks_col(category):
    col = FilteredElementCollector(document)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

colPipeSystem = make_col(BuiltInCategory.OST_PipingSystem)

colDuctSystem = make_col(BuiltInCategory.OST_DuctSystem)
colInsulation = make_eks_col(BuiltInCategory.OST_PipeInsulations)



def get_elements():
	categories = [BuiltInCategory.OST_MechanicalEquipment, 	#Оборудование

				  BuiltInCategory.OST_PlumbingFixtures, 	#Сантехнические приборы
				  BuiltInCategory.OST_Sprinklers,			#Спринклеры
				  BuiltInCategory.OST_PipeFitting,			#Соединительные детали трубопроводов
				  BuiltInCategory.OST_PipeAccessory,		#Арматура трубопроводов
				  BuiltInCategory.OST_PipeInsulations,		#Материалы изоляции труб
				  BuiltInCategory.OST_FlexPipeCurves,		#Гибкие трубы"
				  BuiltInCategory.OST_PipeCurves,			#Трубы

				  BuiltInCategory.OST_DuctCurves,			#Воздуховоды
				  BuiltInCategory.OST_DuctFitting, 			#Соединительные детали воздуховодов
				  BuiltInCategory.OST_DuctAccessory, 		#Арматрура воздуховодов
				  BuiltInCategory.OST_DuctInsulations,		#Материалы изоляции воздуховодов
				  BuiltInCategory.OST_FlexDuctCurves, 		#Гибкие воздуховоды
				  BuiltInCategory.OST_DuctTerminal]			#Воздухораспределители

	category_filter = ElementMulticategoryFilter(List[BuiltInCategory](categories))
	return FilteredElementCollector(document).WherePasses(category_filter).WhereElementIsNotElementType().ToElements()


def get_connectors(element):
	if hasattr(element, "ConnectorManager"):
		if element.ConnectorManager:
			return element.ConnectorManager.Connectors

	if hasattr(element, "MEPModel") and element.MEPModel:
		if element.MEPModel.ConnectorManager:
			return element.MEPModel.ConnectorManager.Connectors

	if hasattr(element, "MEPSystem") and element.MEPSystem:
		if element.MEPSystem.ConnectorManager:
			return element.MEPSystem.ConnectorManager.Connectors

	return []


def get_type_system(element):

	connector = next((c for c in get_connectors(element) if c.Domain == Domain.DomainHvac or c.Domain == Domain.DomainPiping), None)
	if connector:
		return get_type_system_name(connector)


def get_type_system_name(element):
	if hasattr(element, "HostElementId"):
		mep_type = get_type_system_name(document.GetElement(element.HostElementId))
		return mep_type.GetParamValueOrDefault("ФОП_ВИС_Сокращение для системы")
	if hasattr(element, "MEPSystem") and element.MEPSystem:
		mep_type = document.GetElement(element.MEPSystem.GetTypeId())
		return mep_type.GetParamValueOrDefault("ФОП_ВИС_Сокращение для системы")


def update_system_name(element):
	if element.GetParam(SharedParamsConfig.Instance.MechanicalSystemName).IsReadOnly:
		return

	system_name = element.GetParamValueOrDefault(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
	if not system_name:
		super_component = element.SuperComponent
		if super_component:
			system_name = super_component.GetParamValueOrDefault(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)

	if system_name:
		# Т11 3,Т11 4 -> Т11
		# Т11 3,Т12 4 -> Т11, Т12
		system_name = ", ".join(set([s.split(" ")[0] for s in system_name.split(",")]))

	type_system_name = get_type_system(element)



	if type_system_name:
		system_name = type_system_name



	if hasattr(element, "GetSubComponentIds"):
		sub_elements = [document.GetElement(element_id) for element_id in element.GetSubComponentIds()]

		for sub_element in sub_elements:

			sub_element.SetParamValue(SharedParamsConfig.Instance.MechanicalSystemName, str(system_name))
			sub_element_ids.append(sub_element.Id)

	if element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
		system_name = get_type_system_name(document.GetElement(element.HostElementId))



	if element.Id not in sub_element_ids:
		element.SetParamValue(SharedParamsConfig.Instance.MechanicalSystemName, str(system_name))


def update_element(elements):
	report_rows = set()
	for element in elements:
		edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
		if edited_by and edited_by != __revit__.Application.Username:
			report_rows.add(edited_by)
			continue

		update_system_name(element)

	return report_rows


def script_execute():
	# настройка атрибутов
	project_parameters = ProjectParameters.Create(__revit__.Application)
	project_parameters.SetupRevitParams(document, SharedParamsConfig.Instance.MechanicalSystemName)

	with Transaction(document) as transaction:
		transaction.Start("Обновление атрибута \"{}\"".format(SharedParamsConfig.Instance.MechanicalSystemName.Name))

		elements = get_elements()
		report_rows = update_element(elements)
		if report_rows:
			output1 = output.get_output()
			output1.set_title("Обновление атрибута \"{}\"".format(SharedParamsConfig.Instance.MechanicalSystemName.Name))

			print "Некоторые элементы не были обработаны, так как были заняты пользователями:"
			print "\r\n".join(report_rows)

		transaction.Commit()


script_execute()