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
import paraSpec

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
	connectors = get_connectors(element)
	hvac_connector = None

	for connector in connectors:
		if connector.Domain == Domain.DomainHvac or connector.Domain == Domain.DomainPiping:
			hvac_connector = connector
			if hvac_connector:
				if get_type_system_name(hvac_connector) != None:
					return get_type_system_name(hvac_connector)

	return None


def get_type_system_name(element):
	if hasattr(element, "HostElementId"):
		mep_type = get_type_system_name(document.GetElement(element.HostElementId))
		return mep_type.GetParamValueOrDefault("ФОП_ВИС_Сокращение для системы")
	if hasattr(element, "MEPSystem") and element.MEPSystem:
		mep_type = document.GetElement(element.MEPSystem.GetTypeId())
		return mep_type.GetParamValueOrDefault("ФОП_ВИС_Сокращение для системы")



def rename_sub(element):
	if hasattr(element, "GetSubComponentIds"):
		super_component = element.SuperComponent
		if super_component:
			return

		system_name = element.GetParamValueOrDefault("ADSK_Имя системы")

		sub_elements = [document.GetElement(element_id) for element_id in element.GetSubComponentIds()]
		for sub_element in sub_elements:

			sub_element.SetParamValue(SharedParamsConfig.Instance.MechanicalSystemName, str(system_name))
			rename_sub_sub(sub_element, system_name)
			sub_element_ids.append(sub_element.Id)

def rename_sub_sub(element, system_name):
	if not hasattr(element, "GetSubComponentIds"):
		return
	for element_id in element.GetSubComponentIds():
		element = document.GetElement(element_id)
		element.SetParamValue(SharedParamsConfig.Instance.MechanicalSystemName, str(system_name))
		rename_sub_sub(element, system_name)


def update_system_name(element):
	if element.GetParam(SharedParamsConfig.Instance.MechanicalSystemName).IsReadOnly:
		return

	system_name = element.GetParamValueOrDefault(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)

	if not system_name:
		try:
			super_component = element.SuperComponent
			if super_component:
				system_name = super_component.GetParamValueOrDefault(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
		except Exception:
			pass

	if system_name:
		# Т11 3,Т11 4 -> Т11
		# Т11 3,Т12 4 -> Т11, Т12
		system_name = ", ".join(set([s.split(" ")[0] for s in system_name.split(",")]))



	type_system_name = get_type_system(element)

	if element.Category.IsId(BuiltInCategory.OST_PipeInsulations):
		type_system_name = get_type_system_name(document.GetElement(element.HostElementId))

	if type_system_name:
		system_name = type_system_name

	if element.Id.IntegerValue == 2517492:
		print system_name

	if element.Id.IntegerValue == 2517492:
		print type_system_name


	if system_name == None:
		if document.ProjectInformation.GetParamValueOrDefault('ФОП_ВИС_Имя внесистемных элементов') != None:
			system_name = document.ProjectInformation.GetParamValueOrDefault('ФОП_ВИС_Имя внесистемных элементов')

	element.SetParamValue(SharedParamsConfig.Instance.MechanicalSystemName, str(system_name))
	element.LookupParameter('ФОП_ВИС_Имя системы').Set(str(system_name))



def update_element(elements):
	report_rows = set()
	for element in elements:
		edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
		if edited_by and edited_by != __revit__.Application.Username:
			report_rows.add(edited_by)
			continue
		update_system_name(element)

	#отдельно проходимся по суб-элементам, чтоб не перекрыть имена случайно основной проверкой
	for element in elements:
		edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
		if edited_by and edited_by != __revit__.Application.Username:
			continue

		try:
			rename_sub(element)
			name = element.LookupParameter('ADSK_Имя системы').AsString()
			element.LookupParameter('ФОП_ВИС_Имя системы').Set(name)
		except Exception:
			print element.Id

	return report_rows


def script_execute():
	# настройка атрибутов
	project_parameters = ProjectParameters.Create(__revit__.Application)
	project_parameters.SetupRevitParams(document, SharedParamsConfig.Instance.MechanicalSystemName,
										SharedParamsConfig.Instance.VISSystemShortName,
										SharedParamsConfig.Instance.VISOutSystemName
										)

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

status = paraSpec.check_parameters()
if not status:
    script_execute()
