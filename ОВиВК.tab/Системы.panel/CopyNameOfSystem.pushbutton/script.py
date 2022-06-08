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


def make_col(category):
    col = FilteredElementCollector(document)\
                            .OfCategory(category)\
                            .WhereElementIsElementType()\
                            .ToElements()
    return col


colPipeSystem = make_col(BuiltInCategory.OST_PipingSystem)

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


def  get_system_reduction(system_name):
	if system_name == None:
		system_reduction = 'None'
	else:
		for system in colPipeSystem:
			sys_abb = str(system.GetParamValueOrDefault(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM))

			if sys_abb in system_name:
				system_reduction = system.LookupParameter('ФОП_ВИС_Сокращение для системы').AsString()
				if system_reduction == None:
					system_reduction = 'None'

	return system_reduction


sub_element_ids = []
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

	system_reduction = get_system_reduction(system_name)

	if system_reduction != "None":
		system_name = system_reduction



	if hasattr(element, "GetSubComponentIds"):
		sub_elements = [document.GetElement(element_id) for element_id in element.GetSubComponentIds()]
		for sub_element in sub_elements:
			sub_element.SetParamValue(SharedParamsConfig.Instance.MechanicalSystemName, str(system_name))
			sub_element_ids.append(sub_element.Id)

	if element.Id not in sub_element_ids:
		element.SetParamValue(SharedParamsConfig.Instance.MechanicalSystemName, str(system_name))

def update_element(elements):
	report_rows = set()
	for element in elements:
		try:
			edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
			if edited_by and edited_by != __revit__.Application.Username:
				report_rows.add(edited_by)
				continue

			update_system_name(element)

		except: # надеюсь это маловероятный сценарий
			pass

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