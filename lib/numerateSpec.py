# -*- coding: utf-8 -*-
import math

import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from dosymep.Bim4Everyone.SharedParams import *
from dosymep.Bim4Everyone import ElementExtensions
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep_libs.bim4everyone import *
from dosymep.Revit import *

doc = __revit__.ActiveUIDocument.Document  # type: Document
active_view = doc.ActiveView

class SpecificationSettings:
    definition = None
    rollback_itemized = False
    rollback_header = False
    hidden = []
    sort_para_group_indexes = None
    position_index = None
    group_index = None

    def __init__(self, definition):
        self.sort_para_group_indexes, self.position_index, self.group_index = get_sort_rules(definition)


        self.definition = definition
        # если заголовки показаны изначально или если спека изначально развернута - сворачивать назад не нужно
        self.rollback_itemized = not definition.IsItemized
        self.rollback_header = not definition.ShowHeaders

        # Разворачиваем спецификацию полностью
        definition.IsItemized = True
        definition.ShowHeaders = True

        # Собираем номера скрытых столбцов и разворачиваем их
        i = 0
        while i < definition.GetFieldCount():
            if definition.GetField(i).IsHidden == True:
                self.hidden.append(i)
            definition.GetField(i).IsHidden = False
            i += 1

    def repair_specification(self):
        self.definition.IsItemized = not self.rollback_itemized
        self.definition.ShowHeaders = not self.rollback_header

        i = 0
        while i < self.definition.GetFieldCount():
            if i in self.hidden:
                self.definition.GetField(i).IsHidden = True
            i += 1

def get_sort_rules(definition):
    sortGroupInd = []
    posInShed = False
    groupingInd = False
    index = 0

    for scheduleGroupField in definition.GetFieldOrder():
        scheduleField = definition.GetField(scheduleGroupField)
        if scheduleField.GetName() == "ФОП_ВИС_Позиция":
            posInShed = True
            FOP_pos_ind = index
        if scheduleField.GetName() == "ФОП_ВИС_Группирование":
            groupingInd = index

        index += 1

    index = 0
    for field in definition.GetFieldOrder():
        for scheduleSortGroupField in definition.GetSortGroupFields():
            if scheduleSortGroupField.FieldId.ToString() == field.ToString():
                sortGroupInd.append(index)

        index += 1

    if posInShed == False or groupingInd == False:

        forms.alert('С добавленными параметрами "ФОП_ВИС_Позиция", "ФОП_ВИС_Группирование" и "ФОП_ВИС_Примечание"',
                    "Ошибка", exitscript=True)
    return [sortGroupInd, FOP_pos_ind, groupingInd]

# возвращает значение данных по которым идет сортировка слитыми в единую строку
def get_sort_rule_string(row, specification_settings, vs):
    return ''.join(vs.GetCellText(SectionType.Body, row, ind) for ind in specification_settings.sort_para_group_indexes)

def process_row(row, specification_settings,old_schedule_string, position_number):
    new_sort_rule = get_sort_rule_string(row, specification_settings, active_view)
    element_id = active_view.GetCellText(SectionType.Body, row, specification_settings.position_index)
    if element_id.isdigit():
        group = active_view.GetCellText(SectionType.Body, row, specification_settings.group_index)
        element = doc.GetElement(ElementId(int(element_id)))
        if new_sort_rule != old_schedule_string:
            position_number += 1

        if '_Узел_' not in group:
            element.SetParamValue(SharedParamsConfig.Instance.VISPosition, str(position_number))
        else:
            element.SetParamValue(SharedParamsConfig.Instance.VISPosition, '')

    return new_sort_rule, position_number

def check_edited_elements(elements):
    # Возвращает имя занявшего элемент или None
    def get_element_editor_name(element):
        user_name = __revit__.Application.Username
        edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
        if edited_by is None:
            return None

        if edited_by.lower() == user_name.lower():
            return None
        return edited_by

    report_rows = []

    # если что-то из элементов занято, то дальнейшая обработка не имеет смысла, нужно освобождать спеку
    for element in elements:
        name = get_element_editor_name(element)
        if name is not None:
            report_rows.append(name)

    if len(report_rows) > 0:
        report_message = (("Нумерация/заполнение примечаний не были выполнены, "
                          "так как часть элементов спецификации занята пользователями: {}")
                          .format(", ".join(report_rows)))
        forms.alert(report_message, "Ошибка", exitscript=True)

def process_areas():
    ducts = FilteredElementCollector(doc, active_view.Id).OfCategory(BuiltInCategory.OST_DuctCurves).ToElements()

    # Создаем словарь для хранения воздуховодов по их позициям
    duct_dict = {}

    # Перебираем все воздуховоды и добавляем их в словарь по признаку одинаковой позиции
    for duct in ducts:
        duct_position = duct.GetParamValue(SharedParamsConfig.Instance.VISPosition)
        duct_area = UnitUtils.ConvertFromInternalUnits(
            duct.GetParamValueOrDefault(BuiltInParameter.RBS_CURVE_SURFACE_AREA),
            UnitTypeId.SquareMeters)

        if duct_position in duct_dict:
            duct_dict[duct_position] += duct_area
        else:
            duct_dict[duct_position] = duct_area

    # Обновляем параметр VISNote для каждого воздуховода
    for duct in ducts:
        duct_position = duct.GetParamValue(SharedParamsConfig.Instance.VISPosition)
        formatted_area = "{:.2f}".format(duct_dict[duct_position]).rstrip('0').rstrip('.') + ' м²'
        duct.SetParamValue(SharedParamsConfig.Instance.VISNote, formatted_area)

def numerate(doNumbers, doAreas):
    if active_view.Category is None or not active_view.Category.IsId(BuiltInCategory.OST_Schedules):
        forms.alert("Нумерация и вынесение площади воздуховодов сработают только на активном виде целевой спецификации",
                    "Ошибка", exitscript=True)

    definition = active_view.Definition
    section_data = active_view.GetTableData().GetSectionData(SectionType.Body)
    elements = FilteredElementCollector(doc, doc.ActiveView.Id)
    #если что-то из элементов занято, то дальнейшая обработка не имеет смысла, нужно освобождать спеку
    check_edited_elements(elements)

    with revit.Transaction("Запись айди"):
        specification_settings = SpecificationSettings(definition)

        for element in elements:
            element.SetParamValue(SharedParamsConfig.Instance.VISPosition, str(element.Id.IntegerValue))

    position_number = 0 # Стартовый значение для номера
    old_sort_rule = '' # Стартовая пустая строка сортировки
    with revit.Transaction("Запись номера"):
        row = section_data.FirstRowNumber
        while row <= section_data.LastRowNumber:
            old_sort_rule, position_number = process_row(row,
                                                            specification_settings,
                                                            old_sort_rule,
                                                            position_number)
            row += 1

        specification_settings.repair_specification()

        if doAreas is True:
            process_areas()

        if doNumbers is False:
            for element in elements:
                element.SetParamValue(SharedParamsConfig.Instance.VISPosition, '')

