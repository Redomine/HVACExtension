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
        self.definition = definition

        self.position_index = self.__get_schedule_parameter_index(SharedParamsConfig.Instance.VISPosition.Name)

        self.group_index = self.__get_schedule_parameter_index(SharedParamsConfig.Instance.VISGrouping.Name)

        self.sort_para_group_indexes = self.__get_sorting_params_indexes()

        # если заголовки показаны изначально или если спека изначально развернута - сворачивать назад не нужно
        self.rollback_itemized = not definition.IsItemized
        self.rollback_header = not definition.ShowHeaders

    def show_all_specification(self):
        # Разворачиваем спецификацию полностью
        self.definition.IsItemized = True
        self.definition.ShowHeaders = True

        # Собираем номера скрытых столбцов и разворачиваем их
        i = 0
        while i < self.definition.GetFieldCount():
            if self.definition.GetField(i).IsHidden == True:
                self.hidden.append(i)
            self.definition.GetField(i).IsHidden = False
            i += 1

    def repair_specification(self):
        self.definition.IsItemized = not self.rollback_itemized
        self.definition.ShowHeaders = not self.rollback_header

        i = 0
        while i < self.definition.GetFieldCount():
            if i in self.hidden:
                self.definition.GetField(i).IsHidden = True
            i += 1

    def __get_sorting_params_indexes(self):
        sorting_params_indexes = []
        index = 0
        for field in self.definition.GetFieldOrder():
            for scheduleSortGroupField in self.definition.GetSortGroupFields():
                if scheduleSortGroupField.FieldId.ToString() == field.ToString():
                    sorting_params_indexes.append(index)

            index += 1

        return sorting_params_indexes

    def __get_schedule_parameter_index(self, name):
        index = 0

        for schedule_parameter in self.definition.GetFieldOrder():
            schedule_parameter = self.definition.GetField(schedule_parameter)
            if schedule_parameter.GetName() == name:
                #print doc.GetElement(schedule_parameter.ParameterId).IsInstance
                return index
            index += 1

        forms.alert('В таблице нет параметра {}.'.format(name),
                    "Ошибка", exitscript=True)

class SpecificationFiller:
    doc = None
    active_view = None

    def __init__(self, doc, active_view):
        self.doc = doc
        self.active_view = active_view

    # возвращает значение данных по которым идет сортировка слитыми в единую строку
    def __get_sort_rule_string(self, row, specification_settings, vs):
        return ''.join(vs.GetCellText(SectionType.Body, row, ind) for ind in specification_settings.sort_para_group_indexes)

    def __process_row(self, row, specification_settings, old_schedule_string, position_number):
        new_sort_rule = self.__get_sort_rule_string(row, specification_settings, self.active_view)
        element_id = self.active_view.GetCellText(SectionType.Body, row, specification_settings.position_index)
        if element_id.isdigit():
            group = self.active_view.GetCellText(SectionType.Body, row, specification_settings.group_index)
            element = self.doc.GetElement(ElementId(int(element_id)))
            if new_sort_rule != old_schedule_string:
                position_number += 1

            if '_Узел_' not in group:
                element.SetParamValue(SharedParamsConfig.Instance.VISPosition, str(position_number))
            else:
                element.SetParamValue(SharedParamsConfig.Instance.VISPosition, '')

        return new_sort_rule, position_number

    def __check_edited_elements(self, elements):
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

    def __process_areas(self, ):
        ducts = FilteredElementCollector(
            self.doc,
            self.active_view.Id).OfCategory(BuiltInCategory.OST_DuctCurves).ToElements()

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

    def __fill_id_to_schedule_param(self, specification_settings, elements):
        with revit.Transaction("Запись айди"):
            specification_settings.show_all_specification()

            for element in elements:
                element.SetParamValue(SharedParamsConfig.Instance.VISPosition, str(element.Id.IntegerValue))

    def __fill_position_and_note(self, specification_settings, elements, fill_areas, fill_numbers):
        with revit.Transaction("Запись номера"):
            section_data = self.active_view.GetTableData().GetSectionData(SectionType.Body)
            row = section_data.FirstRowNumber

            position_number = 0  # Стартовый значение для номера
            old_sort_rule = ''  # Стартовая пустая строка сортировки
            while row <= section_data.LastRowNumber:
                old_sort_rule, position_number = self.__process_row(row,
                                                                    specification_settings,
                                                                    old_sort_rule,
                                                                    position_number)
                row += 1

            specification_settings.repair_specification()

            if fill_areas is True:
                self.__process_areas()

            if fill_numbers is False:
                for element in elements:
                    element.SetParamValue(SharedParamsConfig.Instance.VISPosition, '')

    def __check_position_param(self, elements):
        for element in elements:
            if not element.IsExistsParam(SharedParamsConfig.Instance.VISPosition):
                forms.alert(
                    'Параметр "ФОП_ВИС_Позиция" для некоторых элементов спецификации является параметром типа. '
                    'Нумерация этих элементов будет пропущена.',
                    "Внимание")
                return

    def __check_duct_note_param(self, elements):
        for element in elements:
            if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and not element.IsExistsParam(SharedParamsConfig.Instance.VISNote):
                forms.alert(
                    'Параметр "ФОП_ВИС_Примечание" является параметром типа воздуховодов. '
                    'Примечания не будут заполнены',
                    "Внимание")
                return

    def numerate(self, fill_numbers = False, fill_areas = False):
        if self.active_view.Category is None or not active_view.Category.IsId(BuiltInCategory.OST_Schedules):
            forms.alert("Нумерация и вынесение площади воздуховодов сработают только на активном виде целевой спецификации",
                        "Ошибка", exitscript=True)

        elements = FilteredElementCollector(self.doc, self.doc.ActiveView.Id)

        # Проверяем параметры
        self.__check_position_param(elements)
        self.__check_duct_note_param(elements)

        #если что-то из элементов занято, то дальнейшая обработка не имеет смысла, нужно освобождать спеку
        self.__check_edited_elements(elements)

        # собираем информацию по работе активного вида
        specification_settings = SpecificationSettings(self.active_view.Definition)

        # Заполняем в параметр позиции айди элементов
        self.__fill_id_to_schedule_param(specification_settings, elements)

        # Заполняем позицию и примечания
        self.__fill_position_and_note(specification_settings, elements, fill_areas, fill_numbers)

