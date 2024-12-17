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

class SpecificationSettings:
    """
    Класс для настройки спецификаций в документе.

    Атрибуты:
        definition (ScheduleDefinition): Определение спецификации.
        rollback_itemized (bool): Флаг для возврата состояния развернутости спецификации.
        rollback_header (bool): Флаг для возврата состояния отображения заголовков.
        hidden (list): Список скрытых столбцов.
        sort_para_group_indexes (list): Индексы параметров сортировки.
        position_index (int): Индекс параметра позиции.
        group_index (int): Индекс параметра группы.
    """

    definition = None
    rollback_itemized = False
    rollback_header = False
    hidden = []
    sort_para_group_indexes = None
    position_index = None
    group_index = None
    elements = []
    ducts = []
    duct_fittings = []

    def __init__(self, definition):
        """
        Инициализация экземпляра класса SpecificationSettings.

        Аргументы:
            definition (ScheduleDefinition): Определение спецификации.
        """
        self.definition = definition

        self.position_index = self.__get_schedule_parameter_index(SharedParamsConfig.Instance.VISPosition.Name)

        self.group_index = self.__get_schedule_parameter_index(SharedParamsConfig.Instance.VISGrouping.Name)

        self.sort_para_group_indexes = self.__get_sorting_params_indexes()

        # Если заголовки показаны изначально или если спека изначально развернута - сворачивать назад не нужно
        self.rollback_itemized = not definition.IsItemized
        self.rollback_header = not definition.ShowHeaders

    def show_all_specification(self):
        """
        Разворачивает спецификацию полностью.
        """
        self.definition.IsItemized = True
        self.definition.ShowHeaders = True

        # Собираем номера скрытых столбцов и разворачиваем их
        i = 0
        while i < self.definition.GetFieldCount():
            if self.definition.GetField(i).IsHidden:
                self.hidden.append(i)
            self.definition.GetField(i).IsHidden = False
            i += 1

    def repair_specification(self):
        """
        Возвращает спецификацию в исходное состояние.
        """
        self.definition.IsItemized = not self.rollback_itemized
        self.definition.ShowHeaders = not self.rollback_header

        i = 0
        while i < self.definition.GetFieldCount():
            if i in self.hidden:
                self.definition.GetField(i).IsHidden = True
            i += 1

    def __get_sorting_params_indexes(self):
        """
        Возвращает индексы параметров сортировки.

        Возвращает:
            list: Список индексов параметров сортировки.
        """
        sorting_params_indexes = []
        index = 0
        for field in self.definition.GetFieldOrder():
            for scheduleSortGroupField in self.definition.GetSortGroupFields():
                if scheduleSortGroupField.FieldId.ToString() == field.ToString():
                    sorting_params_indexes.append(index)

            index += 1

        return sorting_params_indexes

    def __get_schedule_parameter_index(self, name):
        """
        Возвращает индекс параметра спецификации по имени.

        Аргументы:
            name (str): Имя параметра.

        Возвращает:
            int: Индекс параметра.
        """
        index = 0

        for schedule_parameter in self.definition.GetFieldOrder():
            schedule_parameter = self.definition.GetField(schedule_parameter)
            if schedule_parameter.GetName() == name:
                return index
            index += 1

        forms.alert('В таблице нет параметра {}.'.format(name),
                    "Ошибка", exitscript=True)

class SpecificationFiller:
    """
    Класс для заполнения спецификаций в документе.

    Атрибуты:
        doc (Document): Документ, в котором происходит заполнение спецификаций.
        active_view (View): Активный вид, используемый для заполнения спецификаций.
    """

    doc = None
    active_view = None

    def __init__(self, doc, active_view):
        """
        Инициализация экземпляра класса SpecificationFiller.

        Аргументы:
            doc (Document): Документ, в котором происходит заполнение спецификаций.
            active_view (View): Активный вид, используемый для заполнения спецификаций.
        """
        self.doc = doc
        self.active_view = active_view

    def __get_sort_rule_string(self, row, specification_settings, vs):
        """
        Возвращает значение данных по которым идет сортировка слитыми в единую строку.

        Аргументы:
            row (int): Номер строки.
            specification_settings (SpecificationSettings): Настройки спецификации.
            vs (ViewSchedule): Активный вид спецификации.

        Возвращает:
            str: Строка, содержащая значения данных для сортировки.
        """
        return ''.join(vs.GetCellText(SectionType.Body, row, ind) for ind in specification_settings.sort_para_group_indexes)

    def __process_row(self, row, specification_settings, old_schedule_string, position_number):
        """
        Обрабатывает строку спецификации.

        Аргументы:
            row (int): Номер строки.
            specification_settings (SpecificationSettings): Настройки спецификации.
            old_schedule_string (str): Старая строка сортировки.
            position_number (int): Текущий номер позиции.

        Возвращает:
            tuple: Новая строка сортировки и обновленный номер позиции.
        """
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
        """
        Проверяет, заняты ли элементы другими пользователями.

        Аргументы:
            elements (list): Список элементов для проверки.
        """
        def get_element_editor_name(element):
            """
            Возвращает имя пользователя, занявшего элемент, или None.

            Аргументы:
                element (Element): Элемент для проверки.

            Возвращает:
                str или None: Имя пользователя или None, если элемент не занят.
            """
            user_name = __revit__.Application.Username
            edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
            if edited_by is None:
                return None

            if edited_by.lower() == user_name.lower():
                return None
            return edited_by

        report_rows = []

        for element in elements:
            name = get_element_editor_name(element)
            if name is not None:
                report_rows.append(name)

        if len(report_rows) > 0:
            report_message = (("Нумерация/заполнение примечаний не были выполнены, "
                              "так как часть элементов спецификации занята пользователями: {}")
                              .format(", ".join(report_rows)))
            forms.alert(report_message, "Ошибка", exitscript=True)

    def __get_fitting_area(self, element):
        area = 0

        for solid in element.GetSolid():
            for face in solid.faces:
                area += face.area

        area = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters)

        if area > 0:
            false_area = 0
            connectors = element.get_connectors()
            for connector in connectors:
                if connector.shape == ConnectorProfileType.Rectangular:
                    false_area += UnitUtils.ConvertFromInternalUnits(connector.height * connector.width, UnitTypeId.SquareMeters)
                if connector.shape == ConnectorProfileType.Round:
                    false_area += UnitUtils.ConvertFromInternalUnits(connector.radius * connector.radius * math.pi, UnitTypeId.SquareMeters)
                if connector.shape == ConnectorProfileType.Oval:
                    false_area += 0

            # Вычитаем площадь пустоты на местах коннекторов
            area -= false_area

        return area

    def __get_duct_area(self, duct):
        return UnitUtils.ConvertFromInternalUnits(
            duct.GetParamValueOrDefault(BuiltInParameter.RBS_CURVE_SURFACE_AREA),
            UnitTypeId.SquareMeters)

    def __set_area(self, elements):

    def __process_areas(self):
        """
        Обрабатывает площади воздуховодов, их фитингов, и обновляет параметр VISNote.
        """

        # Заполняем площади фитингов, если включен их учет. Иначе - металл и так идет в м2
        info = self.doc.ProjectInformation
        fill_fitting_areas = info.GetParamValueOrDefault(SharedParamsConfig.Instance.VISConsiderDuctFittings) == 1

        area_elements = FilteredElementCollector(
            self.doc,
            self.active_view.Id).OfCategory(BuiltInCategory.OST_DuctCurves).ToElements()

        if fill_fitting_areas:
            duct_fittings = FilteredElementCollector(
            self.doc,
            self.active_view.Id).OfCategory(BuiltInCategory.OST_DuctFitting).ToElements()
            area_elements = area_elements + duct_fittings


        duct_dict = {}

        for area_element in area_elements:
            element_position = area_element.GetParamValue(SharedParamsConfig.Instance.VISPosition)
            if area_element.Category.IsId(BuiltInCategory.OST_DuctCurves):
                element_area = self.__get_duct_area(area_element)
            else:
                element_area = self.__get_fitting_area(area_element)

            if element_position in duct_dict:
                duct_dict[element_position] += element_area
            else:
                duct_dict[element_position] = element_area

        for area_element in area_elements:
            element_position = area_element.GetParamValue(SharedParamsConfig.Instance.VISPosition)
            formatted_area = "{:.2f}".format(duct_dict[element_position]).rstrip('0').rstrip('.') + ' м²'
            area_element.SetParamValue(SharedParamsConfig.Instance.VISNote, formatted_area)

    def __fill_id_to_schedule_param(self, specification_settings, elements):
        """
        Заполняет параметр позиции айди элементов.

        Аргументы:
            specification_settings (SpecificationSettings): Настройки спецификации.
            elements (list): Список элементов для заполнения.
        """
        with revit.Transaction("Запись айди"):
            specification_settings.show_all_specification()

            for element in elements:
                element.SetParamValue(SharedParamsConfig.Instance.VISPosition, str(element.Id.IntegerValue))

    def __fill_values(self, specification_settings, elements, fill_areas, fill_numbers):
        """
        Заполняет позицию и примечания для элементов.

        Аргументы:
            specification_settings (SpecificationSettings): Настройки спецификации.
            elements (list): Список элементов для заполнения.
            fill_areas (bool): Флаг для заполнения площадей.
            fill_numbers (bool): Флаг для заполнения номеров.
        """
        with revit.Transaction("Запись номера"):
            section_data = self.active_view.GetTableData().GetSectionData(SectionType.Body)
            row = section_data.FirstRowNumber

            position_number = 0
            old_sort_rule = ''
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
        """
        Проверяет наличие параметра позиции у элементов.

        Аргументы:
            elements (list): Список элементов для проверки.
        """
        for element in elements:
            if not element.IsExistsParam(SharedParamsConfig.Instance.VISPosition):
                forms.alert(
                    'Параметр "ФОП_ВИС_Позиция" для некоторых элементов спецификации является параметром типа. '
                    'Нумерация этих элементов будет пропущена.',
                    "Внимание")
                return

    def __check_duct_note_param(self, elements):
        """
        Проверяет наличие параметра примечания у воздуховодов.

        Аргументы:
            elements (list): Список элементов для проверки.
        """
        for element in elements:
            if element.Category.IsId(BuiltInCategory.OST_DuctCurves) and not element.IsExistsParam(SharedParamsConfig.Instance.VISNote):
                forms.alert(
                    'Параметр "ФОП_ВИС_Примечание" является параметром типа воздуховодов. '
                    'Примечания не будут заполнены',
                    "Внимание")
                return

    def fill_position_and_notes(self, fill_numbers=False, fill_areas=False):
        """
        Основной метод для заполнения позиций и примечаний в спецификации.

        Аргументы:
            fill_numbers (bool): Флаг для заполнения номеров.
            fill_areas (bool): Флаг для заполнения площадей.
        """
        if self.doc.IsFamilyDocument:
            forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

        if self.active_view.Category is None or not self.active_view.Category.IsId(BuiltInCategory.OST_Schedules):
            forms.alert("Нумерация и вынесение площади воздуховодов сработают только на активном виде целевой спецификации",
                        "Ошибка", exitscript=True)

        elements = FilteredElementCollector(self.doc, self.doc.ActiveView.Id)

        # Проверяем параметры
        self.__check_position_param(elements)
        self.__check_duct_note_param(elements)

        # Если хоть один элемент спеки на редактировании - отменяем выполнение, нужно освобождать
        self.__check_edited_elements(elements)

        specification_settings = SpecificationSettings(self.active_view.Definition)

        # Заполняем айди в параметр позиции элементов для их чтения
        self.__fill_id_to_schedule_param(specification_settings, elements)

        # заполням значения нумерации и примечаний для воздуховодов
        self.__fill_values(specification_settings, elements, fill_areas, fill_numbers)

