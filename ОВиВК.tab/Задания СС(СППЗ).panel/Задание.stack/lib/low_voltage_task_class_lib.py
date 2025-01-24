#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
import glob
import re
import sys
import json
import os
import ctypes
import codecs

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import forms
from pyrevit import revit
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.SharedParams import *
from collections import defaultdict

from dosymep_libs.bim4everyone import *
from System.Collections.Generic import List

from datetime import datetime, timedelta
from System import Environment

class EditedReport:
    edited_reports = []
    status_report = ''
    edited_report = ''

    def __init__(self, doc):
        self.doc = doc

    def get_element_editor_name(self, element):
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

    def is_elemet_edited(self, element):
        """
        Проверяет, заняты ли элементы другими пользователями.

        Args:
            elements (list): Список элементов для проверки.
        """

        update_status = WorksharingUtils.GetModelUpdatesStatus(self.doc, element.Id)

        if update_status == ModelUpdatesStatus.UpdatedInCentral:
            self.status_report = "Вы владеете элементами, но ваш файл устарел. Выполните синхронизацию. "

        name = self.get_element_editor_name(element)
        if name is not None and name not in self.edited_reports:
            self.edited_reports.append(name)

        if name is not None or update_status == ModelUpdatesStatus.UpdatedInCentral:
            return True

        return False

    def show_report(self):
        if len(self.edited_reports) > 0:
            self.edited_report = \
                ("Часть элементов спецификации занята пользователями: {}".format(", ".join(self.edited_reports)))

        if self.edited_report != '' or self.status_report != '':
            report_message = (self.status_report +
                              ('\n' if (self.edited_report and self.status_report) else '') +
                              self.edited_report)
            forms.alert(report_message, "Ошибка", exitscript=True)

class LowVoltageSystemData:
    id = ''
    valve_base_name = ''
    autor_name = ''
    json_name = ''
    creation_date = ''
    deletion_date = ''

    def __init__(self,
                 id,
                 creation_date='',
                 valve_base_name = '',
                 autor_name = __revit__.Application.Username,
                 json_name = '',
                 deletion_date = ''
                 ):

        self.id = id
        self.valve_base_name = valve_base_name
        self.autor_name = autor_name
        self.json_name = json_name
        self.creation_date = creation_date
        self.deletion_date = deletion_date

    def to_dict(self):
        return {
            "id": str(self.id),
            "valve_base_name": self.valve_base_name,
            "autor_name": self.autor_name,
            "json_name": self.json_name,
            "creation_date": self.creation_date,
            "deletion_date": self.deletion_date
        }

    def insert(self, doc, time):
        element = doc.GetElement(self.id)

        if element is not None:
            element.SetParamValue("ФОП_ВИС_СС Марка задания", self.json_name)
            element.SetParamValue("ФОП_ВИС_СС Дата задания", self.creation_date)
        else:
            if self.deletion_date == '' and self.creation_date != time:
                self.deletion_date = time

class JsonOperator:
    def __init__(self,doc, uiapp):
        self.doc = doc
        self.uiapp = uiapp

    def get_document_path(self):
        path = \
            "W:/Проектный институт/Отд.стандарт.BIM и RD/BIM-Ресурсы/5-Надстройки/Bim4Everyone/A101/MEP/EquipmentNumbering/"

        if not (os.path.exists(path) and os.access(path, os.R_OK) and os.access(path, os.W_OK)):
            version = self.uiapp.VersionNumber
            documents_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

            path = os.path.join(documents_path,
                                'dosymep',
                                str(version),
                                'RevitMEPNumeration/')

            report = ('Нет доступа к сетевому диску. Файлы задания обрабатываются из папки: {} \n'
                      'Открыть папку с файлами?').format(path)

            # Вызов MessageBox из Windows API
            result = ctypes.windll.user32.MessageBoxW(0, report, "Внимание", 4)

            if result == 6:  # IDYES
                os.startfile(path)



        project_name = self.get_project_name()
        project_path = path + project_name

        self.create_folder_if_not_exist(project_path)

        return project_path

    def create_folder_if_not_exist(self, project_path):
        if not os.path.exists(project_path):
            os.makedirs(project_path)

    def send_json_data(self, data, new_file_path):
        project_name = self.get_project_name()

        time = self.get_moscow_time()

        new_file_path = new_file_path + "\СС(СППЗ)_" + project_name + "_" + time + ".json"

        # Преобразование списка объектов в список словарей
        data_dicts = [item.to_dict() for item in data]

        # Запись в JSON
        with codecs.open(new_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(data_dicts, json_file, ensure_ascii=False, indent=4)

    def get_json_data(self, project_path):
        # Находим все JSON-файлы в директории
        json_files = glob.glob(os.path.join(project_path, "*.json"))
        if not json_files:
            return {}

        # Находим файл с самым поздним временем модификации
        latest_file = max(json_files, key=os.path.getmtime)

        # Проверяем, существует ли файл
        if latest_file is not None and os.path.exists(latest_file):
            # Читаем существующие данные
            with codecs.open(latest_file, 'r', encoding='utf-8') as json_file:
                existing_data = json.load(json_file)

                # Конвертируем данные в объекты класса LowVoltageSystemData
                return [
                    LowVoltageSystemData(
                        id=ElementId(int(item["id"])),
                        creation_date=item["creation_date"],
                        valve_base_name=item["valve_base_name"],
                        autor_name=item["autor_name"],
                        json_name=item["json_name"],
                        deletion_date=item["deletion_date"]
                    )
                    for item in existing_data
                ]
        return []  # Если файл не найден, возвращаем пустой список

    def get_moscow_time(self):
        # Получаем текущее время в UTC
        utc_time = datetime.utcnow()

        # Добавляем 3 часа для перехода в часовой пояс Москвы (UTC+3)
        moscow_time = utc_time + timedelta(hours=3)

        # Форматируем время
        formatted_time = moscow_time.strftime("%Y-%m-%d")

        return formatted_time

    def get_project_name(self):
        # Получаем имя пользователя
        username = __revit__.Application.Username

        # Получаем заголовок документа
        title = self.doc.Title

        # Переводим все буквы в верхний регистр
        username_upper = username.upper()
        title_upper = title.upper()

        # Проверяем, является ли имя пользователя частью заголовка
        if username_upper in title_upper:
            # Убираем имя пользователя и подчеркивание перед ним
            project_name = title_upper.replace('_' + username_upper, '').strip()
        else:
            # Если имя пользователя не является частью заголовка, возвращаем заголовок как есть
            project_name = title

        return project_name

    def create_folder_if_not_exist(self, project_path):
        if not os.path.exists(project_path):
            os.makedirs(project_path)