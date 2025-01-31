#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

from low_voltage_task_class_lib import JsonOperator, EditedReport, LowVoltageSystemData

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

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

operator = JsonOperator(doc, uiapp)

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if doc.IsFamilyDocument:
        forms.alert("Надстройка не предназначена для работы с семействами", "Ошибка", exitscript=True)

    file_folder_path = operator.get_document_path()

    # Находим все JSON-файлы в директории
    json_files = glob.glob(os.path.join(file_folder_path, "*.json"))
    if not json_files:
        forms.alert("Ревизии не найдены", "Ошибка", exitscript=True)

    # Находим файл с самым поздним временем модификации
    latest_file = max(json_files, key=os.path.getmtime)

    file_exists = latest_file is not None and os.path.exists(latest_file)

    if not file_exists:
        forms.alert('Файлов-заданий не существует',
                    "Ошибка", exitscript=True)
    else:
        report = ("Данное действие невозможно отменить. "
                  "Данные последней ревизии будут удалены и могут быть только добавлены заново. Удалить файл?")

        result = operator.show_dialog(report)

        if result:
            os.remove(latest_file)

script_execute()
