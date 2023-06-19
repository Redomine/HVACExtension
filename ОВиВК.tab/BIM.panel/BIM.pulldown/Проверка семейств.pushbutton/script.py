# -*- coding: utf-8 -*-


__title__ = 'Код в семействах'
__doc__ = "Обновляет код внутри семейств"
import clr
import sys
import System
from System.Collections.Generic import *
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitNodes")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Microsoft.Office.Interop import Excel
from pyrevit import revit
from rpw.ui.forms import SelectFromList
from System import Guid
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Revit import ParamExtensions
from Redomine import *

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import paraSpec


commandData = ExternalCommandData

App = commandData.Application


exel = Excel.ApplicationClass()
tablePath = "C:\\Users\\Mankaev_r\\Desktop\\Проверка базы данных семейств\\Тестовая база данных.xlsx"
try:
    workbook = exel.Workbooks.Open(tablePath, ReadOnly=True)
except Exception:
    print 'Ошибка открытия таблицы, проверьте ее целостность'
    exel.ActiveWorkbook.Close(True)
    sys.exit()

try:
    worksheet = workbook.Sheets["Лист2"]
except Exception:
    print 'Не найден лист с названием Импорт, проверьте файл формы.'
    exel.ActiveWorkbook.Close(True)
    sys.exit()

xlrange = worksheet.Range["A1", "AZ500"]

exel.ActiveWorkbook.Close(True)




Path = "C:\\Users\\Mankaev_r\\Desktop\\Проверка базы данных семейств\\Тест 1\\Взр_Ршт_Решетка_Прямоугольная_Вытяжная_Арктос.rfa"

doc = __revit__.Application.OpenDocumentFile(Path)

set = doc.FamilyManager.Parameters
manager = doc.FamilyManager

t = Transaction(doc, 'Test')

t.Start()

for param in set:
    if str(param.Definition.Name) == "Код по классификатору": classCode = param


manager.Set(classCode, "Test")

t.Commit()


doc.Close(True)
