#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Импорт немоделируемых'
__doc__ = "Генерирует в модели элементы в соответствии с их ведомостью"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import sys
import System

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit

from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert





doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col 
    
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)
# create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_GenericModel)
collector.OfClass(FamilySymbol)
famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()





is_temporary_in = False

for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)

    if famsymb.Family.Name == '_Якорный элемент':
        temporary = famsymb
        is_temporary_in = True

if is_temporary_in == False:
    print 'Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()
paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения' ,'ФОП_ВИС_Масса', 'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число', 'ФОП_ВИС_Узел', 'ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка',
             'ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ТИП_Назначение', 'ФОП_ТИП_Число', 'ФОП_ТИП_Единица измерения',
             'ФОП_ТИП_Код', 'ФОП_ТИП_Наименование работы']


#проверка на наличие нужных параметров
map = doc.ParameterBindings
it = map.ForwardIterator()
while it.MoveNext():
    newProjectParameterData = it.Key.Name
    if str(newProjectParameterData) in paraNames:
        paraNames.remove(str(newProjectParameterData))
if len(paraNames) > 0:
    print 'Необходимо добавить параметры'
    for name in paraNames:
        print name
    sys.exit()


exel = Excel.ApplicationClass()
filepath = select_file()





def setElement(element, name, setting):
    try:
        if setting == None:
            pass
        else:
            element.LookupParameter(name).Set(setting)
            if name == 'ADSK_Единица измерения':
                element.LookupParameter('ФОП_ТИП_Единица измерения').Set(setting)
            if name == 'ФОП_ВИС_Число':
                element.LookupParameter('ФОП_ТИП_Число').Set(setting)
            if name == 'ФОП_ВИС_Наименование комбинированное':
                element.LookupParameter('ФОП_ТИП_Назначение').Set(setting)
    except Exception:
        pass


def new_position(calculation_elements):
    #создаем заглушки по элементов собранных из таблицы
    loc = XYZ(0, 0, 0)

    temporary.Activate()
    for element in calculation_elements:
        familyInst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)

    #собираем список из созданных заглушек
    colModel = make_col(BuiltInCategory.OST_GenericModel)
    Models = []
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемент':
            try:
                element.CreatedPhaseId = phaseid
            except Exception:
                print 'Не удалось присовить стадию спецификация, проверьте список стадий'

            Models.append(element)

    #для первого элмента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for element in calculation_elements:
        dummy = Models[0]
        setElement(dummy, 'ADSK_Имя системы', element[0])
        setElement(dummy, 'ФОП_ТИП_Код', element[1])
        setElement(dummy, 'ФОП_ТИП_Наименование работы', element[2])
        setElement(dummy, 'ФОП_ВИС_Группирование', element[3])
        setElement(dummy, 'ФОП_ВИС_Наименование комбинированное', element[4])
        setElement(dummy, 'ADSK_Марка', element[5])
        setElement(dummy, 'ADSK_Код изделия', element[6])
        setElement(dummy, 'ADSK_Завод-изготовитель', element[7])
        setElement(dummy, 'ADSK_Единица измерения', element[8])
        setElement(dummy, 'ФОП_ВИС_Число', element[9])
        setElement(dummy, 'ADSK_Масса', element[10])
        setElement(dummy, 'ADSK_Примечание', element[11])
        Models.pop(0)

ADSK_System_Names = []
System_Named = True

try:
    workbook = exel.Workbooks.Open(filepath)
except Exception:
    sys.exit()
sheet_name = 'Импорт'

try:
    worksheet = workbook.Sheets[sheet_name]
except Exception:
    print 'Не найден лист с названием Импорт, проверьте файл формы.'
    sys.exit()

xlrange = worksheet.Range["A1", "AZ500"]

ADSK_System = 0
FOP_class = 1
FOP_work = 2
FOP_Group = 3
FOP_Name = 4
ADSK_Mark = 5
ADSK_Art = 6
ADSK_Maker = 7
ADSK_Izm = 8
FOP_Number = 9
FOP_Mass = 10
ADSK_Comment = 11

report_rows = set()


for element in colModel:
    edited_by = element.LookupParameter('Редактирует').AsString()
    if edited_by and edited_by != __revit__.Application.Username:
        print "Якорные элементы не были обработаны, так как были заняты пользователями:"
        print edited_by
        sys.exit()

with revit.Transaction("Добавление расчетных элементов"):

    #при каждом повторе расчета удаляем старые версии
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемент':
            doc.Delete(element.Id)

    calculation_elements = []
    row = 2
    while True:
        if xlrange.value2[row, FOP_Name] == None:
            break
        System = xlrange.value2[row, ADSK_System]
        Class = xlrange.value2[row, FOP_class]
        Work = xlrange.value2[row, FOP_work]
        Group = xlrange.value2[row, FOP_Group]
        Name = xlrange.value2[row, FOP_Name]
        Mark = xlrange.value2[row, ADSK_Mark]
        Art = xlrange.value2[row, ADSK_Art]
        Maker = xlrange.value2[row, ADSK_Maker]
        Izm = xlrange.value2[row, ADSK_Izm]
        Number = xlrange.value2[row, FOP_Number]
        Mass = xlrange.value2[row, FOP_Mass]
        Comment = xlrange.value2[row, ADSK_Comment]

        row += 1
        calculation_elements.append([System, Class, Work, Group, Name, Mark, Art, Maker, Izm, Number, Mass, Comment])



    for phase in doc.Phases:
        if phase.Name == 'Спецификация':
            phaseid = phase.Id

    # в следующем блоке генерируем новые экземпляры пустых семейств куда уйдут расчеты
    new_position(calculation_elements)


exel.ActiveWorkbook.Close(True)
Marshal.ReleaseComObject(worksheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(exel)








