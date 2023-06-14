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
import paraSpec

from Microsoft.Office.Interop import Excel
from Redomine import *
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



def setElement(element, name, setting):
    if name == 'ФОП_ВИС_Число':
        try:
            element.LookupParameter(name).Set(setting)
        except:
            element.LookupParameter(name).Set(0)
    if name == 'ФОП_ВИС_Масса':
        pass
    else:
        if setting == None or setting == 'None':
            setting = ''
        element.LookupParameter(name).Set(str(setting))


def new_position(calculation_elements, phaseid):

    #создаем заглушки по элементов собранных из таблицы
    loc = XYZ(0, 0, 0)

    temporary.Activate()
    for element in calculation_elements:
        familyInst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)

    #собираем список из созданных заглушек
    colModel = make_col(BuiltInCategory.OST_GenericModel)
    Models = []

    fws = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in fws:
        if ws.Name == '99_Немоделируемые элементы':
            WORKSET_ID = ws.Id


    for element in colModel:
        try:
            if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемент':
                ews = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                ews.Set(WORKSET_ID.IntegerValue)
        except Exception:
                 print 'Не удалось присвоить рабочий набор "99_Немоделируемые элементы", проверьте список наборов'




        # if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемент':
        #     try:
        #         element.CreatedPhaseId = phaseid
        #
        #         ews = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        #
        #         print ews.AsValueString()
        #         ews.Set(WORKSET_ID.ItnegerValue)
        #         print ews.AsValueString()
        #
        #     except Exception:
        #         print 'Не удалось присвоить стадию спецификация, проверьте список стадий'
            Models.append(element)


    index = 1
    #для первого элмента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for position in calculation_elements:
        posGroup = str(position.group) + '_' + str(index)
        index+=1
        if position.group == None:
            posGroup = 'None'
        posName = position.name
        if position.name == None:
            posName = 'None'
        posMark = position.mark
        if position.mark == None:
            posMark = 'None'

        group = posGroup + posName + posMark
        dummy = Models[0]
        setElement(dummy, 'ФОП_Блок СМР', position.corp)
        setElement(dummy, 'ФОП_Секция СМР', position.sec)
        setElement(dummy, 'ФОП_Этаж', position.floor)
        setElement(dummy, 'ФОП_ВИС_Имя системы', position.system)
        setElement(dummy, 'ФОП_ВИС_Группирование', group)
        setElement(dummy, 'ФОП_ВИС_Наименование комбинированное', position.name)
        setElement(dummy, 'ФОП_ВИС_Марка', position.mark)
        setElement(dummy, 'ФОП_ВИС_Код изделия', position.art)
        setElement(dummy, 'ФОП_ВИС_Завод-изготовитель', position.maker)
        setElement(dummy, 'ФОП_ВИС_Единица измерения', position.izm)
        setElement(dummy, 'ФОП_ВИС_Число', position.number)
        setElement(dummy, 'ФОП_ВИС_Масса', position.mass)
        setElement(dummy, 'ФОП_ВИС_Примечание', position.comment)
        setElement(dummy, 'ФОП_Экономическая функция', position.EF)

        Models.pop(0)





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

class shedulePosition:
    def __init__(self, corp, sec, floor, system, group, name, mark, art, maker, izm, number, mass, comment, EF):
        self.corp = corp
        self.sec = sec
        self.floor = floor
        self.system = system
        self.group = group
        self.name = name
        self.mark = mark
        self.art = art
        self.maker = maker
        self.izm = izm
        self.number = number
        self.mass = mass
        self.comment = comment
        self.EF = EF

def script_execute():

    exel = Excel.ApplicationClass()
    filepath = select_file()


    ADSK_System_Names = []
    System_Named = True

    try:
        workbook = exel.Workbooks.Open(filepath)
    except Exception:
        print 'Ошибка открытия таблицы, проверьте ее целостность'
        sys.exit()
    sheet_name = 'Импорт'

    try:
        worksheet = workbook.Sheets[sheet_name]
    except Exception:
        print 'Не найден лист с названием Импорт, проверьте файл формы.'
        sys.exit()

    xlrange = worksheet.Range["A1", "AZ500"]

    FOP_Corp = 0
    FOP_Sec = 1
    FOP_Floor = 2
    FOP_System = 3
    FOP_Group = 4
    FOP_Name = 5
    FOP_Mark = 6
    FOP_Art = 7
    FOP_Maker = 8
    FOP_Izm = 9
    FOP_Number = 10
    FOP_Mass = 11
    FOP_Comment = 12
    FOP_EF = 13

    report_rows = set()

    for element in colModel:
        if element.LookupParameter('Редактирует'):
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
            newPos = shedulePosition(corp = xlrange.value2[row, FOP_Corp],
                            sec = xlrange.value2[row, FOP_Sec],
                            floor = xlrange.value2[row, FOP_Floor],
                            system = xlrange.value2[row, FOP_System],
                            group = xlrange.value2[row, FOP_Group],
                            name = xlrange.value2[row, FOP_Name],
                            mark = xlrange.value2[row, FOP_Mark],
                            art = xlrange.value2[row, FOP_Art],
                            maker = xlrange.value2[row, FOP_Maker],
                            izm = xlrange.value2[row, FOP_Izm],
                            number = xlrange.value2[row, FOP_Number],
                            mass = xlrange.value2[row, FOP_Mass],
                            comment = xlrange.value2[row, FOP_Comment],
                            EF = xlrange.value2[row, FOP_EF])


            row += 1

            calculation_elements.append(newPos)



        # for phase in doc.Phases:
        #     if phase.Name == 'Спецификация':
        #         phaseid = phase.Id




        # в следующем блоке генерируем новые экземпляры пустых семейств куда уйдут расчеты
        new_position(calculation_elements, phaseid)





    exel.ActiveWorkbook.Close(True)
    Marshal.ReleaseComObject(worksheet)
    Marshal.ReleaseComObject(workbook)
    Marshal.ReleaseComObject(exel)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

status = paraSpec.check_parameters()

if not status:
    script_execute()





