#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Пустой элемент'
__doc__ = "Генерирует в модели пустой якорный элемент"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")


import sys
import System
import dosymep
import paraSpec


clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)


from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from System import Guid
from pyrevit import revit


from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert





doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
selectedIds = uidoc.Selection.GetElementIds()


collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_GenericModel)
collector.OfClass(FamilySymbol)
famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()

def make_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

def setElement(element, name, setting):
    if setting == 'ФОП_ВИС_Масса':
        if setting == 'None':
            setting = ''

        if setting == None:
            setting = ''
    if name == 'ФОП_ВИС_Число' or name == 'ФОП_ВИС_Масса':
        element.LookupParameter(name).Set(setting)
    else:
        try:
            element.LookupParameter(name).Set(str(setting))
        except:
            element.LookupParameter(name).Set(setting)


class collapsedElements:
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

def new_position(element):
    #ищем стадию спеки для присвоения
    phaseOk = False
    for phase in doc.Phases:
        if phase.Name == 'Спецификация':
            phaseid = phase.Id
            phaseOk = True

    if not phaseOk:
        print 'Не удалось присвоить стадию спецификация, проверьте список стадий'
        sys.exit()

    # создаем заглушки по элементов собранных из таблицы
    loc = XYZ(0, 0, 0)

    temporary.Activate()
    familyInst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)
    # собираем список из созданных заглушек
    colModel = make_col(BuiltInCategory.OST_GenericModel)
    for model in colModel:
        if model.LookupParameter('Семейство').AsValueString() == '_Якорный элемен(пустой)':
            model.CreatedPhaseId = phaseid


    dummy = familyInst
    setElement(dummy, 'ФОП_Номер корпуса', element.corp)
    setElement(dummy, 'ФОП_Номер секции', element.sec)
    setElement(dummy, 'ФОП_Этаж', element.floor)
    setElement(dummy, 'ФОП_ВИС_Имя системы', element.system)
    setElement(dummy, 'ФОП_ВИС_Группирование', element.group)
    setElement(dummy, 'ФОП_ВИС_Наименование комбинированное', element.name)
    setElement(dummy, 'ФОП_ВИС_Марка', element.mark)
    setElement(dummy, 'ФОП_ВИС_Код изделия', element.art)
    setElement(dummy, 'ФОП_ВИС_Завод-изготовитель', element.maker)
    setElement(dummy, 'ФОП_ВИС_Единица измерения', element.izm)
    setElement(dummy, 'ФОП_ВИС_Число', element.number)
    setElement(dummy, 'ФОП_ВИС_Масса', element.mass)
    setElement(dummy, 'ФОП_ВИС_Примечание', element.comment)
    setElement(dummy, 'ФОП_Экономическая функция', element.EF)


def script_execute():
    with revit.Transaction("Добавление расчетных элементов"):
        element = doc.GetElement(selectedIds[0])
        corp = element.LookupParameter('ФОП_Номер корпуса').AsString()
        sec = element.LookupParameter('ФОП_Номер секции').AsString()
        floor = element.LookupParameter('ФОП_Этаж').AsString()
        system = element.LookupParameter('ФОП_ВИС_Имя системы').AsString()
        parentGroup = element.LookupParameter('ФОП_ВИС_Группирование').AsString()
        parentEF = element.LookupParameter('ФОП_Экономическая функция').AsString()
        group = parentGroup + '_1'


        hollowElement= collapsedElements(corp=corp,
                             sec=sec,
                             floor=floor,
                             system=system,
                             group=group,
                             name='',
                             mark='',
                             art= '',
                             maker='',
                             izm='',
                             number='',
                             mass='',
                             comment='',
                             EF=parentEF)

        new_position(hollowElement)

is_temporary_in = False

for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)


    if famsymb.Family.Name == '_Якорный элемен(пустой)':
        temporary = famsymb
        is_temporary_in = True



if is_temporary_in == False:
    print 'Не обнаружен якорный элемен(пустой). Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()


status = paraSpec.check_parameters()
viewIsShed = view.Category.IsId(BuiltInCategory.OST_Schedules)

if not status:
    if viewIsShed:
        if 0 == selectedIds.Count:
            print 'Выделите целевой элемент'
        else:
            script_execute()
    else:
        print 'Добавление пустого элемента возможно только на целевой спецификации'

