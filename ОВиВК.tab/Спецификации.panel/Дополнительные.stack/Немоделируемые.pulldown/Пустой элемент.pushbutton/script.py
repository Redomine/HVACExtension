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
import checkAnchor


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
from Redomine import *


from System.Runtime.InteropServices import Marshal
from rpw.ui.forms import select_file
from rpw.ui.forms import TextInput
from rpw.ui.forms import SelectFromList
from rpw.ui.forms import Alert




#Исходные данные
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
uidoc = __revit__.ActiveUIDocument
selectedIds = uidoc.Selection.GetElementIds()
nameOfModel = '_Якорный элемент'
description = 'Пустая строка'


def script_execute():
    with revit.Transaction("Добавление расчетных элементов"):
        element = doc.GetElement(selectedIds[0])
        corp = element.LookupParameter('ФОП_Блок СМР').AsString()
        sec = element.LookupParameter('ФОП_Секция СМР').AsString()
        floor = element.LookupParameter('ФОП_Этаж').AsString()
        system = element.LookupParameter('ФОП_ВИС_Имя системы').AsString()
        parentGroup = element.LookupParameter('ФОП_ВИС_Группирование').AsString()
        parentEF = element.LookupParameter('ФОП_Экономическая функция').AsString()
        group = parentGroup + '_1'



        hollowElement= rowOfSpecification(corp=corp,
                             sec=sec,
                             floor=floor,
                             system=system,
                             group=group,
                             name='',
                             mark='',
                             art= '',
                             maker='',
                             unit='',
                             number='',
                             mass='',
                             comment='',
                             EF=parentEF)

        new_position([hollowElement], temporary, nameOfModel, description)


temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, nameOfModel)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

if temporary == None:
    print 'Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()


status = paraSpec.check_parameters()
viewIsShed = view.Category.IsId(BuiltInCategory.OST_Schedules)

if not status:
    if viewIsShed:
        if 0 == selectedIds.Count:
            print 'Выделите целевой элемент'
        else:
            anchor = checkAnchor.check_anchor(showText = False)
            if anchor:
                script_execute()
    else:
        print 'Добавление пустого элемента возможно только на целевой спецификации'

