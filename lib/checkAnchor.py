#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление параметров'
__doc__ = "Добавление параметров в семейство для за полнения спецификации и экономической функции"


import clr

import Redomine

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)


from Autodesk.Revit.DB import *
import os
from pyrevit import revit
import sys
from Autodesk.Revit.DB import FamilySource, IFamilyLoadOptions

doc = Redomine.getDocIfItsWorkshared()

IFamilyLoadOptions

actualVersion = '1.01'

# anchor1Path = str(os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\_Якорный элемент.rfa"
# anchor2Path = str(os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\_Якорный элемен(пустой).rfa"
# anchor3Path = str(os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\_Якорный элемен(металл и краска).rfa"

anchor1Name = "_Якорный элемент"
# anchor2Name = "_Якорный элемен(пустой)"
# anchor3Name = "_Якорный элемен(металл и краска)"
# anchor4Name = "_Якорный элемент(Расходники)"

famNames = [anchor1Name]

def check_anchor(showText = True):
    with revit.Transaction("Проверка якорных элементов"):
        if WorksetTable.IsWorksetNameUnique(doc, '99_Немоделируемые элементы'):
            targetWorkset = Workset.Create(doc, '99_Немоделируемые элементы')
            print 'Был создан рабочий набор "99_Немоделируемые элементы". Откройте диспетчер рабочих наборов и снимите галочку с параметра "Видимый на всех видах". ' \
                  'В данном рабочем наборе будут создаваться немоделируемые элементы и требуется исключить их видимость.'
        else:
            fws = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
            for ws in fws:
                if ws.Name == '99_Немоделируемые элементы' and ws.IsVisibleByDefault:

                    print 'Рабочий набор "99_Немоделируемые элементы" на данный момент отображается на всех видах. Откройте диспетчер рабочих наборов и снимите галочку с параметра "Видимый на всех видах". ' \
                  'В данном рабочем наборе будут создаваться немоделируемые элементы и требуется исключить их видимость.'

        collector = FilteredElementCollector(doc)
        collector.OfCategory(BuiltInCategory.OST_GenericModel)
        collector.OfClass(FamilySymbol)
        famtypeitr = collector.GetElementIdIterator()
        famtypeitr.Reset()

        anchor1 = False
        # anchor2 = False
        # anchor3 = False
        # anchor4 = False
        wayToPrint = False

        for element in famtypeitr:
            famtypeID = element
            famsymb = doc.GetElement(famtypeID)

            famName = famsymb.Family.Name
            if famName in famNames:
                if famsymb.LookupParameter("Версия"):
                    famVersion = famsymb.LookupParameter("Версия").AsString()
                else:
                    famVersion = 0

                if famName == anchor1Name:
                    anchor1 = True


                if famVersion != actualVersion:
                    if showText:
                        print 'Версия семейства ' + famsymb.Family.Name + ' расходится с актуальной, обновите его из шаблона.'
                        wayToPrint = True


        if showText:
            if not anchor1:
                print "Не загружено семейство " + anchor1Name
                wayToPrint = True


            if wayToPrint == True:
                print 'Для внутренних инженеров актуальные семейства и шаблоны лежат по пути:'
                print 'W:\\Департаменты\\Проектный институт\\Отдел стандартизации BIM и RD\\BIM-Ресурсы\\2-Стандарты\\11 - Спецификации ОВ-ВК'
                print 'Для инженеров подрядных организаций, обратитесь к ведущему BIM-координатору, или проверьте облако семейств'

        else:
            if not anchor1:
                return False
            else:
                if famVersion != actualVersion:
                    return False
                else:
                    return True