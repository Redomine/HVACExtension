#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Добавление параметров'
__doc__ = "Добавление параметров в семейство для за полнения спецификации и экономической функции"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

from Autodesk.Revit.DB import *
import os
from pyrevit import revit
import sys


doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView



spFile = doc.Application.OpenSharedParameterFile()

#проверяем тот ли файл общих параметров подгружен
spFileName = str(doc.Application.SharedParametersFilename)
spFileName = spFileName.split('\\')
spFileName = spFileName[-1]



if "ФОП_v1.txt" != spFileName:
    try:
        doc.Application.SharedParametersFilename = str(os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\ФОП_v1.txt"
    except Exception:
        print 'По стандартному пути не найден файл общих параметров, обратитесь в бим отдел или замените вручную на ФОП_v1.txt'

paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения' ,'ФОП_ВИС_Масса', 'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число', 'ФОП_ВИС_Узел', 'ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка',
             'ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ТИП_Назначение', 'ФОП_ТИП_Число', 'ФОП_ТИП_Единица измерения',
             'ФОП_ТИП_Код', 'ФОП_ТИП_Наименование работы', 'ФОП_ВИС_Имя трубы из сегмента']

pipeparaNames = ['ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка', 'ФОП_ВИС_Имя трубы из сегмента']

projectparaNames = ['ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб']

catFittings = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctFitting)
catPipeFittings = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting)
catPipeCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves)
catCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctCurves)
catFlexCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FlexDuctCurves)
catFlexPipeCurves = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FlexPipeCurves)
catTerminals = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctTerminal)
catAccessory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctAccessory)
catPipeAccessory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory)
catEquipment = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment)
catInsulations = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctInsulations)
catPipeInsulations = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeInsulations)
catPlumbingFixtures = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlumbingFixtures)
catSprinklers = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sprinklers)
catInformation = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ProjectInformation)

def make_type_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsElementType()\
                            .ToElements()
    return col



colTerminals = make_type_col(BuiltInCategory.OST_DuctTerminal)
colAccessory = make_type_col(BuiltInCategory.OST_DuctAccessory)
colPipeAccessory = make_type_col(BuiltInCategory.OST_PipeAccessory)
colEquipment = make_type_col(BuiltInCategory.OST_MechanicalEquipment)
colPlumbingFixtures= make_type_col(BuiltInCategory.OST_PlumbingFixtures)

collections = [colTerminals, colAccessory, colPipeAccessory, colEquipment, colPlumbingFixtures]

colPipeTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()

cats = [catFittings, catPipeFittings, catPipeCurves, catCurves, catFlexCurves, catFlexPipeCurves, catTerminals, catAccessory,
        catPipeAccessory, catEquipment, catInsulations, catPipeInsulations, catPlumbingFixtures, catSprinklers]

ductCats = [catCurves, catInsulations]

nodeCats = [catTerminals, catAccessory, catPipeAccessory, catEquipment, catPlumbingFixtures]

pipeCats = [catPipeCurves]

infCats = [catInformation]




#проверка на наличие нужных параметров

map = doc.ParameterBindings
it = map.ForwardIterator()
while it.MoveNext():
    newProjectParameterData = it.Key.Name
    if str(newProjectParameterData) in paraNames:
        paraNames.remove(str(newProjectParameterData))



uiDoc = __revit__.ActiveUIDocument
sel = uiDoc.Selection

catSet = doc.Application.Create.NewCategorySet()
catDuctSet = doc.Application.Create.NewCategorySet()
catPipeSet = doc.Application.Create.NewCategorySet()
nodeCatSet = doc.Application.Create.NewCategorySet()
infCatSet = doc.Application.Create.NewCategorySet()

for cat in cats:
    catSet.Insert(cat)

for cat in ductCats:
    catDuctSet.Insert(cat)

for cat in pipeCats:
    catPipeSet.Insert(cat)

for cat in nodeCats:
    nodeCatSet.Insert(cat)

for cat in infCats:
    infCatSet.Insert(cat)

with revit.Transaction("Добавление параметров"):
    if len(paraNames) > 0:
        addedNames = []
        for name in paraNames:
            for dG in spFile.Groups:
                group = "03_ВИС"
                if "ФОП_ТИП" in name:
                    group = "10_ТИП_Классификатор видов работ"
                if str(dG.Name) == group:
                    myDefinitions = dG.Definitions
                    eDef = myDefinitions.get_Item(name)
                    if name == "ФОП_ВИС_Минимальная толщина воздуховода":
                        newIB = doc.Application.Create.NewTypeBinding(catDuctSet)
                    elif name in pipeparaNames:
                        newIB = doc.Application.Create.NewTypeBinding(catPipeSet)

                    elif name == "ФОП_ВИС_Узел":
                        newIB = doc.Application.Create.NewTypeBinding(nodeCatSet)

                    elif name == 'ФОП_ВИС_Запас изоляции' or name == 'ФОП_ВИС_Запас воздуховодов/труб':
                        newIB = doc.Application.Create.NewInstanceBinding(infCatSet)

                    else:
                        newIB = doc.Application.Create.NewInstanceBinding(catSet)
                    #print name
                    #print eDef
                    doc.ParameterBindings.Insert(eDef, newIB)

                    if name in pipeparaNames:
                        for element in colPipeTypes:
                            if name == 'ФОП_ВИС_Днар х Стенка':
                                element.LookupParameter(name).Set(1)
                            else:
                                element.LookupParameter(name).Set(0)

                    if name == "ФОП_ВИС_Узел":
                        for collection in collections:
                            for element in collection:
                                element.LookupParameter(name).Set(0)

                    if name == 'ФОП_ВИС_Запас изоляции' or name == 'ФОП_ВИС_Запас воздуховодов/труб':
                        inf = doc.ProjectInformation.LookupParameter(name)
                        inf.Set(10)







