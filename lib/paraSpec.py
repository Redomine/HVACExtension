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

doc = __revit__.ActiveUIDocument.Document  # type: Document
paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения', 'ФОП_ВИС_Масса',
             'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число', 'ФОП_ВИС_Узел', 'ФОП_ВИС_Ду',
             'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка',
             'ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ТИП_Назначение', 'ФОП_ТИП_Число',
             'ФОП_ТИП_Единица измерения',
             'ФОП_ТИП_Код', 'ФОП_ТИП_Наименование работы', 'ФОП_ВИС_Имя трубы из сегмента', 'ФОП_ВИС_Позиция',
             'ФОП_ВИС_Площади воздуховодов в примечания',
             'ФОП_ВИС_Нумерация позиций', 'ФОП_ВИС_Расчет комплектов заделки', 'ФОП_ВИС_Расчет краски и грунтовки',
             'ФОП_ВИС_Расчет металла для креплений']


def check_parameters():
    # проверка на наличие нужных параметров
    map = doc.ParameterBindings
    it = map.ForwardIterator()
    while it.MoveNext():
        newProjectParameterData = it.Key.Name
        #print it.Key.ParameterGroup
        if str(newProjectParameterData) in paraNames:
            paraNames.remove(str(newProjectParameterData))

    if len(paraNames) > 0:
        try:
            script_execute()
            print 'Были добавлен параметры, перезапустите скрипт'
            return True
        except Exception:
            print 'Не удалось добавить параметры'

def make_type_col(category):
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsElementType()\
                            .ToElements()
    return col

def script_execute():
    spFile = doc.Application.OpenSharedParameterFile()
    # проверяем тот ли файл общих параметров подгружен
    spFileName = str(doc.Application.SharedParametersFilename)
    spFileName = spFileName.split('\\')
    spFileName = spFileName[-1]

    if "ФОП_v1.txt" != spFileName:
        try:
            doc.Application.SharedParametersFilename = str(os.environ['USERPROFILE']) + "\\AppData\\Roaming\\pyRevit\\Extensions\\04.OV-VK.extension\\ФОП_v1.txt"
        except Exception:
            print 'По стандартному пути не найден файл общих параметров, обратитесь в BIM-отдел или замените вручную на ФОП_v1.txt'
            sys.exit()


    pipeparaNames = ['ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка', 'ФОП_ВИС_Имя трубы из сегмента',
                     'ФОП_ВИС_Расчет краски и грунтовки']

    projectparaNames = ['ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ВИС_Нумерация позиций',
                        'ФОП_ВИС_Площади воздуховодов в примечания',  'ФОП_ВИС_Расчет комплектов заделки']

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

    ductandpipeCats = [catCurves, catPipeCurves]

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
    ductandpipeCatSet = doc.Application.Create.NewCategorySet()

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

    for cat in ductandpipeCats:
        ductandpipeCatSet.Insert(cat)

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
                        if name == 'ФОП_ВИС_Расчет металла для креплений':
                            newIB = doc.Application.Create.NewTypeBinding(ductandpipeCatSet)
                        elif name in pipeparaNames:
                            newIB = doc.Application.Create.NewTypeBinding(catPipeSet)

                        elif name == "ФОП_ВИС_Узел":
                            newIB = doc.Application.Create.NewTypeBinding(nodeCatSet)

                        elif name in projectparaNames:
                            newIB = doc.Application.Create.NewInstanceBinding(infCatSet)
                        else:
                            newIB = doc.Application.Create.NewInstanceBinding(catSet)

                        doc.ParameterBindings.Insert(eDef, newIB, BuiltInParameterGroup.PG_DATA)

                        if name in pipeparaNames:
                            for element in colPipeTypes:
                                if name == 'ФОП_ВИС_Днар х Стенка':
                                    element.LookupParameter(name).Set(1)
                                else:
                                    element.LookupParameter(name).Set(0)


                        if name == 'ФОП_ВИС_Запас изоляции' or name == 'ФОП_ВИС_Запас воздуховодов/труб':
                            inf = doc.ProjectInformation.LookupParameter(name)
                            inf.Set(10)









