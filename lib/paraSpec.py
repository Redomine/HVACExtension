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

class shared_parameter:
    def insert(self, eDef, set, istype):
        if istype:
            newBinding = doc.Application.Create.NewTypeBinding(set)
        else:
            newBinding = doc.Application.Create.NewInstanceBinding(set)
        doc.ParameterBindings.Insert(eDef, newBinding, BuiltInParameterGroup.PG_DATA)

    def reinsert(self, eDef, set, istype):
        if istype:
            newBinding = doc.Application.Create.NewTypeBinding(set)
        else:
            newBinding = doc.Application.Create.NewInstanceBinding(set)
        doc.ParameterBindings.Reinsert(eDef, newBinding, BuiltInParameterGroup.PG_DATA)

    def __init__(self, name, definition):
        self.name = name
        self.group = "03_ВИС"
        if "ФОП_ТИП" in self.name:
            self.group = "10_ТИП_Классификатор видов работ"
        self.eDef = definition.get_Item(name)
        self.set = catSet
        self.istype = False

        if name == "ФОП_ВИС_Минимальная толщина воздуховода":
            self.istype = True
            self.set = catDuctSet
        if name == 'ФОП_ВИС_Расчет металла для креплений':
            self.istype = True
            self.set = ductandpipeCatSet
        if name == 'ФОП_ВИС_Совместно с воздуховодом':
            self.istype = True
            self.set = ductInsCatSet
        if name == 'ФОП_ВИС_Узел':
            self.istype = True
            self.set = nodeCatSet
        if name in pipeparaNames:
            self.istype = True
            self.set = catPipeSet
        if name in projectparaNames:
            self.istype = False
            self.set = infCatSet

        self.insert(self.eDef, self.set, self.istype)

def check_parameters():
    # проверка на наличие нужных параметров
    map = doc.ParameterBindings
    it = map.ForwardIterator()
    while it.MoveNext():
        newProjectParameterData = it.Key.Name
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

def get_cats(list_of_cats):
    cats = []
    for cat in list_of_cats:
        cats.append(doc.Settings.Categories.get_Item(cat))
    set = doc.Application.Create.NewCategorySet()
    for cat in cats:
        set.Insert(cat)
    return set


paraNames = ['ФОП_ВИС_Группирование', 'ФОП_ВИС_Единица измерения', 'ФОП_ВИС_Масса',
             'ФОП_ВИС_Минимальная толщина воздуховода',
             'ФОП_ВИС_Наименование комбинированное', 'ФОП_ВИС_Число', 'ФОП_ВИС_Узел', 'ФОП_ВИС_Ду',
             'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка',
             'ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ТИП_Назначение', 'ФОП_ТИП_Число',
             'ФОП_ТИП_Единица измерения',
             'ФОП_ТИП_Код', 'ФОП_ТИП_Наименование работы', 'ФОП_ВИС_Имя трубы из сегмента', 'ФОП_ВИС_Позиция',
             'ФОП_ВИС_Площади воздуховодов в примечания',
             'ФОП_ВИС_Нумерация позиций', 'ФОП_ВИС_Расчет комплектов заделки', 'ФОП_ВИС_Расчет краски и грунтовки',
             'ФОП_ВИС_Расчет металла для креплений', 'ФОП_ВИС_Совместно с воздуховодом', 'ФОП_ВИС_Марка']

pipeparaNames = ['ФОП_ВИС_Ду', 'ФОП_ВИС_Ду х Стенка', 'ФОП_ВИС_Днар х Стенка', 'ФОП_ВИС_Имя трубы из сегмента',
                 'ФОП_ВИС_Расчет краски и грунтовки']

projectparaNames = ['ФОП_ВИС_Запас изоляции', 'ФОП_ВИС_Запас воздуховодов/труб', 'ФОП_ВИС_Нумерация позиций',
                    'ФОП_ВИС_Площади воздуховодов в примечания', 'ФОП_ВИС_Расчет комплектов заделки']

# сеты категорий идущие по стандарту

catSet = get_cats([BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeCurves,
                   BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_FlexDuctCurves,
                   BuiltInCategory.OST_FlexPipeCurves,
                   BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_DuctAccessory,
                   BuiltInCategory.OST_PipeAccessory,
                   BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_DuctInsulations,
                   BuiltInCategory.OST_PipeInsulations,
                   BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_Sprinklers])

# нестандартные сеты категорий
catDuctSet = get_cats([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_DuctInsulations])
ductInsCatSet = get_cats([BuiltInCategory.OST_DuctInsulations])
nodeCatSet = get_cats([BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_DuctAccessory,
                       BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
                       BuiltInCategory.OST_PlumbingFixtures])
catPipeSet = get_cats([BuiltInCategory.OST_PipeCurves])
infCatSet = get_cats([BuiltInCategory.OST_ProjectInformation])
ductandpipeCatSet = get_cats([BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves])





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


    #проверка на наличие нужных параметров
    map = doc.ParameterBindings
    it = map.ForwardIterator()
    while it.MoveNext():
        newProjectParameterData = it.Key.Name
        if str(newProjectParameterData) in paraNames:
            paraNames.remove(str(newProjectParameterData))

    uiDoc = __revit__.ActiveUIDocument
    sel = uiDoc.Selection

    for dG in spFile.Groups:
        vis_group = "03_ВИС"
        type_group = "10_ТИП_Классификатор видов работ"
        if str(dG.Name) == vis_group:
            visDefinitions = dG.Definitions
        if str(dG.Name) == type_group:
            typeDefinitions = dG.Definitions

    colPipeTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()




    with revit.Transaction("Добавление параметров"):
        for name in paraNames:
            eDef = visDefinitions
            if "ФОП_ТИП" in name:
                eDef = typeDefinitions

            parameter = shared_parameter(name, eDef)


            if name in pipeparaNames:
                for element in colPipeTypes:
                    if name == 'ФОП_ВИС_Днар х Стенка':
                        element.LookupParameter(name).Set(1)
                    else:
                        element.LookupParameter(name).Set(0)

                if name == 'ФОП_ВИС_Запас изоляции' or name == 'ФОП_ВИС_Запас воздуховодов/труб':
                    inf = doc.ProjectInformation.LookupParameter(name)
                    inf.Set(10)











