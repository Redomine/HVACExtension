#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Расходники изоляции'
__doc__ = "Генерирует в модели элементы расходных материалов для изоляции"


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
colPipeInsul = make_col(BuiltInCategory.OST_PipeInsulations)
colDuctInsul = make_col(BuiltInCategory.OST_DuctInsulations)
nameOfModel = '_Якорный элемент'
description = 'Расходники изоляции'


class insulType:
    def __init__(self, typeName, name1, name2, name3, mark1, mark2, mark3, unit1, unit2, unit3,
                 maker1, maker2, maker3, expenditure1, expenditure2, expenditure3, isArea1, isArea2, isArea3):
        self.typeName = typeName

        self.name1 = name1
        self.name2 = name2
        self.name3 = name3

        self.mark1 = mark1
        self.mark2 = mark2
        self.mark3 = mark3

        self.unit1 = unit1
        self.unit2 = unit2
        self.unit3 = unit3

        self.maker1 = maker1
        self.maker2 = maker2
        self.maker3 = maker3

        self.expenditure1 = expenditure1
        self.expenditure2 = expenditure2
        self.expenditure3 = expenditure3

        self.isArea1 = isArea1
        self.isArea2 = isArea2
        self.isArea3 = isArea3

class insulPosition:
    def __init__(self, corp, sec, floor, system, name, EF, key):
        self.corp = corp
        self.sec = sec
        self.floor = floor
        self.system = system

        self.typeName = name
        self.EF = EF
        self.key = key

        self.area = 0
        self.length = 0

class objectToGenerate:
    def getNumber(self, expenditure, motherArea):
        return float(expenditure) * float(motherArea)

    def __init__(self, corp, sec, floor, system, group, name, mark,
                 art, maker, unit, number, mass, comment, EF, motherArea, typeName, expenditure):
        self.corp = corp
        self.sec = sec
        self.floor = floor
        self.system = system
        self.group = group
        self.name = name
        self.mark = mark
        self.art = art
        self.maker = maker
        self.unit = unit

        self.mass = mass
        self.comment = comment
        self.EF = EF
        self.motherArea = motherArea
        self.typeName = typeName
        self.expenditure = expenditure

        self.number = self.getNumber(self.expenditure, self.motherArea)


def isToGenerate(insTypeName, expenditure):
    if expenditure != None:
        if float(expenditure) > 0.0001:
            if insTypeName != None:
                if insTypeName != '':
                    return True

def get_area(element):
    information = doc.ProjectInformation
    isol_reserve = 1 + (lookupCheck(information, 'ФОП_ВИС_Запас изоляции').AsDouble() / 100)  # запас площадей

    if element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
        connectors = getConnectors(element)
        for connector in connectors:
            for el in connector.AllRefs:
                if el.Owner.Category.IsId(BuiltInCategory.OST_DuctFitting):
                    return round((get_fitting_area(el.Owner) * isol_reserve), 2)

    area = element.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble()
    area = round((fromRevitToSquareMeters(area) * isol_reserve), 2)
    return area

def get_length(element):
    information = doc.ProjectInformation
    isol_reserve = 1 + (lookupCheck(information, 'ФОП_ВИС_Запас изоляции').AsDouble() / 100)  # запас площадей

    length = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
    length= round((fromRevitToMeters(length) * isol_reserve), 2)
    return length

def checkExpenditure(number):
    number = str(number)
    if ',' in number:
        number = number.replace(',','.')
    return number

def script_execute():
    with revit.Transaction("Добавление расходных элементов"):
        colModel = make_col(BuiltInCategory.OST_GenericModel)
        #список элементов которые будут сгенерированы
        calculation_elements = []
        insulationsTypeList = []
        insulationsObjectsList = []
        # при каждом повторе расчета удаляем старые версии
        remove_models(colModel, nameOfModel, description)


        collections = [colDuctInsul, colPipeInsul]

        #собираем типы изоляции и выписываем данные их расходников
        for collection in collections:
            for element in collection:
                elemType = doc.GetElement(element.GetTypeId())

                typeName = elemType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()

                name1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Наименование').AsString()

                mark1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Марка').AsString()
                maker1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Изготовитель').AsString()
                unit1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Ед. изм.').AsString()
                expenditure1 = checkExpenditure(elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Расход на м2').AsValueString())
                isArea1 = True
                if elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Расход по м.п.').AsInteger() == 1:
                    isArea1 = False


                name2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Наименование').AsString()
                mark2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Марка').AsString()
                maker2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Изготовитель').AsString()
                unit2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Ед. изм.').AsString()
                expenditure2 = checkExpenditure(elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Расход на м2').AsValueString())
                isArea2 = True
                if elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Расход по м.п.').AsInteger() == 1:
                    isArea2 = False


                name3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Наименование').AsString()
                mark3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Марка').AsString()
                maker3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Изготовитель').AsString()
                unit3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Ед. изм.').AsString()
                expenditure3 = checkExpenditure(elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Расход на м2').AsValueString())
                isArea3 = True
                if elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Расход по м.п.').AsInteger() == 1:
                    isArea3 = False

                toAppend = True
                for insulation in insulationsTypeList:
                    if insulation.typeName == typeName:
                        toAppend = False

                if toAppend:
                    newInsulation = insulType(typeName, name1, name2, name3, mark1, mark2, mark3,
                                                   unit1, unit2, unit3, maker1, maker2, maker3, expenditure1,
                                                   expenditure2, expenditure3, isArea1, isArea2, isArea3)
                    insulationsTypeList.append(
                        newInsulation
                     )

        #перебираем элементы изоляции, забираем функции, системы и группирование и добавляем площадь для последующих расчетов
        for collection in collections:
            for element in collection:
                elemType = doc.GetElement(element.GetTypeId())
                typeName = elemType.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()

                corp = element.LookupParameter('ФОП_Блок СМР').AsString()
                sec = element.LookupParameter('ФОП_Секция СМР').AsString()
                floor = element.LookupParameter('ФОП_Этаж').AsString()
                EF = element.LookupParameter('ФОП_Экономическая функция').AsString()
                system = element.LookupParameter('ФОП_ВИС_Имя системы').AsString()


                key = str(corp) + "_" + str(sec) + "_" + str(floor) + "_" + str(EF) + "_" + str(system) + "_" + \
                      str(typeName)

                area = get_area(element)

                length = get_length(element)


                toAppend = True
                for insulationsObject in insulationsObjectsList:
                    if insulationsObject.key == key:
                        toAppend = False
                        objectToArea = insulationsObject

                if toAppend:
                    newInsulationObject = insulPosition(corp, sec, floor, system, typeName, EF, key)
                    newInsulationObject.area = area
                    newInsulationObject.length = length
                    insulationsObjectsList.append(
                        newInsulationObject
                     )
                else:
                    objectToArea.area = objectToArea.area + area
                    objectToArea.length = objectToArea.length + length





        #взаимно перебираем объекты в спеке с типами изоляции и создаем объекты под генерацию
        for insulationsObject in insulationsObjectsList:

            for insulationType in insulationsTypeList:
                if insulationType.typeName == insulationsObject.typeName:
                    group = "12. Расходники изоляции"
                    if isToGenerate(insulationType.name1, insulationType.expenditure1):
                        number = float(insulationType.expenditure1) * float(insulationsObject.area)
                        if insulationType.isArea1:
                            areaOrLength = insulationsObject.area
                        else:
                            areaOrLength = insulationsObject.length

                        newSheduleObj = objectToGenerate(
                            corp=insulationsObject.corp,
                            sec=insulationsObject.sec,
                            floor=insulationsObject.floor,
                            system=insulationsObject.system,
                            group=group,
                            name=insulationType.name1,
                            mark=insulationType.mark1,
                            art='',
                            maker=insulationType.maker1,
                            unit=insulationType.unit1,
                            number=0,
                            mass=None,
                            comment='',
                            EF=insulationsObject.EF,
                            motherArea=areaOrLength,
                            typeName= insulationsObject.typeName,
                            expenditure = insulationType.expenditure1
                        )
                        calculation_elements.append(newSheduleObj)
                    if isToGenerate(insulationType.name2, insulationType.expenditure2):
                        if insulationType.isArea1:
                            areaOrLength = insulationsObject.area
                        else:
                            areaOrLength = insulationsObject.length

                        newSheduleObj = objectToGenerate(
                            corp=insulationsObject.corp,
                            sec=insulationsObject.sec,
                            floor=insulationsObject.floor,
                            system=insulationsObject.system,
                            group=group,
                            name=insulationType.name2,
                            mark=insulationType.mark2,
                            art='',
                            maker=insulationType.maker2,
                            unit=insulationType.unit2,
                            number=0,
                            mass=None,
                            comment='',
                            EF=insulationsObject.EF,
                            motherArea=areaOrLength,
                            typeName= insulationsObject.typeName,
                            expenditure=insulationType.expenditure2
                        )
                        calculation_elements.append(newSheduleObj)
                    if isToGenerate(insulationType.name3, insulationType.expenditure3):

                        if insulationType.isArea1:
                            areaOrLength = insulationsObject.area
                        else:
                            areaOrLength = insulationsObject.length

                        newSheduleObj = objectToGenerate(
                            corp=insulationsObject.corp,
                            sec=insulationsObject.sec,
                            floor=insulationsObject.floor,
                            system=insulationsObject.system,
                            group=group,
                            name=insulationType.name3,
                            mark=insulationType.mark3,
                            art='',
                            maker=insulationType.maker3,
                            unit=insulationType.unit3,
                            number=0,
                            mass=None,
                            comment='',
                            EF=insulationsObject.EF,
                            motherArea=areaOrLength,
                            typeName= insulationsObject.typeName,
                            expenditure=insulationType.expenditure3

                        )
                        calculation_elements.append(newSheduleObj)

        #проходимся по списку объектов под генерацию и создаем их
        new_position(calculation_elements, temporary, nameOfModel, description)

temporary = isFamilyIn(BuiltInCategory.OST_GenericModel, nameOfModel)

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

if temporary == None:
    print 'Не обнаружен якорный элемент. Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()



status = paraSpec.check_parameters()
if not status:
    anchor = checkAnchor.check_anchor(showText = False)
    if anchor:
        script_execute()


