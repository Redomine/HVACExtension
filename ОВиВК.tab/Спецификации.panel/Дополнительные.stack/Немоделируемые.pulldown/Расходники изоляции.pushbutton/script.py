#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Расчет краски и креплений'
__doc__ = "Генерирует в модели элементы с расчетом количества соответствующих материалов"


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

colPipeInsul = make_col(BuiltInCategory.OST_PipeInsulations)
colDuctInsul = make_col(BuiltInCategory.OST_DuctInsulations)


# create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_GenericModel)
collector.OfClass(FamilySymbol)
famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()


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


def remove_models(colModel):
    try:
        for element in colModel:
            edited_by = element.LookupParameter('Редактирует').AsString()
            if edited_by and edited_by != __revit__.Application.Username:
                print "Якорные элементы не были обработаны, так как были заняты пользователями:"
                print edited_by
                sys.exit()
    except Exception:
        pass
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемен(Расходники)':
            doc.Delete(element.Id)

def setElement(element, name, setting):

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



def new_position(calculation_elements):
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
            if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемен(Расходники)':
                ews = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                ews.Set(WORKSET_ID.IntegerValue)
                Models.append(element)
        except Exception:
                 print 'Не удалось присвоить рабочий набор "99_Немоделируемые элементы", проверьте список наборов'

    index = 1
    #для первого элмента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for position in calculation_elements:
        posGroup = str(position.group) + '_' + str(index)
        index+=1
        if position.group == None:
            posGroup = 'None'
        posName = position.typeName
        if position.typeName == None:
            posName = 'None'
        posMark = position.mark
        if position.mark == None:
            posMark = 'None'

        group = posGroup
                #+ posName + posMark
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
        setElement(dummy, 'ФОП_ВИС_Единица измерения', position.unit)
        setElement(dummy, 'ФОП_ВИС_Число', position.number)
        setElement(dummy, 'ФОП_ВИС_Масса', position.mass)
        setElement(dummy, 'ФОП_ВИС_Примечание', position.comment)
        setElement(dummy, 'ФОП_Экономическая функция', position.EF)

        Models.pop(0)


def isToGenerate(insTypeName, expenditure):
    if expenditure != None:
        if float(expenditure) > 0.0001:
            if insTypeName != None:
                if insTypeName != '':
                    return True



def getConnectors(element):
    connectors = []
    try:
        a = element.ConnectorManager.Connectors.ForwardIterator()
        while a.MoveNext():
            connectors.append(a.Current)
    except:
        try:
            a = element.MEPModel.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
        except:
            a = element.MEPSystem.ConnectorManager.Connectors.ForwardIterator()
            while a.MoveNext():
                connectors.append(a.Current)
    return connectors

def get_fitting_area(element):
    if element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        options = Options()
        geoms = element.get_Geometry(options)

        for g in geoms:
            solids = g.GetInstanceGeometry()
        area = 0

        for solid in solids:
            if isinstance(solid, Line) or isinstance(solid, Arc):
                continue
            for face in solid.Faces:
                area = area + face.Area

        area = fromRevitToSquareMeters(area)
        connectors = getConnectors(element)

        for connector in connectors:
            try:
                H = connector.Height
                B = connector.Width
                S = (H * B)
                S = fromRevitToSquareMeters(S)
                area = area - S
            except Exception:
                R = connector.Radius
                S = (3.14 * R * R)
                S = fromRevitToSquareMeters(S)
                area = (area - S)
        area = round(area, 2)
    return area

information = doc.ProjectInformation
isol_reserve = 1 + (lookupCheck(information, 'ФОП_ВИС_Запас изоляции').AsDouble() / 100)  # запас площадей
length_reserve = 1 + (lookupCheck(information, 'ФОП_ВИС_Запас воздуховодов/труб').AsDouble() / 100)  # запас длин

def get_area(element):
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
    length = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
    length= round((fromRevitToMeters(length) * length_reserve), 2)
    print element.Id
    print length
    return length

def script_execute():
    with revit.Transaction("Добавление расходных элементов"):
        colModel = make_col(BuiltInCategory.OST_GenericModel)
        #список элементов которые будут сгенерированы
        calculation_elements = []
        insulationsTypeList = []
        insulationsObjectsList = []
        # при каждом повторе расчета удаляем старые версии
        remove_models(colModel)


        collections = [colDuctInsul, colPipeInsul]

        #собираем типы изоляции и выписываем данные их расходников
        for collection in collections:
            for element in collection:
                elemType = doc.GetElement(element.GetTypeId())

                typeName = elemType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString()

                name1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Наименование').AsValueString()
                mark1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Марка').AsValueString()
                maker1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Изготовитель').AsValueString()
                unit1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Ед. изм.').AsValueString()
                expenditure1 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Расход на м2').AsValueString()
                isArea1 = True
                if elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 1_Расход по м.п.').AsInteger() == 1:
                    isArea1 = False


                name2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Наименование').AsValueString()
                mark2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Марка').AsValueString()
                maker2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Изготовитель').AsValueString()
                unit2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Ед. изм.').AsValueString()
                expenditure2 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Расход на м2').AsValueString()
                isArea2 = True
                if elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 2_Расход по м.п.').AsInteger() == 1:
                    isArea2 = False


                name3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Наименование').AsValueString()
                mark3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Марка').AsValueString()
                maker3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Изготовитель').AsValueString()
                unit3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Ед. изм.').AsValueString()
                expenditure3 = elemType.LookupParameter('ФОП_ВИС_Изол_Расходник 3_Расход на м2').AsValueString()
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
                typeName = elemType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString()

                corp = element.LookupParameter('ФОП_Блок СМР').AsValueString()
                sec = element.LookupParameter('ФОП_Секция СМР').AsValueString()
                floor = element.LookupParameter('ФОП_Этаж').AsValueString()
                EF = element.LookupParameter('ФОП_Экономическая функция').AsValueString()
                system = element.LookupParameter('ФОП_ВИС_Имя системы').AsValueString()


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
                    objectToArea.length = objectToArea.area + length




        #взаимно перебираем объекты в спеке с типами изоляции и создаем объекты под генерацию
        for insulationsObject in insulationsObjectsList:

            for insulationType in insulationsTypeList:
                if insulationType.typeName == insulationsObject.typeName:
                    group = "12. Расходники изоляции" + "_" + insulationsObject.typeName
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
        new_position(calculation_elements)


is_temporary_in = False

for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)


    if famsymb.Family.Name == '_Якорный элемен(Расходники)':
        temporary = famsymb
        is_temporary_in = True

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

if is_temporary_in == False:
    print 'Не обнаружен якорный элемен(металл и краска). Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()



status = paraSpec.check_parameters()
if not status:
    script_execute()

