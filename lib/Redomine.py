#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Библиотека стандартных функций'
__doc__ = "pass"

import os.path as op
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Revit.Geometry import *
import sys
import paraSpec
from Autodesk.Revit.DB import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output
from pyrevit import script


doc = __revit__.ActiveUIDocument.Document  # type: Document
import os.path as op
import clr

output = script.get_output()


def getParameter(element, paraName):
    try:
        parameter = element.LookupParameter(paraName)
        return parameter
    except:
        return None

def setIfNotRO(parameter, value):
    if not parameter.IsReadOnly:
        parameter.Set(value)



def getDefCols():
    colFittings = make_col(BuiltInCategory.OST_DuctFitting)
    colPipeFittings = make_col(BuiltInCategory.OST_PipeFitting)
    colPipeCurves = make_col(BuiltInCategory.OST_PipeCurves)
    colCurves = make_col(BuiltInCategory.OST_DuctCurves)
    colFlexCurves = make_col(BuiltInCategory.OST_FlexDuctCurves)
    colFlexPipeCurves = make_col(BuiltInCategory.OST_FlexPipeCurves)
    colTerminals = make_col(BuiltInCategory.OST_DuctTerminal)
    colAccessory = make_col(BuiltInCategory.OST_DuctAccessory)
    colPipeAccessory = make_col(BuiltInCategory.OST_PipeAccessory)
    colEquipment = make_col(BuiltInCategory.OST_MechanicalEquipment)
    colInsulations = make_col(BuiltInCategory.OST_DuctInsulations)
    colPipeInsulations = make_col(BuiltInCategory.OST_PipeInsulations)
    colPlumbingFixtures = make_col(BuiltInCategory.OST_PlumbingFixtures)
    colSprinklers = make_col(BuiltInCategory.OST_Sprinklers)

    collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves, colInsulations,
                   colPipeInsulations, colPipeCurves, colSprinklers, colAccessory,
                   colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]

    return collections




output = script.get_output()

def isItFamily():
    try:
        manager = doc.FamilyManager
        return True
    except Exception:
        pass

def lookupCheck(element, paraName, isExit = True):
    type = 'экземпляра '
    try:
        if element.GetTypeId():
            type = 'типа '
    except:
        pass

    parameter = getParameter(element, paraName)

    if parameter:
        return parameter
    else:
        if isExit:
            print 'Параметр ' + type + paraName + ' не назначен для категории ' + element.Category.Name + ' (ID элемента на котором найдена ошибка ' + output.linkify(element.Id) +")"
            sys.exit()
        else:
            return None
def getSharedParameter (element, paraName, replaceName):
    if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
        if paraName == 'ADSK_Наименование':
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)
            if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента').AsInteger() == 1:

                parameter = element.LookupParameter('Описание сегмента').AsString()
                return parameter

    report = paraName.split('_')[1]
    replaceDef = isReplasing(replaceName)
    name = paraName
    if replaceDef:
            name = replaceIsValid(element, paraName, replaceName)
    try:
        parameter = element.LookupParameter(name).AsString()
        if parameter == None:
            parameter = 'None'
    except Exception:
        #try:

        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter(name) == None:
            parameter = "None"
        else:
            parameter = ElemType.LookupParameter(name).AsString()

    nullParas = ['ADSK_Завод-изготовитель', 'ADSK_Марка', 'ADSK_Код изделия']

    if not parameter:
        parameter = 'None'

    if parameter == 'None' or parameter == None:
        if paraName in nullParas:
            parameter = ''
        else:
            parameter = 'None'

    parameter = str(parameter)

    return parameter

def fromRevitToMeters(number):
    meters = (number * 304.8)/1000
    return meters

def fromRevitToMilimeters(number):
    mms = number * 304.8
    return mms
def fromRevitToSquareMeters(number):
    square = number * 0.092903
    return square


#создает коллекцию из экземпляров категории
def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

def isReplasing(replaceName ):
    replaceDefiniton = doc.ProjectInformation.LookupParameter(replaceName).AsString()
    if replaceDefiniton == None:
        return False
    if replaceDefiniton == 'None':
        return False
    if replaceDefiniton == '':
        return False
    else:
        return replaceDefiniton

def replaceIsValid (element, paraName, replaceName):
    replaceparameter = doc.ProjectInformation.LookupParameter(replaceName).AsString()
    if not element.LookupParameter(replaceparameter):
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if not ElemType.LookupParameter(replaceparameter):
            print output.linkify(element.Id)
            print 'Назначеного параметра замены ' + replaceparameter + ' нет у одной из категорий, назначенных для исходного параметр ' + paraName
            sys.exit()

    return replaceparameter

def get_ADSK_Izm(element):
    paraName = 'ADSK_Единица измерения'
    replaceName = 'ФОП_ВИС_Замена параметра_Единица измерения'
    ADSK_Izm = getSharedParameter(element, paraName, replaceName)
    return ADSK_Izm

def get_ADSK_Name(element):
    paraName = 'ADSK_Наименование'
    replaceName = 'ФОП_ВИС_Замена параметра_Наименование'
    ADSK_Name = getSharedParameter(element, paraName, replaceName)
    return ADSK_Name

def get_ADSK_Maker(element):
    paraName = 'ADSK_Завод-изготовитель'
    replaceName = 'ФОП_ВИС_Замена параметра_Завод-изготовитель'
    ADSK_Maker = getSharedParameter(element, paraName, replaceName)
    return ADSK_Maker

def get_ADSK_Mark(element):
    paraName = 'ADSK_Марка'
    replaceName = 'ФОП_ВИС_Замена параметра_Марка'
    ADSK_Mark = getSharedParameter(element, paraName, replaceName)
    return ADSK_Mark

def get_ADSK_Code(element):
    paraName = 'ADSK_Код изделия'
    replaceName = 'ФОП_ВИС_Замена параметра_Код изделия'
    ADSK_Code = getSharedParameter(element, paraName, replaceName)
    return ADSK_Code

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


#возвращает площадь фитинга воздуховода
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


#заполняет ячейки в сгенерированном немоделируемом
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


#генерирует пустые элементы в рабочем наборе немоделируемых
def new_position(calculation_elements, temporary, famName, description):
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
            if element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() == famName:
                if element.LookupParameter('ФОП_ВИС_Назначение').AsValueString() == '':
                    ews = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                    ews.Set(WORKSET_ID.IntegerValue)
                    Models.append(element)
        except Exception:
                 print 'Не удалось присвоить рабочий набор "99_Немоделируемые элементы", проверьте список наборов'

    index = 1
    #для первого элмента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for position in calculation_elements:
        group = position.group
        if description != 'Пустая строка':
            posGroup = str(position.group) + '_' + str(position.name) + '_' + str(position.mark) + '_' + str(index)
            index+=1
            group = posGroup

        if description != 'Пустая строка':
            dummy = Models[0]
        else:
            dummy = familyInst
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
        setElement(dummy, 'ФОП_ВИС_Назначение', description)
        Models.pop(0)

#для прогона новых ревизий генерации немоделируемых: стирает элмент с переданным именем модели
def remove_models(colModel, famName, description):
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
        if element.LookupParameter('ФОП_ВИС_Назначение'):
            currentName = element.LookupParameter('Семейство').AsValueString()
            currentDescription = element.LookupParameter('ФОП_ВИС_Назначение').AsValueString()
        if  currentName == famName:
            if currentDescription == description:
                doc.Delete(element.Id)

#класс содержащий все ячейки типовой спецификации
class rowOfSpecification:
    def __init__(self, corp, sec, floor, system, group, name, mark, art, maker, unit, number, mass, comment, EF):
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
        self.number = number
        self.mass = mass
        self.comment = comment
        self.EF = EF

#возвращает famsymbol если семейство есть в проекте, None если нет
def isFamilyIn(builtin, name):
    # create a filtered element collector set to Category OST_Mass and Class FamilySymbol
    collector = FilteredElementCollector(doc)
    collector.OfCategory(builtin)
    collector.OfClass(FamilySymbol)
    famtypeitr = collector.GetElementIdIterator()
    famtypeitr.Reset()

    is_temporary_in = False
    for element in famtypeitr:
        famtypeID = element
        famsymb = doc.GetElement(famtypeID)

        if famsymb.Family.Name == name:
            temporary = famsymb
            is_temporary_in = True
            return temporary
    else:
        return None