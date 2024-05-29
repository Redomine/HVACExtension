# -*- coding: utf-8 -*-


__title__ = 'Перенос значений'
__doc__ = "Переносит между собой значения параметров в активной спецификации"
import clr
import sys
import System
from System.Collections.Generic import *
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitNodes")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from pyrevit import revit
from rpw.ui.forms import SelectFromList
from System import Guid
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Revit import ParamExtensions
from Redomine import *
from pyrevit import forms

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import paraSpec

class paramCell:
    def __init__(self, paraIndex,sortGroupInd ,sortname):
        self.index = paraIndex
        self.sortGroupInd = sortGroupInd
        self.name = sortname
        self.unitType = ''
        self.displayUnitType = ''

class projectParam:
    def __init__(self, name, unit):
        self.name = name
        self.unit = unit
report_rows = set()




def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

def get_duct_area(element):
    length_reserve = 1 + (doc.ProjectInformation.LookupParameter(
        'ФОП_ВИС_Запас воздуховодов/труб').AsDouble() / 100)  # запас длин
    if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
        fop_number = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903) * length_reserve
        fop_number = round(fop_number, 2)
    return fop_number

def getParaInd(paraName, definition):
    sortGroupInd = []
    index = 0

    for scheduleGroupField in definition.GetFieldOrder():
        scheduleField = definition.GetField(scheduleGroupField)
        if scheduleField.GetName() == paraName:
            paraIndex = index
            paraFormat = scheduleField.GetFormatOptions() # type: FormatOptions
            paraType = None
            if not paraFormat.UseDefault:
                try:
                    paraType = paraFormat.GetUnitTypeId()
                except:
                    paraType = paraFormat.DisplayUnits
        index += 1

    index = 0
    for field in definition.GetFieldOrder():
        for scheduleSortGroupField in definition.GetSortGroupFields():
            if scheduleSortGroupField.FieldId.ToString() == field.ToString():
                sortGroupInd = index
        index += 1


    try:
        param = paramCell(paraIndex, sortGroupInd, paraName)
        param.unitType = paraType

    except:
        print 'Параметр ' + paraName + ' не обнаружен в таблице'
        sys.exit()
    return param


def getParamsInShed(definition):
    paramList = []
    for scheduleGroupField in definition.GetFieldOrder():
        scheduleField = definition.GetField(scheduleGroupField)
        paramList.append(scheduleField.GetName())
    return paramList

def isNoneUnitType(element, parameterObj):
    if parameterObj.unitType == None:
        targetParam = element.LookupParameter(parameterObj.name)
        if not targetParam:
            ElemTypeId = element.GetTypeId()
            ElemType = doc.GetElement(ElemTypeId)
            targetParam = ElemType.LookupParameter(parameterObj.name)

        if targetParam:
            definition = targetParam.Definition
            if not str(targetParam.StorageType) == 'String':
                try:
                    unit = definition.UnitType
                except:
                    unit = targetParam.GetUnitTypeId()

                parameterObj.unitType = unit






errorList = []
projectParams = []
def execute():
    uiapp = DocumentManager.Instance.CurrentUIApplication
    # app = uiapp.Application
    uidoc = __revit__.ActiveUIDocument

    definition = vs.Definition
    tData = vs.GetTableData()
    tsData = tData.GetSectionData(SectionType.Header)
    sectionData = vs.GetTableData().GetSectionData(SectionType.Body)

    sortColumnHeaders = []
    sortGroupNamesInds = []
    headerIndexes = []
    report_rows = set()


    with revit.Transaction("Запись айди"):
        elementsOnView = FilteredElementCollector(doc, doc.ActiveView.Id)
        try:
            for element in elementsOnView:
                if not isElementEditedBy(element):
                    position = element.get_Parameter(Guid('3f809907-b64c-4a8d-be5e-06709ee28386'))
                    position.Set(str(element.Id.IntegerValue))
                else:
                    pass
        except:
            print 'Не удалось обработотать параметр ФОП_ВИС_Позиция для элементов в спецификации. Проверьте, назначен ли он для них'
            sys.exit()

        #перебираем элементы на активном виде и для начала прописываем айди в позицию
        parameters = getParamsInShed(definition)
        #for element in elementsOnView:

        startParamName = forms.SelectFromList.show(parameters, title="Выберите исходный параметр", button_name='Применить')
        endParamName = forms.SelectFromList.show(parameters, title="Выберите целевой параметр", button_name='Применить')

        if startParamName == endParamName:
            print 'Имя исходного и целевого параметра одинаковое'
            sys.exit()

        rollback_itemized = False
        rollback_header = False

        # если заголовки показаны изначально или если спека изначально развернута - сворачивать назад не нужно
        if definition.IsItemized == False:
            rollback_itemized = True
        definition.IsItemized = True

        if definition.ShowHeaders == False:
            rollback_header = True
        definition.ShowHeaders = True

        hidden = []
        i = 0
        while i < definition.GetFieldCount():
            if definition.GetField(i).IsHidden == True:
                hidden.append(i)
            definition.GetField(i).IsHidden = False
            i += 1

    with revit.Transaction("Перенос параметров"):
        paraObj = getParaInd(startParamName, definition)
        endParaObj = getParaInd(endParamName, definition)
        posParaObj = getParaInd('ФОП_ВИС_Позиция', definition)

        row = sectionData.FirstRowNumber
        while row <= sectionData.LastRowNumber:

            try:  # могут быть неправильные номера строк, заголовки там и тд - пропускаем их
                elId = vs.GetCellText(SectionType.Body, row, posParaObj.index)
                sheduleElement = doc.GetElement(ElementId(int(elId)))
            except:
                row+=1
                continue

            isNoneUnitType(sheduleElement, paraObj)
            isNoneUnitType(sheduleElement, endParaObj)


            try:
                startParamValue = sheduleElement.GetParamValue(startParamName)
            except:
                startParamValue = vs.GetCellText(SectionType.Body, row, paraObj.index)


            try:
                startParamValue = UnitUtils.ConvertFromInternalUnits(startParamValue, paraObj.unitType)
            except:
                pass

            targetParam = sheduleElement.LookupParameter(endParamName)
            if not targetParam:
                ElemTypeId = sheduleElement.GetTypeId()
                ElemType = doc.GetElement(ElemTypeId)
                targetParam = ElemType.LookupParameter(endParamName)
            if not targetParam:
                print 'У элементов спецификации не существует целевого параметра. Возможно вы выбрали расчетное значение.'
                row+=1
                continue

            if targetParam.IsReadOnly:
                error = 'Целевой параметр недоступен для редактирования'
                if error not in errorList:
                    errorList.append(error)
                row+=1
                continue

            if str(targetParam.StorageType) == 'Double':
                if startParamValue == None or startParamValue == '': startParamValue = 0
                startParamValue = UnitUtils.ConvertToInternalUnits(float(startParamValue), endParaObj.unitType)
                targetParam.Set(float(startParamValue))
            if str(targetParam.StorageType) == 'Integer':
                if startParamValue == None or startParamValue == '': startParamValue = 0
                targetParam.Set(int(startParamValue))
            if str(targetParam.StorageType) == 'String':
                if startParamValue == None: startParamValue = ''
                targetParam.Set(str(startParamValue))
            row+=1

        for element in elementsOnView:
            position = element.get_Parameter(Guid('3f809907-b64c-4a8d-be5e-06709ee28386'))
            position.Set('')

        for error in errorList:
            print error

        if rollback_itemized == True:
            definition.IsItemized = False

        if rollback_header == True:
            definition.ShowHeaders = False

        i = 0
        while i < definition.GetFieldCount():
            if i in hidden:
                definition.GetField(i).IsHidden = True
            i += 1


doc = __revit__.ActiveUIDocument.Document  # type: Document
vs = doc.ActiveView

if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

try:
    if vs.Category.IsId(BuiltInCategory.OST_Schedules):
        vsShedule = True
    else:
        vsShedule = False
except:
    vsShedule = False



if vsShedule:
    status = paraSpec.check_parameters()
    if not status:
        execute()
        if len(report_rows) > 0:
            for report in report_rows:
                print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report
else:
    print "Применяйте скрипт на активном виде спецификации"
