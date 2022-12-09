# -*- coding: utf-8 -*-


__title__ = '(_в разработке)\nПеренос значений'
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

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

class paramCell:
    def __init__(self, paraIndex,sortGroupInd ,sortname):
        self.index = paraIndex
        self.sortGroupInd = sortGroupInd
        self.name = sortname


report_rows = []
def isElementEditedBy(element):
    try:
        edited_by = element.GetParamValue(BuiltInParameter.EDITED_BY)
    except Exception:
        edited_by = __revit__.Application.Username

    if edited_by and edited_by != __revit__.Application.Username:
        if edited_by not in report_rows:
            report_rows.add(edited_by)
        return True
    return False




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
        index += 1

    index = 0
    for field in definition.GetFieldOrder():
        for scheduleSortGroupField in definition.GetSortGroupFields():
            if scheduleSortGroupField.FieldId.ToString() == field.ToString():
                sortGroupInd = index
        index += 1

    try:
        param = paramCell(paraIndex, sortGroupInd, paraName)
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

        elementsOnView = FilteredElementCollector(doc, doc.ActiveView.Id)
        try:
            for element in elementsOnView:
                if not isElementEditedBy(element):
                    position = element.get_Parameter(Guid('3f809907-b64c-4a8d-be5e-06709ee28386'))
                    position.Set(str(element.Id.IntegerValue))
                else:
                    pass
        except:
            print 'Не удалось обратотать параметр ФОП_ВИС_Позиция для элементов в спецификации. Проверьте, назначен ли он для них'


    with revit.Transaction("Перенос параметров"):

        #перебираем элементы на активном виде и для начала прописываем айди в позицию
        parameters = getParamsInShed(definition)
        #for element in elementsOnView:

        startParamName = SelectFromList('Выберите исходный параметр', parameters)
        endParamName = SelectFromList('Выберите целевой параметр', parameters)

        if startParamName == endParamName:
            print 'Имя исходного и целевого параметра одинаковое'
            sys.exit()

        paraObj = getParaInd(startParamName, definition)
        posParaObj = getParaInd('ФОП_ВИС_Позиция', definition)



        row = sectionData.FirstRowNumber
        while row <= sectionData.LastRowNumber:

            try:  # могут быть неправильные номера строк, заголовки там и тд - пропускаем их
                elId = vs.GetCellText(SectionType.Body, row, posParaObj.index)
                sheduleElement = doc.GetElement(ElementId(int(elId)))

                try:
                    startParamValue = sheduleElement.GetParamValue(startParamName)
                except:
                    startParamValue = vs.GetCellText(SectionType.Body, row, paraObj.index)

                targetParam = sheduleElement.LookupParameter(endParamName)

                if str(targetParam.StorageType) == 'Double':
                    if startParamValue == None: startParamValue = 0
                    targetParam.Set(float(startParamValue))
                if str(targetParam.StorageType) == 'Integer':
                    if startParamValue == None: startParamValue = 0
                    targetParam.Set(int(startParamValue))
                if str(targetParam.StorageType) == 'String':
                    if startParamValue == None: startParamValue = ''
                    targetParam.Set(str(startParamValue))

                #element.SetParamValue(targetParam, startParamValue)
                position = sheduleElement.get_Parameter(Guid('3f809907-b64c-4a8d-be5e-06709ee28386'))
                position.Set(str(''))
            except:
                pass
            row+=1


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

try:
    if vs.Category.IsId(BuiltInCategory.OST_Schedules):
        vsShedule = True
    else:
        vsShedule = False
except:
    vsShedule = False

if vsShedule:
    execute()
    if len(report_rows) > 0:
        for report in report_rows:
            print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report
else:
    print "Применяйте скрипт на активном виде спецификации"
