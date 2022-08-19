# -*- coding: utf-8 -*-
import clr
import sys
import System
from System.Collections.Generic import *
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitNodes")
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

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *


def getSortGroupInd():

    sortGroupInd = []

    posInShed = False
    index = 0
    for scheduleGroupField in definition.GetFieldOrder():
        scheduleField = definition.GetField(scheduleGroupField)
        if scheduleField.GetName() == "ФОП_ВИС_Позиция":
            posInShed = True
            FOP_pos_ind = index
        index += 1

    index = 0
    for field in definition.GetFieldOrder():

        for scheduleSortGroupField in definition.GetSortGroupFields():
            if scheduleSortGroupField.FieldId.ToString() == field.ToString():
                sortGroupInd.append(index)

        index += 1

    if posInShed == False:
        print "В таблице нет параметра ФОП_ВИС_Позиция"
        sys.exit()
    return [sortGroupInd, FOP_pos_ind]



doc = __revit__.ActiveUIDocument.Document  # type: Document
uiapp = DocumentManager.Instance.CurrentUIApplication
#app = uiapp.Application
uidoc = __revit__.ActiveUIDocument
vs = doc.ActiveView
definition = vs.Definition
tData = vs.GetTableData()
tsData = tData.GetSectionData(SectionType.Header)
sectionData = vs.GetTableData().GetSectionData(SectionType.Body)

def make_col(category):
    col = FilteredElementCollector(doc) \
        .OfCategory(category) \
        .WhereElementIsNotElementType() \
        .ToElements()
    return col

sortColumnHeaders = []
sortGroupNamesInds = []
headerIndexes = []

with revit.Transaction("Запись айди"):
    definition.IsItemized = True
    definition.ShowHeaders = True
    hidden = []
    i = 0
    while i < definition.GetFieldCount():
        if definition.GetField(i).IsHidden == True:
            hidden.append(i)
        definition.GetField(i).IsHidden = False
        i += 1



    #получаем элементы на активном виде и для начала прописываем айди в позицию
    elements = FilteredElementCollector(doc, doc.ActiveView.Id)
    for element in elements:
        ADSK_Pos = element.get_Parameter(Guid('3f809907-b64c-4a8d-be5e-06709ee28386'))
        ADSK_Pos.Set(str(element.Id.IntegerValue))


startIndex = 0 #Стартовый значение для номера
with revit.Transaction("Запись номера"):
    #получаем по каким столбикам мы сортируем
    sortGroupInd = getSortGroupInd()[0] #список параметров с сортировкой
    FOP_pos_ind = getSortGroupInd()[1] #индекс столбика с позицией
    row = sectionData.FirstRowNumber
    column = sectionData.FirstColumnNumber
    oldSheduleString = None

    while row <= sectionData.LastRowNumber:
        newSheduleString = ''
        for ind in sortGroupInd:
            newSheduleString = newSheduleString + vs.GetCellText(SectionType.Body, row, ind)

        #получаем элемент по записанному айди
        elId = vs.GetCellText(SectionType.Body, row, FOP_pos_ind)

        if newSheduleString != oldSheduleString:
            if elId:
                try:
                    if int(elId):
                        startIndex += 1
                except Exception: #если вместо айди прилетает текст, пропускаем
                    pass


        try:
            pos = element = doc.GetElement(ElementId(int(elId)))
            pos.LookupParameter('ФОП_ВИС_Позиция').Set(str(startIndex))
        except Exception:
            pass

        row += 1

        oldSheduleString = newSheduleString

    definition.IsItemized = False
    definition.ShowHeaders = False

    i = 0
    while i < definition.GetFieldCount():
        if i in hidden:
            definition.GetField(i).IsHidden = True
        i += 1



    if doc.ProjectInformation.LookupParameter('ФОП_ВИС_Площади воздуховодов в примечания').AsInteger() == 1:
        colCurves = make_col(BuiltInCategory.OST_DuctCurves)

        duct_dict = {}
        for element in colCurves:
            index = element.LookupParameter('ФОП_ВИС_Позиция').AsString()
            if index not in duct_dict:
                duct_dict[index] = element.LookupParameter('ФОП_ТИП_Число').AsDouble()
            else:
                duct_dict[index] = duct_dict[index] + element.LookupParameter('ФОП_ТИП_Число').AsDouble()

        for element in colCurves:
            note = element.LookupParameter('ADSK_Примечание')
            index = element.LookupParameter('ФОП_ВИС_Позиция').AsString()
            note.Set(str(duct_dict[index])+' м²')
