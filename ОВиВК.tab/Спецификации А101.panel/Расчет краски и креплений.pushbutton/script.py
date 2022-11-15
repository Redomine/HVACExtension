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
    
colPipes = make_col(BuiltInCategory.OST_PipeCurves)
colCurves = make_col(BuiltInCategory.OST_DuctCurves)
colModel = make_col(BuiltInCategory.OST_GenericModel)
colSystems = make_col(BuiltInCategory.OST_DuctSystem)
colInsul = make_col(BuiltInCategory.OST_DuctInsulations)
# create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_GenericModel)
collector.OfClass(FamilySymbol)
famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()


class settings:
    def __init__(self,
                 group,
                 name,
                 unit,
                 param_name,
                 collection,
                 is_type
    ):
        pass


#settings('8. Расчетные элементы', "Металлические крепления для воздуховодов", 'кг.', 'ФОП_ВИС_Расчет металла для креплений', colCurves, False)

generation_list = [
    # для добавления элементов внести в список на генерацию по форме ниже
    # Если число формируется по отдельному алгоритму - добавляем формулу в функцию, иначе выпадет 1
    # [Группирование, Имя, Единица измерения, ФОП_Имя параметра, Коллекция, Один элемент на систему?]
    ['12. Расчетные элементы', "Металлические крепления для воздуховодов", 'кг.', 'ФОП_ВИС_Расчет металла для креплений', colCurves, False],
    ['12. Расчетные элементы', "Металлические крепления для трубопроводов", 'кг.', 'ФОП_ВИС_Расчет металла для креплений', colPipes, False],
    ['12. Расчетные элементы', "Изоляция для фланцев и стыков", 'м².', 'ФОП_ВИС_Совместно с воздуховодом', colInsul, False],
    ['12. Расчетные элементы', "Краска антикоррозионная за два раза БТ-177", 'кг.', 'ФОП_ВИС_Расчет краски и грунтовки', colPipes, False],
    ['12. Расчетные элементы', "Грунт ГФ-031", 'кг.', 'ФОП_ВИС_Расчет краски и грунтовки', colPipes, False],
    ['12. Расчетные элементы', 'Комплект заделки отверстий с восстановлением предела огнестойкости', 'компл.', 'ФОП_ВИС_Комплекты заделки', colCurves, True],
    ['12. Расчетные элементы', 'Комплект заделки отверстий с восстановлением предела огнестойкости', 'компл.', 'ФОП_ВИС_Комплекты заделки', colPipes, True]
]

def setElement(element, name, setting):
    if name == "ФОП_ВИС_Масса":
        element.LookupParameter(name).Set(str(setting))
    if name == 'ADSK_Единица измерения':
        element.LookupParameter('ФОП_ТИП_Единица измерения').Set(str(setting))
        element.LookupParameter('ФОП_ВИС_Единица измерения').Set(str(setting))
    try:
        if setting == None:
            pass
        else:
            element.LookupParameter(name).Set(setting)
            if name == 'ФОП_ВИС_Число':
                element.LookupParameter('ФОП_ТИП_Число').Set(str(setting))
            if name == 'ФОП_ВИС_Наименование комбинированное':
                element.LookupParameter('ФОП_ТИП_Назначение').Set(setting)
    except Exception:
        pass

def get_ADSK_Name(element):
    if element.LookupParameter('ADSK_Наименование'):
        ADSK_Name = element.LookupParameter('ADSK_Наименование').AsString()
        if ADSK_Name == None or ADSK_Name == "":
            ADSK_Name = "None"
    else:
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)

        if ElemType.LookupParameter('ADSK_Наименование') == None\
                or ElemType.LookupParameter('ADSK_Наименование').AsString() == None \
                or ElemType.LookupParameter('ADSK_Наименование').AsString() == "":
            ADSK_Name = "None"
        else:
            ADSK_Name = ElemType.LookupParameter('ADSK_Наименование').AsString()

    if str(element.Category.Name) == 'Трубы':
        ElemTypeId = element.GetTypeId()
        ElemType = doc.GetElement(ElemTypeId)
        if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента'):
            if ElemType.LookupParameter('ФОП_ВИС_Имя трубы из сегмента').AsInteger() == 1:
                ADSK_Name = element.LookupParameter('Описание сегмента').AsString()



    return ADSK_Name

def duct_material(element):
    area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903) / 100
    if element.GetParamValue(BuiltInParameter.RBS_EQ_DIAMETER_PARAM) == element.GetParamValue(BuiltInParameter.RBS_HYDRAULIC_DIAMETER_PARAM):
        D = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        P = 3.14 * D
    else:
        A = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        B = 304.8 * element.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        P = 2 * (A + B)

    if P < 1001:
        kg = area * 65
    elif P < 1801:
        kg = area * 122
    else:
        kg = area * 225

    return kg

def pipe_material(element):
    lenght = (304.8 * element.GetParamValue(BuiltInParameter.CURVE_ELEM_LENGTH))/1000

    D = 304.8 * element.GetParamValue(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)

    if D < 25:
        kg = 0.9 * lenght
    elif D < 33:
        kg = 0.73 * lenght
    elif D < 41:
        kg = 0.64 * lenght
    elif D < 51:
        kg = 0.67 * lenght
    elif D < 66:
        kg = 0.53 * lenght
    elif D < 81:
        kg = 0.7 * lenght
    elif D < 101:
        kg = 0.64 * lenght
    elif D < 126:
        kg = 1.16 * lenght
    else:
        kg = 0.96 * lenght
    return kg

def insul_stock(element):
    area = element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA)
    if area == None:
        area = 0
    area = area * 0.092903 * 0.03
    return area

def grunt(element):
    area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
    number = area / 10
    return number

def colorBT(element):
    area = (element.GetParamValue(BuiltInParameter.RBS_CURVE_SURFACE_AREA) * 0.092903)
    number = area * 0.2 * 2
    return number

class calculation_element:
    def get_number(self, element, name):
        Number = 1
        if name == "Металлические крепления для трубопроводов" and element in colPipes:
            Number = pipe_material(element)
        if name == "Металлические крепления для воздуховодов" and element in colCurves:
            Number = duct_material(element)
        if name == "Изоляция для фланцев и стыков" and element in colInsul:
            Number = insul_stock(element)
        if name == "Краска антикоррозионная за два раза БТ-177" and element in colPipes:
            Number = colorBT(element)
        if name == "Грунт ГФ-031" and element in colPipes:
            Number = grunt(element)

        return Number

    def __init__(self, element, collection, parameter, Name):
        self.System = 'None'
        self.Class = 'None'
        self.Work = 'None'
        self.Group = '8. Расчетные элементы'
        self.Name = Name
        self.Mark = ''
        self.Art = ''
        self.Maker = ''
        self.Izm = 'None'
        self.Number = 0
        self.Mass = ''
        self.Comment = ''
        self.EF = 'None'
        self.parentId = element.Id.IntegerValue

        for generation in generation_list:
            if collection in generation and parameter in generation:
                self.Izm = generation[2]
                isType = generation[5]

        self.EF = str(element.LookupParameter('ФОП_Экономическая функция').AsString())
        self.System = str(element.LookupParameter('ADSK_Имя системы').AsString())
        if parameter == 'ФОП_ВИС_Совместно с воздуховодом':
            pass

        self.Number = self.get_number(element, Name)

        elemType = doc.GetElement(element.GetTypeId())
        if element in colInsul and elemType.LookupParameter('ФОП_ВИС_Совместно с воздуховодом').AsInteger() == 1:
            self.Name = 'Изоляция ' + get_ADSK_Name(element) + ' для фланцев и стыков'


def new_position(calculation_elements):
    #ищем стадию спеки для присвоения
    phaseOk = False
    for phase in doc.Phases:
        if phase.Name == 'Спецификация':
            phaseid = phase.Id
            phaseOk = True

    if not phaseOk:
        print 'Не удалось присвоить стадию спецификация, проверьте список стадий'
        sys.exit()

    # создаем заглушки по элементов собранных из таблицы
    loc = XYZ(0, 0, 0)

    temporary.Activate()
    for element in calculation_elements:
        familyInst = doc.Create.NewFamilyInstance(loc, temporary, Structure.StructuralType.NonStructural)
    # собираем список из созданных заглушек
    colModel = make_col(BuiltInCategory.OST_GenericModel)
    Models = []
    for element in colModel:
        if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемен(металл и краска)':
            element.CreatedPhaseId = phaseid


            Models.append(element)

    # для первого элмента списка заглушек присваиваем все параметры, после чего удаляем его из списка
    for element in calculation_elements:
        group = str(element[3]) + str(element[4]) + str(element[5])
        dummy = Models[0]
        setElement(dummy, 'ADSK_Имя системы', element[0])
        setElement(dummy, 'ФОП_ТИП_Код', element[1])
        setElement(dummy, 'ФОП_ТИП_Наименование работы', element[2])
        setElement(dummy, 'ФОП_ВИС_Группирование', group)
        setElement(dummy, 'ФОП_ВИС_Наименование комбинированное', element[4])
        setElement(dummy, 'ADSK_Марка', element[5])
        setElement(dummy, 'ADSK_Код изделия', element[6])
        setElement(dummy, 'ADSK_Завод-изготовитель', element[7])
        setElement(dummy, 'ADSK_Единица измерения', element[8])
        setElement(dummy, 'ФОП_ВИС_Число', element[9])
        setElement(dummy, 'ФОП_ВИС_Масса', element[10])
        setElement(dummy, 'ADSK_Примечание', element[11])
        setElement(dummy, 'ФОП_Экономическая функция', element[12])
        try:
            setElement(dummy, 'ID Родительских элементов', element[13])
        except:
            pass

        Models.pop(0)

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
        if element.LookupParameter('Семейство').AsValueString() == '_Якорный элемен(металл и краска)':
            doc.Delete(element.Id)
def is_object_to_generate(element, genCol, collection, parameter, generation_list = generation_list):
    if element in genCol:
        for generation in generation_list:
            if collection in generation and parameter in generation:
                try:
                    elemType = doc.GetElement(element.GetTypeId())
                    if elemType.LookupParameter(parameter).AsInteger() == 1:
                        return True
                except Exception:
                    if element.LookupParameter(parameter).AsInteger() == 1:
                        return True

def collapse_list(lists):
    singles = []
    for gen in generation_list:
        name = gen[1]
        isSingle = gen[5]
        if isSingle:
            singles.append(name)

    dict = {}

    for list in lists:
        system = list[0]
        EF = list[12]
        parentId = list[13]
        name = list[4]
        number = list[9]
        if name in singles:
            number = 1
        Key = EF + "_" + system + "_" + name

        if Key not in dict:
            dict[Key] = list
        else:
            dict[Key][9] = dict[Key][9] + number
            if name in singles:
                dict[Key][9] = 1
            dict[Key][13] = str(dict[Key][13]) + ';' + str(parentId)

    collapsed_list = []
    for x in dict:
        collapsed_list.append(dict[x])
    return collapsed_list


def script_execute():
    with revit.Transaction("Добавление расчетных элементов"):
        # при каждом повторе расчета удаляем старые версии
        remove_models(colModel)

        #список элементов которые будут сгенерированы
        calculation_elements = []

        collections = [colInsul, colPipes, colCurves]

        #тут мы перебираем элементы из коллекций по дурацкому алгоритму соответствия списку параметризации
        for collection in collections:
            for element in collection:
                elemType = doc.GetElement(element.GetTypeId())
                for generation in generation_list:
                    name = generation[1]
                    parameter = generation[3]
                    genCol = generation[4]

                    if is_object_to_generate(element, genCol, collection, parameter):

                        definition = calculation_element(element, collection, parameter, name)
                        definitionList = [definition.System, definition.Class, definition.Work,
                                          definition.Group, definition.Name, definition.Mark,
                                          definition.Art, definition.Maker, definition.Izm, definition.Number,
                                          definition.Mass, definition.Comment, definition.EF, definition.parentId]

                        calculation_elements.append(definitionList)

        calculation_elements = collapse_list(calculation_elements)
        new_position(calculation_elements)

is_temporary_in = False

for element in famtypeitr:
    famtypeID = element
    famsymb = doc.GetElement(famtypeID)


    if famsymb.Family.Name == '_Якорный элемен(металл и краска)':
        temporary = famsymb
        is_temporary_in = True
        print element


if is_temporary_in == False:
    print 'Не обнаружен якорный элемен(металл и краска). Проверьте наличие семейства или восстановите исходное имя.'
    sys.exit()


status = paraSpec.check_parameters()
if not status:
    script_execute()

