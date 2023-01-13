#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = '(_в разработке)\nПривязка к помещениям'
__doc__ = "Прописывает в параметр ФОП_Помещение принадлежность элемента к помещению"


import os.path as op
import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
clr.ImportExtensions(dosymep.Revit.Geometry)

from dosymep.Bim4Everyone import *
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Revit.Geometry import *

import sys
import paraSpec
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Redomine import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output


doc = __revit__.ActiveUIDocument.Document



uidoc = __revit__.ActiveUIDocument


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
colPlumbingFixtures= make_col(BuiltInCategory.OST_PlumbingFixtures)
colSprinklers = make_col(BuiltInCategory.OST_Sprinklers)

possibleSubs = [colSprinklers, colAccessory,
               colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]

priorityCollections = [colCurves, colFlexCurves, colFlexPipeCurves, colPipeCurves, colSprinklers, colAccessory,
               colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures, colFittings, colPipeFittings]

secondPriority = [colInsulations, colPipeInsulations]

levelCol = make_col(BuiltInCategory.OST_Levels)

colView= FilteredElementCollector(doc)\
                            .OfCategory(BuiltInCategory.OST_Views).ToElements()

_options = Options(
            ComputeReferences = True,
            IncludeNonVisibleObjects = False,
            DetailLevel = ViewDetailLevel.Fine)


for view in colView:
    if str(view.ViewType) == 'ThreeD':
        viewFamilyTypeId = view.GetTypeId()
        if viewFamilyTypeId.IntegerValue != -1:
            break

class documentLink:
    def __init__(self, linkDoc, collection, transform):
        self.linkDoc = linkDoc
        self.name = self.linkDoc.Title
        self.collection = collection
        self.transform = transform


roomCols = []

roomOfThisCol = documentLink(doc, make_col(BuiltInCategory.OST_Rooms), None)
roomCols.append(roomOfThisCol)
linksCol = make_col(BuiltInCategory.OST_RvtLinks)

for link in linksCol:
    try:
        linkedDoc = link.GetLinkDocument()
        linkTransform = link.GetTransform()
        col = FilteredElementCollector(link.GetLinkDocument()) \
            .OfCategory(BuiltInCategory.OST_Rooms) \
            .WhereElementIsNotElementType() \
            .ToElements()

        roomOfLinkCol = documentLink(linkedDoc, col, linkTransform)

        roomCols.append(roomOfLinkCol)
    except:
        pass


class line:
    def __init__(self, x1, y1,z1, x2, y2, z2):
        self.x1 = x1
        self.y1 = y1
        self.z1 = z1
        self.x2 = x2
        self.y2 = y2
        self.z2 = z2


        self.linename_v1 = str(x1)+str(y1)+str(x2)+str(y2)
        self.linename_v2 = str(x2)+str(y2)+str(x1)+str(y1)

class elementPoint:
    def getElementCenter(self, lines):
        for elemLine in lines:
            self.x = (elemLine.x1 + elemLine.x2)/2
            self.y = (elemLine.y1 + elemLine.y2)/2

    def insert(self, number):
        self.FOP_room.Set(str(number))


    def __init__(self, element):
        self.roomNumber = 'None'
        self.element = element
        self.room = None
        self.FOP_room = element.LookupParameter('ФОП_Помещение')
        self.floorName = ''
        self.mid = 0
        if not 'LocationPoint' in str(element.Location):
            if self.element.Category.IsId(BuiltInCategory.OST_FlexDuctCurves):
                controlPoints = self.element.Location.Curve.ControlPoints
                controlPoint = controlPoints[1]
                self.x = controlPoint[0]
                self.y = controlPoint[1]
                self.z = controlPoint[2]
            else:
                self.elementLines = getTessallatedLine(element.Location.Curve.Tessellate(), element)
                self.z = element.Location.Curve.Origin[2]
                self.getElementCenter(self.elementLines)

        else:
            if element.HasSpatialElementCalculationPoint:
                point = element.GetSpatialElementCalculationPoint()
                self.x = point[0]
                self.y = point[1]
                self.z = point[2]
            else:
                self.elementCenter = element.Location.Point
                self.x = self.elementCenter[0]
                self.y = self.elementCenter[1]
                self.z = self.elementCenter[2]

class flatroom:
    def appendLine(self, line):
        if line.linename_v1 not in self.roomLinesNames and line.linename_v2 not in self.roomLinesNames:

            self.roomLines.append(line)
            self.roomLinesNames.append(line.linename_v1)
            self.roomLinesNames.append(line.linename_v2)
    def __init__(self, roomNumber, roomName, roomId):
        self.downBorder = 0
        self.topBorder = 0
        self.mid = 0
        self.roomLines = []
        self.roomLinesNames = []
        self.roomNumber = roomNumber
        self.roomName = roomName
        self.roomId = roomId
        self.roomLocation = 0

def getTessallatedLine(coord_list, element):
    newLines = []
    current = 'Начало линии'

    for coordinate in coord_list:
        x = coordinate[0]
        y = coordinate[1]
        z = coordinate[2]
        if current == 'Начало линии':
            x1 = x
            y1 = y
            z1 = z
            current = 'Конец линии'
        else:
            x2 = x
            y2 = y
            z2 = z
            current = 'Начало линии'
            segmentLine = line(x1, y1, z1, x2, y2, z2)

            if element.Category.IsId(BuiltInCategory.OST_Rooms):
                if segmentLine.x1 == segmentLine.x2 and segmentLine.y1 == segmentLine.y2:
                    newLines.append(None)
                else:
                    newLines.append(segmentLine)
            else:
                newLines.append(segmentLine)


    return newLines

def getBorders(element):

    box = element.get_BoundingBox(None)
    if element not in roomOfThisCol.collection:
        if box != None:
            transform = element.Document.ActiveProjectLocation.GetTransform()
            box.TransformBoundingBox(transform)
    try:
        z_min = box.Min[2]
        z_max = box.Max[2]
    except:
        z_min = 0
        z_max = 0

    return [z_min, z_max]


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



class stolPoint:
    def __init__(self, x, y):
        self.x = x
        self.y = y

# Given three collinear points p, q, r, the function checks if
# point q lies on line segment 'pr'
def onSegment(p, q, r):
    if ( (q.x <= max(p.x, r.x)) and (q.x >= min(p.x, r.x)) and
           (q.y <= max(p.y, r.y)) and (q.y >= min(p.y, r.y))):
        return True
    return False


def orientation(p, q, r):
    # to find the orientation of an ordered triplet (p,q,r)
    # function returns the following values:
    # 0 : Collinear points
    # 1 : Clockwise points
    # 2 : Counterclockwise

    # See https://www.geeksforgeeks.org/orientation-3-ordered-points/amp/
    # for details of below formula.

    val = (float(q.y - p.y) * (r.x - q.x)) - (float(q.x - p.x) * (r.y - q.y))
    if (val > 0):

        # Clockwise orientation
        return 1
    elif (val < 0):

        # Counterclockwise orientation
        return 2
    else:

        # Collinear orientation
        return 0


# The main function that returns true if
# the line segment 'p1q1' and 'p2q2' intersect.
def doIntersect(p1, q1, p2, q2):
    # Find the 4 orientations required for
    # the general and special cases
    o1 = orientation(p1, q1, p2)
    o2 = orientation(p1, q1, q2)
    o3 = orientation(p2, q2, p1)
    o4 = orientation(p2, q2, q1)

    # General case
    if ((o1 != o2) and (o3 != o4)):
        return True

    # Special Cases

    # p1 , q1 and p2 are collinear and p2 lies on segment p1q1
    if ((o1 == 0) and onSegment(p1, p2, q1)):
        return True

    # p1 , q1 and q2 are collinear and q2 lies on segment p1q1
    if ((o2 == 0) and onSegment(p1, q2, q1)):
        return True

    # p2 , q2 and p1 are collinear and p1 lies on segment p2q2
    if ((o3 == 0) and onSegment(p2, p1, q2)):
        return True

    # p2 , q2 and q1 are collinear and q1 lies on segment p2q2
    if ((o4 == 0) and onSegment(p2, q1, q2)):
        return True

    # If none of the cases
    return False


def isEquipmenInRoom(point, room):
    intersects_v1 = 0
    intersects_v2 = 0
    intersects_v3 = 0
    intersects_v4 = 0
    # print room.mid
    # print str(point.x) + '_' + str(point.y)

    for roomLine in room.roomLines:

        p1 = stolPoint(roomLine.x1, roomLine.y1)
        q1 = stolPoint(roomLine.x2, roomLine.y2)
        p2 = stolPoint(point.x, point.y)

        q2_v1 = stolPoint(point.x, point.y + 10000)
        q2_v2 = stolPoint(point.x, point.y - 10000)
        q2_v3 = stolPoint(point.x + 10000, point.y)
        q2_v4 = stolPoint(point.x - 10000, point.y)


        isV1Intersect = doIntersect(p1, q1, p2, q2_v1)
        isV2Intersect = doIntersect(p1, q1, p2, q2_v2)
        isV3Intersect = doIntersect(p1, q1, p2, q2_v3)
        isV4Intersect = doIntersect(p1, q1, p2, q2_v4)
        # if isV1Intersect or isV2Intersect or isV3Intersect or isV4Intersect:
        #     print str(roomLine.x1) + '_' + str(roomLine.y1) + '_' + str(roomLine.z1)
        #     print str(roomLine.x2) + '_' + str(roomLine.y2) + '_' + str(roomLine.z2)

        if isV1Intersect:
            intersects_v1 += 1
        if isV2Intersect:
            intersects_v2 += 1
        if isV3Intersect:
            intersects_v3 += 1
        if isV4Intersect:
            intersects_v4 += 1

    # print intersects_v1
    # print intersects_v2
    # print intersects_v3
    # print intersects_v4
    checkList = []

    if intersects_v1%2 != 0 or intersects_v2%2 != 0 or intersects_v3%2 != 0 or intersects_v4%2 != 0:
        return True

    # if intersects_v1 == 1  or intersects_v2 == 1 or intersects_v3 == 1 or intersects_v4 == 1:
    #     return True
    return  False

class level:
    def __init__(self, element):
        self.name = element.Name
        self.z_max = 0
        self.mid = 0
        self.z_min = element.ProjectElevation - 60 / 304.8


def getConnectors(element):
    try:
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
    except:
        return None

def getElementLevelName(element):
    if element.Category.IsId(BuiltInCategory.OST_PipeInsulations) \
            or element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
        connectors = getConnectors(element)
        for connector in connectors:
            for el in connector.AllRefs:
                if el.Owner.LookupParameter('Уровень'):
                    level = el.Owner.LookupParameter('Уровень').AsValueString()
                if el.Owner.LookupParameter('Базовый уровень'):
                    level = el.Owner.LookupParameter('Базовый уровень').AsValueString()
    else:
        if element.LookupParameter('Уровень'):
            level = element.LookupParameter('Уровень').AsValueString()
        if element.LookupParameter('Базовый уровень'):
            level = element.LookupParameter('Базовый уровень').AsValueString()

    return level

def isValidSecondary(owner):
    if owner.Category.IsId(BuiltInCategory.OST_MechanicalEquipment) or owner.Category.IsId(
        BuiltInCategory.OST_PipeCurves) or owner.Category.IsId(
        BuiltInCategory.OST_DuctCurves) or owner.Category.IsId(
        BuiltInCategory.OST_FlexPipeCurves) or owner.Category.IsId(
        BuiltInCategory.OST_FlexDuctCurves) or owner.Category.IsId(
        BuiltInCategory.OST_DuctAccessory) or owner.Category.IsId(
        BuiltInCategory.OST_PipeAccessory):
        return True
    else:
        return False

def isFittingsAround(element):
    if element.Category.IsId(BuiltInCategory.OST_PipeFitting) or element.Category.IsId(BuiltInCategory.OST_DuctFitting):
        connectors = getConnectors(element)
        conList = []

        for connector in connectors:
            for el in connector.AllRefs:
                if isValidSecondary(el.Owner):
                    conList.append(True)
        if len(conList) > 0:
            return True
        else:
            return False
    else:
        return False

def selectName(element):
    roomName = 'Не удалось определить имя'
    paraname = doc.ProjectInformation.LookupParameter('ФОП_Привязка к помещениям').AsString()
    if paraname == None or paraname == '':
        paraname = 'Номер'

    try:
        roomName = element.LookupParameter(paraname).AsString()
    except:
        pass

    return roomName

def rename_sub(element):
    if hasattr(element, "GetSubComponentIds"):
        floor = element.LookupParameter('ФОП_Этаж').AsString()
        room = element.LookupParameter('ФОП_Помещение').AsString()
        super_component = element.SuperComponent

        if super_component:
     	    return

        sub_elements = [doc.GetElement(element_id) for element_id in element.GetSubComponentIds()]
        for sub_element in sub_elements:
            if sub_element.Category.IsId(BuiltInCategory.OST_GenericModel):
                continue
            subFloor = sub_element.LookupParameter('ФОП_Этаж')
            subRoom = sub_element.LookupParameter('ФОП_Помещение')

            subFloor.Set(floor)
            subRoom.Set(room)

def execute():
    with revit.Transaction("Привязка к помещениям"):
        #
        # #Вид нужен потому что у уровней есть боксы только когда они видимы
        # newView = View3D.CreateIsometric(doc, viewFamilyTypeId)

        #создаем объекты уровней
        projectLevels = []
        for element in levelCol:
            newLevel = level(element)
            projectLevels.append(newLevel)

        projectLevels.sort(key=lambda x: x.z_min)

        for projectLevel in projectLevels:
            ind = projectLevels.index(projectLevel)
            try:
                projectLevel.z_max = projectLevels[ind + 1].z_min
                projectLevel.mid = (projectLevel.z_max - projectLevel.z_min) / 2 + projectLevel.z_min
            except:
                diff = projectLevel.z_min - projectLevels[ind - 1].z_min
                projectLevel.z_max = projectLevel.z_min + diff
                projectLevel.mid = (projectLevel.z_max - projectLevel.z_min) / 2 + projectLevel.z_min

            projectLevel.mid = (projectLevel.z_max - projectLevel.z_min) / 2 + projectLevel.z_min

        #Перебираем коллекции комнат и создаем на их базе объекты контуров комнат
        rooms = []
        for roomCol in roomCols:
            for element in roomCol.collection: # type: Room
                # calculator = SpatialElementGeometryCalculator(roomCol.linkDoc)
                if element.Location:

                    roomNumber =element.GetParamValue(BuiltInParameter.ROOM_NUMBER)
                    newName = selectName(element)
                    newRoom = flatroom(roomNumber, newName, element.Id)
                    roomGeom = element.get_Geometry(_options)
                    newRoom.roomLocation = element.Location.Point

                    if roomCol.transform != None:
                        roomGeom = roomGeom.GetTransformed(roomCol.transform)
                    for geom in roomGeom:
                        if 'Solid' in str(geom):
                            roomSolid = geom

                    #if newRoom.roomId.IntegerValue == 20497856:
                     #   print len(roomSolid.Faces)

                    borders = getBorders(element)
                    newRoom.downBorder = borders[0] - 60/304.8
                    newRoom.topBorder = borders[1]
                    newRoom.mid = (newRoom.topBorder - newRoom.downBorder)/2 + newRoom.downBorder


                    for face in roomSolid.Faces:
                        Loops = face.GetEdgesAsCurveLoops()
                        for loop in Loops:
                            for roomLine in loop:
                                newLines = getTessallatedLine(roomLine.Tessellate(), element)
                                for newLine in newLines:
                                    if newLine:
                                        if newLine.z1 < newRoom.mid:
                                        #if not str(round(newLine.x1, 3)) == str(round(newLine.x2, 3)) and str(round(newLine.y1, 3)) == str(round(newLine.y2, 3)):
                                            #if newLine.linename_v1 not in newRoom.roomLinesNames and newLine.linename_v2 not in newRoom.roomLinesNames:
                                            newRoom.appendLine(newLine)

                    rooms.append(newRoom)

        #Аналогично проходимся по приоритетным элементам модели и создаем объекты координатных точек
        equipmentPoints = []
        for collection in priorityCollections:
            for element in collection:
                newEquipment = elementPoint(element)
                elementLevel = getElementLevelName(element)

                for projectLevel in projectLevels:
                    if newEquipment.z > projectLevel.z_min and newEquipment.z < projectLevel.z_max:
                        newEquipment.mid = projectLevel.mid
                        newEquipment.floorName = projectLevel.name
                equipmentPoints.append(newEquipment)


        #проверяем во вторичном приоритете элементы без коннекторов, тогда обсчитываем их как обычные
        for collection in secondPriority:
            for element in collection:
                if not getConnectors(element) or isFittingsAround(element):
                    newEquipment = elementPoint(element)
                    elementLevel = getElementLevelName(element)
                    for projectLevel in projectLevels:
                        if elementLevel == str(projectLevel.name):
                            newEquipment.mid = projectLevel.mid
                    equipmentPoints.append(newEquipment)

        #перебираем координатные точки и если они попадают в помещение прописываем имя внутри экземпляра
        for equipmentPoint in equipmentPoints:
            equipmentPoint.roomNumber = 'Вне обозначенных помещений'
            #if equipmentPoint.element.Id.IntegerValue == 1686439:
            for room in rooms:
                    #if room.roomId.IntegerValue == 20497854:
                    #print room.roomLocation
                if equipmentPoint.mid > room.downBorder and equipmentPoint.mid < room.topBorder:
                        #print isEquipmenInRoom(equipmentPoint, room)
                    if isEquipmenInRoom(equipmentPoint, room):
                        equipmentPoint.roomNumber = room.roomName
                        break
                    else:
                        if equipmentPoint.roomNumber == 'None':
                            equipmentPoint.roomNumber = 'Вне обозначенных помещений'

        # вставляем нужное имя в параметр ФОП_Помещение
        for equipmentPoint in equipmentPoints:
            equipmentPoint.insert(equipmentPoint.roomNumber)

        # Запись данных
        for equipmentPoint in equipmentPoints:
            if not isElementEditedBy(equipmentPoint.element):
                fopLevel = equipmentPoint.element.LookupParameter('ФОП_Этаж')
                fopLevel.Set(equipmentPoint.floorName)



        # class connectorSearch:
        #     def reinsert(self):
        #         self.element.LookupParameter('ФОП_Помещение').Set(self.roomname)
        #
        #     def __init__(self, roomname, element):
        #         self.roomname = roomname
        #         self.element = element


        #
        # reinsertList = []
        # #Снова проходим по приоритетным элементам, ищем на коннекторах владельца
        # for collection in priorityCollections:
        #     for element in collection: # type: Element
        #         if element.Id.IntegerValue == 2040776:
        #             roomname = element.LookupParameter('ФОП_Помещение').AsString()
        #             print roomname
        #             if roomname == 'Вне обозначенных помещений':
        #                 print 1
        #                 connectors = getConnectors(element)
        #                 if connectors:
        #                     for connector in connectors:
        #                         for el in connector.AllRefs: #type: Connector
        #                             print el.Owner.LookupParameter('ФОП_Помещение').AsString()
        #                             if el.Owner.LookupParameter('ФОП_Помещение'):
        #
        #                                 elRoomName = el.Owner.LookupParameter('ФОП_Помещение').AsString()
        #
        #                                 if elRoomName != 'Вне обозначенных помещений':
        #                                     print elRoomName
        #                                     reinsertElement = connectorSearch(elRoomName, element)
        #
        #                                     # reinsertElement.element = element
        #                                     # reinsertElement.roomName = elRoomName
        #
        #                                     #
        #                                     # print reinsertElement.element
        #                                     # print reinsertElement.roomName
        #                                     # print reinsertElement
        #
        #                                     reinsertList.append(reinsertElement)
        #
        # for reinsertElement in reinsertList:
        #     reinsertElement.reinsert()
        #
        #
        # sys.exit()

        #Проходимся по второстепенным элементам модели и присваиваем им помещения приоритетных
        for collection in secondPriority:
            for element in collection:
                if getConnectors(element):
                    if not isElementEditedBy(element):
                        elementRoom = element.LookupParameter("ФОП_Помещение")
                        elementFloor = element.LookupParameter("ФОП_Этаж")
                        #try: #Пробуем перебрать через коннекторы, но их может и не быть
                        connectors = getConnectors(element)

                        for connector in connectors:
                                for el in connector.AllRefs:
                                    if isValidSecondary(el.Owner) or el.Owner.Category.IsId(BuiltInCategory.OST_PipeFitting) or el.Owner.Category.IsId(BuiltInCategory.OST_DuctFitting):

                                        if el.Owner.LookupParameter("ФОП_Этаж"):
                                            variant = el.Owner.LookupParameter("ФОП_Этаж").AsString()

                                            elementFloor.Set(str(variant))

                                        if el.Owner.LookupParameter("ФОП_Помещение"):
                                            variant = el.Owner.LookupParameter("ФОП_Помещение").AsString()

                                            elementRoom.Set(str(variant))


        #Ищем суб-элементы и если находим обрабатываем
        for collection in possibleSubs:
            for element in collection:
                if not isElementEditedBy(element):
                    rename_sub(element)






        # #стираем вид
        # doc.Delete(newView.Id)

parametersAdded = paraSpec.check_parameters()

if not parametersAdded:
    execute()
    if len(report_rows) > 0:
        for report in report_rows:
            print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report
