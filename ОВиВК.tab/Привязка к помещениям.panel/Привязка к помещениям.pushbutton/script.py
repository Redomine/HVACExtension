#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Привязка к помещениям'
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

collections = [colFittings, colPipeFittings, colCurves, colFlexCurves, colFlexPipeCurves,  colInsulations, colPipeInsulations, colPipeCurves, colSprinklers, colAccessory,
               colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]



roomCols = []

roomOfThisCol = make_col(BuiltInCategory.OST_Rooms)
roomCols.append(roomOfThisCol)

linksCol = make_col(BuiltInCategory.OST_RvtLinks)

for link in linksCol:
    try:
        col = FilteredElementCollector(link.GetLinkDocument()) \
            .OfCategory(BuiltInCategory.OST_Rooms) \
            .WhereElementIsNotElementType() \
            .ToElements()
        roomCols.append(col)
    except:
        pass



class line:
    def __init__(self, x1, y1, x2, y2):
        self.x1 = x1
        self.y1 = y1
        self.x2 = x2
        self.y2 = y2


def getElementLevel(element):
    pass

class elementPoint:
    def getElementCenter(self, line):
        self.x = (line.x1 + line.x2)/2
        self.y = (line.y1 + line.y2)/2

    def insert(self, number):
        self.FOP_room.Set(number)

    def __init__(self, element):
        self.element = element
        self.room = None
        self.FOP_room = element.LookupParameter('ФОП_Помещение')
        self.pointLevel = getElementLevel(self.element)
        #try:
        if not 'LocationPoint' in str(element.Location):
            self.elementLines = getTessallatedLine(element.Location.Curve.Tessellate())
            self.elementCenter = self.getElementCenter(self.elementLines[0])
        else:
            self.elementCenter = element.Location.Point
            self.x = self.elementCenter[0]
            self.y = self.elementCenter[1]

class flatroom:
    def appendLine(self, lines):
        for line in lines:
            self.roomLines.append(line)

    def __init__(self, roomNumber, roomName, roomId):
        self.downBorder = 0
        self.topBorder = 0
        self.roomLines = []
        self.roomNumber = roomNumber
        self.roomName = roomName
        self.roomId = roomId
def getTessallatedLine(coord_list):
    segmentLines =[]
    current = 'Начало линии'
    for coordinate in coord_list:
        x = coordinate[0]
        y = coordinate[1]
        if current == 'Начало линии':
            x1 = x
            y1 = y
            current = 'Конец линии'
        else:
            x2 = x
            y2 = y
            current = 'Начало линии'
            segmentLine = line(x1, y1, x2, y2)
            segmentLines.append(segmentLine)
    return segmentLines

def getBorders(element):

    box = element.get_BoundingBox(None)
    if element not in roomOfThisCol:
        transform = element.Document.ActiveProjectLocation.GetTransform()

    box.TransformBoundingBox(transform)
    try:
        z_min = box.Min[2]
        z_max = box.Max[2]
    except:
        z_min = 0
        z_max = 0

    return [z_min, z_max]



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




        if isV1Intersect:
            intersects_v1 += 1
        if isV2Intersect:
            intersects_v2 += 1
        if isV3Intersect:
            intersects_v3 += 1
        if isV4Intersect:
            intersects_v4 += 1

    if intersects_v1 == 1 or intersects_v2 == 1 or intersects_v3 == 1 or intersects_v4 == 1:
        return True
    return  False


def execute():
    with revit.Transaction("Привязка к помещениям"):
        rooms = []

        for roomCol in roomCols:
            for element in roomCol: # type: Room
                roomNumber =element.GetParamValue(BuiltInParameter.ROOM_NUMBER)
                roomName = element.GetParamValue(BuiltInParameter.ROOM_NAME)
                newRoom = flatroom(roomNumber, roomName, element.Id)
                borders = getBorders(element)
                newRoom.downBorder = borders[0]
                newRoom.topBorder = borders[1]

                if newRoom.roomId.IntegerValue == 2483864:
                    print newRoom.downBorder
                    print newRoom.topBorder
                    print 1

                if newRoom.roomId.IntegerValue == 6175751:
                    print newRoom.downBorder
                    print newRoom.topBorder
                    print 2
                segments = element.GetBoundarySegments(SpatialElementBoundaryOptions())
                for segmentList in segments:
                    for segment in segmentList:
                        coord_list = segment.GetCurve().Tessellate()
                        segmentLines = getTessallatedLine(coord_list)
                        newRoom.appendLine(segmentLines)
                rooms.append(newRoom)


        equipmentPoints = []
        for collection in collections:
            for element in collection:
                #element.Location.Curve
                newEquipment = elementPoint(element)
                equipmentPoints.append(newEquipment)

        print len(rooms)

        # for equipmentPoint in equipmentPoints:
        #     for room in rooms:
        #
        #         if isEquipmenInRoom(equipmentPoint, room):
        #             equipmentPoint.insert(room.roomNumber)
        #             index = equipmentPoints.index(equipmentPoint)
        #             break
        #             #equipmentPoints.pop(index)
        #         else:
        #             equipmentPoint.insert('None')

execute()
