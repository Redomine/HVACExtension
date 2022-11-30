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
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
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

roomCol = make_col(BuiltInCategory.OST_Rooms)


class line:
    def __init__(self, x1, y1, x2, y2):
        self.x1 = x1
        self.y1 = y1
        self.x2 = x2
        self.y2 = y2

class elementPoint:
    def getElementCenter(self, line):
        self.x = (line.x1 + line.x2)/2
        self.y = (line.y1 + line.y2)/2

    def pointInRoom(self, rooms):
        vector_1 = 0
        vector_2 = 0
        vector_3 = 0
        vector_4 = 0
        for room in rooms:





    def __init__(self, element):
        self.element = element
        self.room = None
        self.x = 0
        self.y = 0
        #try:
        if not 'LocationPoint' in str(element.Location):
            self.elementLines = getTessallatedLine(element.Location.Curve.Tessellate())
            self.elementCenter = self.getElementCenter(self.elementLines[0])
        else:
            self.elementCenter = element.Location.Point





class flatroom:
    def appendLine(self, lines):
        for line in lines:
            self.roomLines.append(line)

    def __init__(self):
        self.roomLines = []

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

def execute():
    rooms = []
    for element in roomCol: # type: Room
        newRoom = flatroom()
        segments = element.GetBoundarySegments(SpatialElementBoundaryOptions())
        for segmentList in segments:
            for segment in segmentList:
                segmentLines = getTessallatedLine(segment.GetCurve().Tessellate())
                newRoom.appendLine(segmentLines)
        rooms.append(newRoom)

    equipment = []
    for collection in collections:
        for element in collection:
            #element.Location.Curve
            newEquipment = elementPoint(element)
            equipment.append(newEquipment)


    for element in equipment:
        element.pointInRoom(rooms)



    for room in rooms:
        for room.roomLine in room.roomLines:
            pass
            #print str(room.roomLine.x1) + '_' + str(room.roomLine.y1) + '_' + str(room.roomLine.x2) + '_' + str(room.roomLine.y2)





execute()
