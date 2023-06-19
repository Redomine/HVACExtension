#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Нумерация и примечания'
__doc__ = "Проставляет в примечаниях к воздуховодам их площади а в позиции номера элементов"

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
from Redomine import *
from numerateSpec import *
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output

doc = __revit__.ActiveUIDocument.Document  # type: Document
view = doc.ActiveView



if isItFamily():
    print 'Надстройка не предназначена для работы с семействами'
    sys.exit()

parametersAdded = paraSpec.check_parameters()



if not parametersAdded:
    numerate(doNumbers = True, doAreas = True)
