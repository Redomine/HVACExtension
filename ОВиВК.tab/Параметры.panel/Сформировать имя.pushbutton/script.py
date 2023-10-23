#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Сформировать имя'
__doc__ = "Генерирует имена и марки арматуры воздуховодов из их маски, и ADSK_ размеров"

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
from System import Guid
from itertools import groupby
from pyrevit import revit
from pyrevit.script import output
import combineNames


combineNames.combine_names()