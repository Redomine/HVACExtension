#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'По смене сечения и расхода'
__doc__ = "Вызывает маркировку элементов по смене сечения и расхода"

import os.path as op
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import markSystems


markSystems.markBySizeAndFlow()