#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Методика расчета'
__doc__ = "Открывает гугл таблицу с исходными формулами для расчетов"


import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

import webbrowser

webbrowser.open('https://docs.google.com/spreadsheets/d/1bi752PrerzCJM7mjaGJFF7ANEKhAgQLd_fdVxecuQS0/edit?usp=sharing', new=0, autoraise=True)




