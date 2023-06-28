#! /usr/bin/env python
# -*- coding: utf-8 -*-

__title__ = 'Полное\nобновление'
__doc__ = "Последовательно обновляет имя системы, функцию и саму спецификацию"

from pyrevit.loader.sessionmgr import execute_command

execute_command("04dotov-vk-овивк-сортировка-обновлениеименисистемы")

execute_command("04dotov-vk-овивк-сортировка-обновлениефункции")

execute_command("04dotov-vk-овивк-сортировка-обновлениеспецификации")
