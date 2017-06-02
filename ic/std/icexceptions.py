#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Исключения.
"""
import exceptions

class icMergeCellError(exceptions.Exception):
    """
    Ошибка обращения к запрещенной области объединенной ячейки.
    """
    def __init__(self, args=None, user=None):
        self.args = args

class icCellAddressInvalidError(exceptions.Exception):
    """
    Ошибка некорректного адреса ячейки.
    """
    def __init__(self, args=None, user=None):
        self.args = args
