#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Исключения.
"""

__version__ = (0, 1, 1, 1)


class icMergeCellError(Exception):
    """
    Ошибка обращения к запрещенной области объединенной ячейки.
    """
    def __init__(self, args=None, user=None):
        self.args = args


class icCellAddressInvalidError(Exception):
    """
    Ошибка некорректного адреса ячейки.
    """
    def __init__(self, args=None, user=None):
        self.args = args
