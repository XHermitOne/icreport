#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os.path

# Режим отладки
DEBUG_MODE = True

# Режим журналирования
LOG_MODE = True

# Кодировка консоли по умолчанию
DEFAULT_ENCODING = 'utf-8'

LOG_FILENAME = '%s/log/virtual_excel.log' % os.path.dirname(__file__)


# Определять адресацию внутри объединенной ячейки как ошибку
DETECT_MERGE_CELL_ERROR = False


def get_cfg_var(sName):
    """
    Прочитать значение переменной конфига.
    @type sName: C{string}
    @param sName: Имя переменной.
    """
    return globals()[sName]


def set_cfg_var(sName, vValue):
    """
    Установить значение переменной конфига.
    @type sName: C{string}
    @param sName: Имя переменной.
    @param vValue: Значение переменной.
    """
    globals()[sName] = vValue
