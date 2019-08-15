#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os.path

# Режим отладки
DEBUG_MODE = True

# Режим журналирования
LOG_MODE = True

# Кодировка консоли по умолчанию
DEFAULT_ENCODING = 'utf-8'

LOG_FILENAME = os.path.join(os.path.dirname(__file__),
                            'log', 'virtual_excel.log')


# Определять адресацию внутри объединенной ячейки как ошибку
DETECT_MERGE_CELL_ERROR = False


def get_cfg_var(name):
    """
    Прочитать значение переменной конфига.
    @type name: C{string}
    @param name: Имя переменной.
    """
    return globals()[name]


def set_cfg_var(name, value):
    """
    Установить значение переменной конфига.
    @type name: C{string}
    @param name: Имя переменной.
    @param value: Значение переменной.
    """
    globals()[name] = value
