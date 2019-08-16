#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Конфигурационный файл.

Параметры:

@type DEBUG_MODE: C{bool}
@var DEBUG_MODE: Режим отладки (вкл./выкл.)
@type LOG_MODE: C{bool}
@var LOG_MODE: Режим журналирования (вкл./выкл.)
"""

import os.path
import datetime

__version__ = (0, 1, 2, 1)

DEFAULT_ENCODING = 'utf-8'

DEBUG_MODE = True
LOG_MODE = True

# Консольный режим работы
NO_GUI_MODE = False

# Имя папки профиля программы
PROFILE_DIRNAME = '.icreport'
# Путь до папки профиля
PROFILE_PATH = os.path.join(os.environ.get('HOME', os.path.dirname(__file__)),
                            PROFILE_DIRNAME)

# Имя файла журнала
LOG_FILENAME = os.path.join(PROFILE_PATH,
                            'icreport_%s.log' % datetime.date.today().isoformat())


def get_glob_var(name):
    """
    Прочитать значение глобальной переменной.
    @type name: C{string}
    @param name: Имя переменной.
    """
    return globals()[name]


def set_glob_var(name, value):
    """
    Установить значение глобальной переменной.
    @type name: C{string}
    @param name: Имя переменной.
    @param value: Значение переменной.
    """
    globals()[name] = value
    return value
