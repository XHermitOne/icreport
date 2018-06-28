#!/usr/bin/env python
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

__version__ = (0, 0, 8, 4)

DEFAULT_ENCODING = 'utf-8'

DEBUG_MODE = True
LOG_MODE = True

# Консольный режим работы
NO_GUI_MODE = False

# Имя папки прфиля программы
PROFILE_DIRNAME = '.icreport'

# Имя файла журнала
LOG_FILENAME = os.path.join(os.environ.get('HOME', os.path.dirname(__file__)+'/log'),
                            PROFILE_DIRNAME,
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
