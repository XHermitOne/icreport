# !/usr/bin/env python
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

__version__ = (0, 0, 7, 3)

DEFAULT_ENCODING = 'utf-8'

DEBUG_MODE = True
LOG_MODE = True

# Имя папки прфиля программы
PROFILE_DIRNAME = '.icreport'

# Имя файла журнала
LOG_FILENAME = os.path.join(os.environ.get('HOME', os.path.dirname(__file__)+'/log'),
                            PROFILE_DIRNAME,
                            'icreport_%s.log' % datetime.date.today().isoformat())
