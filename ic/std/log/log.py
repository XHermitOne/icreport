#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Функции журналирования.


Цветовая расскраска сообщений в коммандной оболочке
производиться только под Linux.
Для Windows систем цветовая раскраска отключена.

Шаблон для использования в современных
командных оболочках и языках
программирования таков: \x1b[...m.
Это ESCAPE-последовательность,
где \x1b обозначает символ ESC
(десятичный ASCII код 27), а вместо "..."
подставляются значения из таблицы,
приведенной ниже, причем они могут
комбинироваться, тогда нужно их
перечислить через точку с запятой.

атрибуты
0 	нормальный режим
1 	жирный
4 	подчеркнутый
5 	мигающий
7 	инвертированные цвета
8 	невидимый

цвет текста
30 	черный
31 	красный
32 	зеленый
33 	желтый
34 	синий
35 	пурпурный
36 	голубой
37 	белый

цвет фона
40 	черный
41 	красный
42 	зеленый
43 	желтый
44 	синий
45 	пурпурный
46 	голубой
47 	белый
"""

import sys
import logging
import os
import os.path
import tempfile
import stat
import traceback
import locale

__version__ = (0, 1, 2, 1)

# Кодировка коммандной оболочки по умолчанию
DEFAULT_ENCODING = sys.stdout.encoding if sys.platform.startswith('win') else locale.getpreferredencoding()

# Цвета в консоли
RED_COLOR_TEXT = '\x1b[31;1m'       # red
GREEN_COLOR_TEXT = '\x1b[32m'       # green
YELLOW_COLOR_TEXT = '\x1b[33;1m'    # yellow
BLUE_COLOR_TEXT = '\x1b[34m'        # blue
PURPLE_COLOR_TEXT = '\x1b[35m'      # purple
CYAN_COLOR_TEXT = '\x1b[36m'        # cyan
WHITE_COLOR_TEXT = '\x1b[37m'       # white
NORMAL_COLOR_TEXT = '\x1b[0m'       # normal

NOT_INIT_LOG_SYS_MSG = u'Не инициализирована система журналирования'

LOG_DATETIME_FMT = '%Y-%m-%d %H:%M:%S'


def print_color_txt(text, color=NORMAL_COLOR_TEXT):
    if isinstance(text, str):
        # text = text.encode(get_default_encoding())
        pass
    if sys.platform.startswith('win'):
        # Для Windows систем цветовая раскраска отключена
        txt = text
    else:
        # Добавление цветовой раскраски
        txt = color + text + NORMAL_COLOR_TEXT
    print(txt)        


# Модуль конфигурации
CONFIG = None


def get_default_encoding():
    """
    Определить актуальную кодировку для вывода текста.

    :return: Актуальная кодировка для вывода текста.
    """
    global CONFIG
    if CONFIG is not None and hasattr(CONFIG, 'DEFAULT_ENCODING'):
        # Приоритетной является явно указанная кодировка в конфигурационном файле
        return CONFIG.DEFAULT_ENCODING
    return DEFAULT_ENCODING


def get_debug_mode():
    """
    Определить актуальный режим отладки.
    По умолчанию считаем что режим выключен.

    :return: True - режим отладки включен / False - режим отладки выключен.
    """
    global CONFIG
    if CONFIG is not None and hasattr(CONFIG, 'DEBUG_MODE'):
        # Приоритетной является явно указанный параметр в конфигурационном файле
        return CONFIG.DEBUG_MODE
    # По умолчанию считаем что режим выключен
    return False


def get_log_mode():
    """
    Определить актуальный режим журналирования.
    По умолчанию считаем что режим выключен

    :return: True - режим журналирования включен / False - режим журналирования выключен.
    """
    global CONFIG
    if CONFIG is not None and hasattr(CONFIG, 'LOG_MODE'):
        # Приоритетной является явно указанный параметр в конфигурационном файле
        return CONFIG.LOG_MODE
    # По умолчанию считаем что режим выключен
    return False


def get_log_filename():
    """
    Имя файла журнала.

    :return:
    """
    global CONFIG
    return CONFIG.LOG_FILENAME if CONFIG and hasattr(CONFIG, 'LOG_FILENAME') else None


def init(config_module=None, log_filename=None):
    """
    Инициализация файла лога.

    :param config_module: Модуль конфигурации.
    :param log_filename: Полное имя файла журнала/лога.
    """
    global CONFIG
    CONFIG = config_module
    
    if not get_log_mode():
        return
    
    if log_filename is None:
        log_filename = CONFIG.LOG_FILENAME if hasattr(CONFIG, 'LOG_FILENAME') else tempfile.mktemp()
        
    # Создать папку логов если она отсутствует
    log_dirname = os.path.normpath(os.path.dirname(log_filename))
    if not os.path.exists(log_dirname):
        os.makedirs(log_dirname)
        
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(levelname)s %(message)s',
                        datefmt=LOG_DATETIME_FMT,
                        filename=log_filename,
                        filemode='a')
    # ВНИМАНИЕ! сразу выставить права для записи/чтения для всех
    # иначе в ряде случаев может не производится запись в файл и поэтому падать
    if os.path.exists(log_filename):
        os.chmod(log_filename,
                 stat.S_IWUSR | stat.S_IRUSR |
                 stat.S_IWGRP | stat.S_IRGRP |
                 stat.S_IWOTH | stat.S_IROTH)

    if get_debug_mode():
        print_color_txt('INFO. Инициализация файла журнала <%s>' % log_filename, GREEN_COLOR_TEXT)


def debug(message=u'', bForcePrint=False, bForceLog=False):
    """
    Вывести ОТЛАДОЧНУЮ информацию.

    :param message: Текстовое сообщение.
    :param bForcePrint: Принудительно вывести на экран.
    :param bForceLog: Принудительно записать в журнале.
    """
    global CONFIG
    
    if CONFIG:
        if get_debug_mode() or bForcePrint:
            print_color_txt('DEBUG. ' + message, BLUE_COLOR_TEXT)
        if get_log_mode() or bForceLog:
            logging.debug(message)
    else:
        print_color_txt(NOT_INIT_LOG_SYS_MSG, PURPLE_COLOR_TEXT)
        print_color_txt('DEBUG. ' + message, BLUE_COLOR_TEXT)


def info(message=u'', bForcePrint=False, bForceLog=False):
    """
    Вывести ТЕКСТОВУЮ информацию.

    :param message: Текстовое сообщение.
    :param bForcePrint: Принудительно вывести на экран.
    :param bForceLog: Принудительно записать в журнале.
    """
    global CONFIG
    
    if CONFIG:
        if get_debug_mode() or bForcePrint:
            print_color_txt('INFO. '+message, GREEN_COLOR_TEXT)
        if get_log_mode() or bForceLog:
            logging.info(message)    
    else:
        print_color_txt(NOT_INIT_LOG_SYS_MSG, PURPLE_COLOR_TEXT)
        print_color_txt('INFO. ' + message, GREEN_COLOR_TEXT)


def error(message=u'', bForcePrint=False, bForceLog=False):
    """
    Вывести ОБЩУЮ информацию.

    :param message: Текстовое сообщение.
    :param bForcePrint: Принудительно вывести на экран.
    :param bForceLog: Принудительно записать в журнале.
    """
    global CONFIG
    
    if CONFIG:
        if get_debug_mode() or bForcePrint:
            print_color_txt('ERROR. '+message, RED_COLOR_TEXT)
        if get_log_mode() or bForceLog:
            logging.error(message)    
    else:
        print_color_txt(NOT_INIT_LOG_SYS_MSG, PURPLE_COLOR_TEXT)
        print_color_txt('ERROR. ' + message, RED_COLOR_TEXT)


def warning(message=u'', bForcePrint=False, bForceLog=False):
    """
    Вывести информацию ОБ ПРЕДУПРЕЖДЕНИИ.

    :param message: Текстовое сообщение.
    :param bForcePrint: Принудительно вывести на экран.
    :param bForceLog: Принудительно записать в журнале.
    """
    global CONFIG
    
    if CONFIG:
        if get_debug_mode() or bForcePrint:
            print_color_txt('WARNING. '+message, YELLOW_COLOR_TEXT)
        if get_log_mode() or bForceLog:
            logging.warning(message)    
    else:
        print_color_txt(NOT_INIT_LOG_SYS_MSG, PURPLE_COLOR_TEXT)
        print_color_txt('WARNING. ' + message, YELLOW_COLOR_TEXT)


def fatal(message=u'', bForcePrint=False, bForceLog=False):
    """
    Вывести информацию ОБ ОШИБКЕ.

    :param message: Текстовое сообщение.
    :param bForcePrint: Принудительно вывести на экран.
    :param bForceLog: Принудительно записать в журнале.
    """
    global CONFIG

    trace_txt = traceback.format_exc()

    try:
        msg = message+u'\n'+trace_txt
    except UnicodeDecodeError:
        if not isinstance(message, str):
            message = str(message)
        if not isinstance(trace_txt, str):
            trace_txt = str(trace_txt)
        msg = message+u'\n'+trace_txt

    if CONFIG:
        if get_debug_mode() or bForcePrint:
            print_color_txt('FATAL. '+msg, RED_COLOR_TEXT)
        if get_log_mode() or bForceLog:
            logging.fatal(msg)
    else:
        print_color_txt(NOT_INIT_LOG_SYS_MSG, PURPLE_COLOR_TEXT)
        print_color_txt('FATAL. ' + msg, RED_COLOR_TEXT)


def service(message=u'', bForcePrint=False, bForceLog=False):
    """
    Вывести СЕРВИСНУЮ информацию.

    :param message: Текстовое сообщение.
    :param bForcePrint: Принудительно вывести на экран.
    :param bForceLog: Принудительно записать в журнале.
    """
    global CONFIG

    if CONFIG:
        if get_debug_mode() or bForcePrint:
            print_color_txt('SERVICE. ' + message, CYAN_COLOR_TEXT)
        if get_log_mode() or bForceLog:
            logging.debug('SERVICE. ' + message)
    else:
        print_color_txt(NOT_INIT_LOG_SYS_MSG, PURPLE_COLOR_TEXT)
        print_color_txt('SERVICE. ' + message, CYAN_COLOR_TEXT)
