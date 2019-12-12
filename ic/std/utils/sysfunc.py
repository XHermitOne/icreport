#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Общесистемные функции.
"""

import os
import sys
import locale
import platform

try:
    # Для Python 2
    import commands as get_procesess_module
except ImportError:
    # Для Python 3
    import subprocess as get_procesess_module

__version__ = (0, 1, 1, 1)


def getPlatform():
    """
    Определение платформы.
    """
    return platform.uname()[0].lower()


def isWindowsPlatform():
    return getPlatform() == 'windows'


def isLinuxPlatform():
    return getPlatform() == 'linux'


def getTerminalCodePage():
    """
    Кодировка командной оболочки по умолчанию.
    :return:
    """
    cmd_encoding = sys.stdout.encoding if isWindowsPlatform() else locale.getpreferredencoding()
    return cmd_encoding


def get_login():
    """
    Имя залогинненного пользователя.
    """
    username = os.environ.get('USERNAME', None)
    if username != 'root':
        return username
    else:
        return os.environ.get('SUDO_USER', None)


def getComputerName():
    """
    Имя компютера. Без перекодировки.
    :return: Получить имя компьютера в сети.
        Имя компьютера возвращается в utf-8 кодировке.
    """
    import socket
    comp_name = socket.gethostname()
    return comp_name


def getPythonMajorVersion():
    """
    Мажорная версия Python.
    """
    return sys.version_info.major


def getPythonMinorVersion():
    """
    Минорная версия Python.
    """
    return sys.version_info.minor


def isPython2():
    """
    Проверка на Python версии 2.
    :return: True - Python версии 2 / False - другая версия Python.
    """
    return sys.version_info.major == 2


def isPython3():
    """
    Проверка на Python версии 3.
    :return: True - Python версии 3 / False - другая версия Python.
    """
    return sys.version_info.major == 3


def getActiveProcessCount(find_process):
    """
    Количество активных выполняемых процессов.
    :param find_process: Строка поиска процесса.
    :param find_process:
    :return: Количество найденных процессов.
    """
    processes_txt = get_procesess_module.getoutput('ps -eo pid,cmd')
    processes = processes_txt.strip().split('\n')
    find_processes = [process for process in processes if find_process in process]
    return len(find_processes)


def isActiveProcess(find_process):
    """
    Проверка на существование активного выполняемого процесса.
    :param find_process: Строка поиска процесса.
    :return: True - есть такой процесс / False - процесс не найден.
    """
    return getActiveProcessCount(find_process) >= 1


def exit_force():
    """
    Принудительное закрытие программы
    """
    sys.exit(0)


def beep(count=1):
    """
    Воспроизвести системный звук.
    :param count: Количество повторений.
    """
    for i in range(count):
        print('\a')
