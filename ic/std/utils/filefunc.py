#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль функций пользователя для работы с файлами.
"""

# Подключение библиотек
import sys
import os
import os.path
import platform
import pwd
import fnmatch

from ic.std.log import log

__version__ = (0, 1, 1, 2)

# Имя папки прфиля программы
DEFAULT_PROFILE_DIRNAME = '.icreport'


def _pathFilter(path, path_filter):
    """
    Фильтрация путей.
    @return: Возвращает True если папок с указанными имена в фильтре нет в пути и
        False если наоборот.
    """
    path = os.path.normpath(path).replace('\\', '/')
    path_lst = path.split('/')
    filter_result = True
    for cur_filter in path_filter:
        if cur_filter in path_lst:
            filter_result = False
            break
    return filter_result


# Папки, которые не надо обрабатывать по умолчанию
DEFAULT_DIR_FILTER = ('.svn', '.SVN', '.Svn', '.idea', '.Idea', '.IDEA')


def getSubDirsFilter(path, dir_filter=DEFAULT_DIR_FILTER):
    """
    Функция возвращает список поддиректорий с отфильтрованными папками.
    @param path: Дeрикторий.
    @param dir_filter: Список недопустимых имен папок.
    @return: В случае ошибки возвращает None.
    """
    try:
        dir_list = [os.path.normpath(os.path.join(path, path)) for path in os.listdir(path)]
        dir_list = [path for path in dir_list if os.path.isdir(path)]
        dir_list = [dir for dir in dir_list if _pathFilter(dir, dir_filter)]
        return dir_list
    except:
        log.fatal(u'Ошибка чтения списка поддиректорий <%s>' % path)
        return None


def getFilesByExt(path, ext):
    """
    Функция возвращает список всех файлов в директории с указанным расширением.
    @param path: Путь.
    @param ext: Расширение, например '.pro'.
    @return: В случае ошибки возвращает None.
    """
    file_list = None
    try:
        path = os.path.abspath(os.path.normpath(path))

        # Приведение расширения к надлежащему виду
        if ext[0] != '.':
            ext = '.' + ext
        ext = ext.lower()

        file_list = [os.path.normpath(os.path.join(path, path)) for path in os.listdir(path)]
        file_list = [path for path in file_list if os.path.isfile(path) and os.path.splitext(path)[1].lower() == ext]
        return file_list
    except:
        log.fatal(u'Ошибка чтения списка файлов директории. ext: <%s>, path: <%s>, list: <%s>' % (ext, path, file_list))
        return None


def getHomePath():
    """
    Путь к домашней директории.
    @return: Строку-путь до папки пользователя.
    """
    os_platform = platform.uname()[0].lower()
    if os_platform == 'windows':
        home_path = os.environ['HOMEDRIVE'] + os.environ['HOMEPATH']
        home_path = home_path.replace('\\', '/')
    elif os_platform == 'linux':
        home_path = os.environ['HOME']
    else:
        log.warning(u'Не поддерживаемая <%s>' % os_platform)
        return None
    return os.path.normpath(home_path)


def getProfilePath(bAutoCreatePath=True):
    """
    Папка профиля программы.
    @param bAutoCreatePath: Создать автоматически путь если его нет?
    @return: Путь до профиля программы.
    """
    home_path = getHomePath()
    if home_path:
        profile_path = os.path.join(home_path, DEFAULT_PROFILE_DIRNAME)
        if not os.path.exists(profile_path) and bAutoCreatePath:
            # Автоматическое создание пути
            try:
                os.makedirs(profile_path)
            except OSError:
                log.fatal(u'Ошибка создания пути профиля <%s>' % profile_path)
        return profile_path
    return os.path.join('~', DEFAULT_PROFILE_DIRNAME)


def get_home_path(username=None):
    """
    Определить домашнюю папку пользователя.
    """
    if sys.platform[:3].lower() == 'win':
        home = os.path.join(os.environ['HOMEDRIVE'], os.environ['HOMEPATH'])
    else:
        if username is None:
            home = os.environ['HOME']
        else:
            user_struct = pwd.getpwnam(username)
            home = user_struct.pw_dir
    return home


HOME_PATH_SIGN = '~'


def normal_path(path, username=None):
    """
    Нормировать путь.
    @param path: Путь.
    @param username: Имя пользователя.
    """
    home_dir = get_home_path(username)
    return os.path.abspath(os.path.normpath(path.replace(HOME_PATH_SIGN, home_dir)))


def file_modify_dt(filename):
    """
    Дата-время изменения файла.
    @param filename: Полное имя файла.
    @return: Дата-время изменения файла или None в случае ошибки.
    """
    if not os.path.exists(filename):
        log.warning(u'Файл <%s> не найден' % filename)
        return None

    try:
        if platform.system() == 'Windows':
            return os.path.getmtime(filename)
        else:
            stat = os.stat(filename)
            return stat.st_mtime
    except:
        log.fatal(u'Ошибка определения даты-времени изменения файла <%s>' % filename)
    return None


def remove_file(filename):
    """
    Удалить файл.
    @param filename: Имя файла.
    @return: True/False.
    """
    if not os.path.exists(filename):
        log.warning(u'Удаление. Файл <%s> не найден' % filename)
        return False

    try:
        os.remove(filename)
        log.info(u'Файл <%s> удален' % filename)
        return True
    except:
        log.fatal(u'Ошибка удаления файла <%s>' % filename)
    return False
