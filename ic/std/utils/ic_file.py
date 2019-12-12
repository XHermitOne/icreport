#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль функций пользователя для работы с файлами.
"""

# --- Подключение пакетов ---
import wx
import os
import os.path
import tempfile
import shutil   # Для реализации высокоуровневых функций работы с файлами
import sys
import time
import glob     # Для поиска файлов по маске/шаблону
import platform

from ic.log import log
from ic.dlg import ic_dlg

import ic.config

__version__ = (1, 1, 1, 2)

_ = wx.GetTranslation


# --- Функции пользователя ---
def getCreateFileTime(filename):
    """
    Время создания файла. Если файла не существует то 0.
    """
    if os.path.exists(filename):
        return os.path.getmtime(filename)
    return 0


def make_dirs(path):
    """
    Корректное создание каталогов по цепочке.
    """
    try:
        if not os.path.exists(path):
            return os.makedirs(path)
    except:
        log.fatal(u'Ошибка создания каталога <%s>' % path)


def changeFileExt(filename, new_ext):
    """
    Поменять у файла расширение.
    :param filename: Полное имя файла.
    :param new_ext: Новое расширение файла (Например: '.bak').
    :return: Возвращает новое полное имя файла.
    """
    try:
        new_name = os.path.splitext(filename)[0] + new_ext
        if os.path.isfile(new_name):
            os.remove(new_name)     # если файл существует, то удалить
        if os.path.exists(filename):
            os.rename(filename, new_name)
            return new_name
    except:
        log.fatal(u'Ошибка изменения расширения файла <%s> -> <%s>' % (filename, new_ext))
    return None


def copyFile(filename, new_filename, bRewrite=True):
    """
    Создает копию файла с новым именем.
    :param filename: Полное имя файла.
    :param new_filename: Новое имя файла.
    :param bRewrite: True-если новый файл уже существует,
        то переписать его молча. False-если новый файл уже существует, 
        то выдать сообщение о подтверждении перезаписи файла.
    :return: Возвращает результат выполнения операции True/False.
    """
    try:
        # --- Проверка существования файла-источника ---
        if not os.path.exists(filename):
            msg = u'Копирование <%s> -> <%s>. Файл <%s> не существует.' % (filename, new_filename, filename)
            log.warning(msg)
            ic_dlg.icWarningBox(u'ОШИБКА', msg)
            return False

        make_dirs(os.path.dirname(new_filename))

        # --- Проверка перезаписи уже существуещего файла ---
        # Выводить сообщение что файл уже существует?
        if not bRewrite:
            # Файл уже существует?
            if os.path.exists(new_filename):
                if ic_dlg.icAskDlg(u'КОПИРВАНИЕ',
                                   u'Файл <%s> уже существует. Переписать?' % new_filename) == wx.NO:
                    return False
        else:
            if os.path.exists(new_filename):
                os.remove(new_filename)

        # --- Реализация копирования файла ---
        if os.path.exists(filename) and os.path.exists(new_filename) and os.path.samefile(filename, new_filename):
            log.warning(u'Попытка скопировать файл <%s> самого в себя' % filename)
        else:
            shutil.copyfile(filename, new_filename)
        return True
    except:
        log.fatal(u'Ошибка копирования файла <%s> -> <%s>' % (filename, new_filename))
        return False


def createBAKFile(filename, bak_ext='.bak'):
    """
    Создает копию файла с новым расширением BAK.
    :param filename: Полное имя файла.
    :param bak_ext: Расширение BAK файла.
    :return: Возвращает результат выполнения операции True/False.
    """
    try:
        if not os.path.exists(filename):
            log.warning(u'Не найден файл <%s> для создания его резервной копии' % filename)
            return False

        bak_name = os.path.splitext(filename)[0] + bak_ext
        return copyFile(filename, bak_name)
    except:
        log.fatal(u'Ошибка создания BAK файла <%s>' % filename)
        return False


def getSubDirs(path):
    """
    Функция возвращает список поддиректорий.
    :param path: Дирикторий.
    :return: В случае ошибки возвращает None.
    """
    try:
        if not os.path.exists(path):
            log.warning(u'Путь <%s> не найден для определения списка поддриекторий' % path)
            return list()
        dir_list = [os.path.join(path, path) for path in os.listdir(path)]
        dir_list = [path for path in dir_list if os.path.isdir(path)]
        return dir_list
    except:
        log.fatal(u'Ошибка чтения списка поддиректорий <%s>' % path)
        return None


def getSubDirsFilter(path, dir_filter=('.svn', '.SVN', '.Svn')):
    """
    Функция возвращает список поддиректорий с отфильтрованными папками.
    :param path: Дирикторий.
    :param dir_filter: Список недопустимых имен папок.
    :return: В случае ошибки возвращает None.
    """
    try:
        if not os.path.exists(path):
            log.warning(u'Не найден путь <%s> для определения списка поддиректорий' % path)
            return list()

        dir_list = [os.path.join(path, path) for path in os.listdir(path)]
        dir_list = [path for path in dir_list if os.path.isdir(path)]
        dir_list = [d for d in dir_list if _pathFilter(d, dir_filter)]
        return dir_list
    except:
        log.fatal(u'Ошибка чтения списка поддиректорий <%s>' % path)
        return None


def getSubDirsFilterSVN(path):
    """
    Функция возвращает список поддиректорий с отфильтрованными папками Subversion.
    :param path: Дирикторий.
    :param path_filter: Список недопустимых имен папок.
    :return: В случае ошибки возвращает None.
    """
    return getSubDirsFilter(path)


def getFilenames(path):
    """
    Функция возвращает список файлов в директории.
    :param path: Дирикторий.
    :return: В случае ошибки возвращает None.
    """
    try:
        if not os.path.exists(path):
            log.warning(u'Не найден путь <%s> для определения списка файлов директории' % path)
            return list()

        file_list = [os.path.join(path, x.lower()) for x in os.listdir(path)]
        file_list = [x for x in file_list if os.path.isfile(x)]
        return file_list
    except:
        log.fatal(u'Ошибка чтения списка файлов <%s>' % path)
    return None


def getFilenamesByExt(path, ext):
    """
    Функция возвращает список всех файлов в директории с указанным расширением.
    :param path: Путь.
    :param ext: Расширение, например '.pro'.
    :return: В случае ошибки возвращает None.
    """
    try:
        path = getCurDirPrj(path)
        if not os.path.exists(path):
            log.warning(u'Путь <%s> не найден для определения списка файлов директории по расширению' % path)
            return list()

        if ext[0] != '.':
            ext = '.' + ext
        ext = ext.lower()
            
        file_list = None
        file_list = [os.path.join(path, file_name) for file_name in os.listdir(path)]
        file_list = [file_name for file_name in file_list if os.path.isfile(file_name) and
                     (os.path.splitext(file_name)[1].lower() == ext)]
        return file_list
    except:
        log.fatal(u'Ошибка чтения списка файлов <ext=%s, path=%s, list=%s>' % (ext, path, file_list))
    return None


def clearFileExt(path, ext):
    """
    Функция УДАЛЯЕТ РЕКУРСИВНО В ПОДДИРЕКТОРИЯХ все файлы в директории с
    заданным расширением.
    :param path: Путь.
    :param ext: Расширение.
    :return: Возвращает результат выполнения операции True/False.
    """
    try:
        ok = True
        dir_list = os.listdir(path)
        for cur_item in dir_list:
            cur_file = path + cur_item
            if os.path.isfile(cur_file) and os.path.splitext(cur_file)[1] == ext:
                os.remove(cur_file)
            elif os.path.isdir(cur_file):
                ok = ok and clearFileExt(cur_file, ext)
        return ok
    except:
        return False        


def getFileExt(filename):
    """
    Получить расширение файла с точкой.
    """
    return os.path.splitext(filename)[1]


def get_current_dir():
    """
    Текущая папка.
    Относительнай путь считается от папки defis.
    :return:
    """
    cur_dir = os.path.dirname(os.path.dirname(ic.config.__file__))
    log.debug(u'Текущая папка определена как <%s>' % cur_dir)
    return cur_dir


def getRelativePath(path):
    """
    Относительный путь.
    Относительнай путь считается от папки defis.
    :param path: Путь.
    """
    path = os.path.normpath(path)
    cur_dir = get_current_dir()
    return path.replace(cur_dir, '.').strip()


def getAbsolutePath(path):
    """
    Абсолютный путь.
    :param path: Путь.
    """
    try:
        cur_dir = get_current_dir()
        if path.startswith('..'):
            path = os.path.join(os.path.dirname(cur_dir), path[2 + len(os.path.sep):])
        elif path.startswith('.'):
            path = os.path.join(cur_dir, path[1 + len(os.path.sep):])
        path = os.path.normpath(path)
        return path
    except:
        log.fatal(u'Ошибка определения абсолютного пути <%s>' % path)
        return path


def RelativePath(path, cur_dir=None):
    """
    Относительный путь. Путь приводится к виду Unix.
    :param path: Путь.
    :param cur_dir: Текущий путь.
    """
    if cur_dir is None:
        import ic.engine.ic_user
        cur_dir = os.path.dirname(ic.engine.ic_user.icGet('PRJ_DIR')).replace('\\', '/').lower()
    if cur_dir:
        path = path.replace('\\', '/').lower().strip()
        return path.replace(cur_dir, '.')
    return path


def getCurDirPrj(path=None):
    """
    Текущий путь. Определяется относительно PRJ_DIR.
    """
    # Нормализация текущего пути
    if path is None:
        try:
            import ic.engine.ic_user
            prj_dir = ic.engine.ic_user.icGet('PRJ_DIR')
            if prj_dir:
                path = os.path.dirname(prj_dir)
            else:
                path = getProfilePath()
        except:
            log.fatal(u'Ошибка определения пути <%s>' % path)
            path = os.getcwd()
    path = path.replace('\\', '/')
    if path[-1] != '/':
        path += '/'
    return path


def AbsolutePath(path, cur_dir=None):
    """ 
    Абсолютный путь. Путь приводится к виду Unix. 
    :param path: Путь.
    :param cur_dir: Текущий путь.
    """
    try:
        if not path:
            log.error(u'Не определен путь для приведения к абсолютному виду')
            return None

        # Нормализация текущего пути
        cur_dir = getCurDirPrj(cur_dir)

        # Коррекция самого пути
        path = os.path.abspath(path.replace('./', cur_dir).strip())
    except:
        log.fatal(u'Ошибка определения абсолютног пути <%s>. Текущая директория <%s>' % (path, cur_dir))
    return path


def PathFile(path, filename):
    """
    Корректное представление общего имени файла.
    :param path: Путь.
    :param filename: Имя файла.
    """
    if not path:
        log.warning(u'Не определен путь для корректировки')
        return filename
    if not filename:
        log.warning(u'Не определено имя файла для корректировки')
        return filename

    path = os.path.normpath(path)
    filename = os.path.normpath(filename)
    relative_path = getRelativePath(path)
    # Этот путь уже присутствует в имени файла
    if filename.find(path) != -1 or filename.find(relative_path) != -1:
        return filename
    return os.path.join(relative_path, filename)


def normPathWin(path):
    """
    Приведение пути к виду Windows.
    """
    if not path:
        return ''
        
    if path.find(' ') > -1 and path[0] != '\'' and path[-1] != '\'':
        return '\'' + os.path.normpath(path).strip() + '\''
    else:
        return os.path.normpath(path).strip()


def normPathUnix(path):
    """
    Приведение пути к виду UNIX.
    """
    return os.path.normpath(path).replace('\\', '/').strip()


def samePathWin(path1, path2):
    """
    Проверка,  path1==path2.
    """
    return bool(normPathWin(path1).lower() == normPathWin(path2).lower())


def _pathFilter(path, path_filter):
    """
    Фильтрация путей.
    :return: Возвращает True если папок с указанными имена в фильтре нет в пути и
        False если наоборот.
    """
    path = os.path.normpath(path).replace('\\', '/')
    path_lst = path.split(os.path.sep)
    filter_result = True
    for cur_filter in path_filter:
        if cur_filter in path_lst:
            filter_result = False
            break
    return filter_result


def _addCopyDirWalk(args, cur_dir, cur_names):
    """
    Функция рекурсивного обхода при добавлении папок и файлов в существующую.
    :param cur_dir: Текущая обрабатываемая папка.
    :param cur_names: Имена файлов и папок в текущей обрабатываемой папке.
    """
    from_dir = args[0]
    to_dir = args[1]
    not_copy_filter = args[2]
    
    if _pathFilter(cur_dir, not_copy_filter):
        paths = [os.path.join(cur_dir, name) for name in cur_names if name not in not_copy_filter]
        for path in paths:
            to_path = path.replace(from_dir, to_dir)
            if not os.path.exists(to_path):
                # Копировать если результирующего файла/папки не существует
                if os.path.isfile(path):
                    # Скопировать файл
                    copyFile(path, to_path)
                elif os.path.isdir(path):
                    # Создать директорию
                    try:
                        os.makedirs(to_path)
                    except:
                        log.fatal(u'Ошибка создания папки <%s>' % to_path)
                        raise


def addCopyDir(src_dirname, dst_dirname, not_copy_filter=('.svn', '.SVN', '.Svn')):
    """
    Дополнить папку dst_dirname файлами и папками из src_dirname
    :param src_dirname: Папка/директория,  которая копируется.
    :param dst_dirname: Папка/директория, в которую копируется src_dirname.
    :param not_copy_filter: Не копировать файлы/папки.
    """
    try:
        os.walk(src_dirname, _addCopyDirWalk, (src_dirname, dst_dirname, not_copy_filter))
        return True
    except:
        log.fatal(u'Ошибка дополнения папки из <%s> в <%s>' % (src_dirname, dst_dirname))
        return False


def copyDir(src_dirname, dst_dirname, bReWrite=False, bAddDir=True):
    """
    Функция папку src_dirname в папку dst_dirname со всеми внутренними поддиректориями
    и файлами.
    :param src_dirname: Папка/директория,  которая копируется.
    :param dst_dirname: Папка/директория, в которую копируется src_dirname.
    :param bReWrite: Указание перезаписи директории,
        если она уже существует.
    :param bAddDir: Указание производить дополнение папки,
        в случае ко когда копируемые файлы/папки существуют.
    :return: Функция возвращает результат выполнения операции True/False.
    """
    try:
        to_dir = os.path.join(dst_dirname, os.path.basename(src_dirname))
        if os.path.exists(to_dir) and bReWrite:
            log.info(u'Удаление папки <%s>' % to_dir)
            shutil.rmtree(to_dir, 1)
        if os.path.exists(to_dir) and bAddDir:
            return addCopyDir(src_dirname, to_dir)
        else:
            log.info(u'Копировние папки <%s> в <%s>' % (src_dirname, to_dir))
            shutil.copytree(src_dirname, to_dir)
        return True
    except:
        log.fatal(u'Ошибка копирования папки из <%s> в <%s>' % (src_dirname, dst_dirname))
        return False


def cloneDir(src_dirname, new_dirname, bReWrite=False):
    """
    Функция переносит все содержимое папки src_dirname в папку с новым именем new_dirname.
    :param src_dirname: Папка/директория,  которая копируется.
    :param new_dirname: Новое имя папки/директории.
    :param bReWrite: Указание перезаписи директории, если она
        уже существует.
    :return: Функция возвращает результат выполнения операции True/False.
    """
    try:
        if os.path.exists(new_dirname) and bReWrite:
            shutil.rmtree(new_dirname, 1)
        os.makedirs(new_dirname)
        for sub_dir in getSubDirs(src_dirname):
            shutil.copytree(sub_dir, new_dirname)
        for file_name in getFilenames(src_dirname):
            copyFile(file_name, os.path.join(new_dirname, os.path.basename(file_name)))
        return True
    except:
        log.fatal(u'Ошибка клонирования папки из <%s> в <%s>' % (src_dirname, new_dirname))
    return False


def isSubDir(dirname1, dirname2):
    """
    Функция проверяет, является ли директория dirname1 поддиректорией dirname2.
    :return: Возвращает True/False.
    """
    dir1 = os.path.abspath(dirname1)
    dir2 = os.path.abspath(dirname2)
    if dir1 == dir2:
        return True
    else:
        sub_dirs = [path for path in [os.path.join(dir2, name) for name in os.listdir(dir2)] if os.path.isdir(path)]
        for cur_sub_dir in sub_dirs:
            find = isSubDir(dirname1, cur_sub_dir)
            if find:
                return find
    return False


def genDefaultBakFileName():
    """
    Генерация имени бак файла по текущему времени.
    """
    return time.strftime('_%d_%m_%Y_%H_%M_%S.bak', time.localtime(time.time()))


def getFilesByMask(filename_mask):
    """
    Список файлов по маске.
    :param filename_mask: Маска файлов. Например C:\Temp\*.dbf.
    :return: Возвращает список строк-полных путей к файлам.
        В случае ошибки None.
    """
    try:
        if isinstance(filename_mask, str):
            dir_path = os.path.dirname(filename_mask)
            if os.path.exists(dir_path):
                filenames = glob.glob(pathname=filename_mask, recursive=False)
                return [os.path.abspath(file_name) for file_name in filenames]
            else:
                log.warning(u'Не найден путь <%s> для определения списка файлов по маске <%s>' % (dir_path, filename_mask))
        elif isinstance(filename_mask, tuple) or isinstance(filename_mask, list):
            filenames = list()
            for file_mask in filename_mask:
                filenames = glob.glob(pathname=filename_mask, recursive=False)
                filenames += [os.path.abspath(file_name) for file_name in filenames]
            return filenames
        else:
            log.warning(u'Не поддерживаемый тип аргумента в функции getFilesByMask')
    except:
        log.fatal(u'Ошибка определения списка файлов по маске <%s>' % str(filename_mask))
    return []


def copyToDir(filename, dst_dirname, bReWrite=True):
    """
    Копировать файл в папку.
    :param filename: Имя файла.
    :param dst_dirname: Папка в которую необходимо скопировать.
    :param bReWrite: True-если новый файл уже существует, 
        то переписать его молча. False-если новый файл уже существует, 
        то выдать сообщение о подтверждении перезаписи файла.
    :return: Возвращает результат выполнения операции True/False.
    """
    return copyFile(filename, os.path.join(dst_dirname,
                                           os.path.basename(filename)), bReWrite)


def delAllFilesFilter(del_dirname, *del_filter):
    """
    Удаление всех файлов из папки с фильтрацией по маске файла. Удаление
    рекурсивное по поддиректориям.
    :param del_dirname: Папка-источник.
    :param del_filter: Список масок файлов которые нужно удалить.
        Например '*_pkl.tab'.
    """
    try:
        # Сначала обработка в поддиректориях
        subdirs = getSubDirs(del_dirname)
        if subdirs:
            for sub_dir in subdirs:
                delAllFilesFilter(sub_dir, *del_filter)
        for file_mask in del_filter:
            del_files = getFilesByMask(os.path.join(del_dirname, file_mask))
            for del_file in del_files:
                os.remove(del_file)
                log.info(u'Удаление файла <%s>' % del_file)
        return True
    except:
        log.fatal(u'Ошибка удаления файлов %s из папки <%s>' % (str(del_filter), del_dirname))
        return None


def getPythonDir():
    """
    Папка в которую установлен Python.
    """
    return os.path.dirname(sys.executable)


def getPythonExe():
    """
    Полный путь к исполняемому интерпретатору Python.
    """
    return sys.executable


def getTempDir():
    """
    Временная директория
    """
    return os.environ['TMP']


def getTempFileName(prefix=None):
    """
    Генерируемое имя временного файла
    """
    return tempfile.mkdtemp(getTempDir(), prefix)


def getHomePath():
    """
    Путь к домашней директории.
    :return: Строку-путь до папки пользователя.
    """
    os_platform = platform.uname()[0].lower()
    if os_platform == 'windows':
        home_path = os.environ['HOMEDRIVE']+os.environ['HOMEPATH']
        home_path = home_path.replace('\\', '/')
    elif os_platform == 'linux':
        home_path = os.environ['HOME']
    else:
        log.warning(u'Не поддерживаемая ОС <%s>' % os_platform)
        return None
    return os.path.normpath(home_path)


def getProfilePath(bAutoCreatePath=True):
    """
    Папка профиля программы DEFIS.
    :param bAutoCreatePath: Создать автоматически путь если его нет?
    :return: Путь до ~/.defis
    """
    home_path = getHomePath()
    if home_path:
        profile_path = os.path.join(home_path, ic.config.PROFILE_DIRNAME)
        if not os.path.exists(profile_path) and bAutoCreatePath:
            # Автоматическое создание пути
            try:
                os.makedirs(profile_path)
            except OSError:
                log.fatal(u'Ошибка создания пути профиля <%s>' % profile_path)
        return profile_path
    return '~/.defis'


def getPrjProfilePath(bAutoCreatePath=True):
    """
    Папка профиля прикладного проекта.
    :param bAutoCreatePath: Создать автоматически путь если его нет?
    :return: Путь до ~/.defis/имя_проекта/
    """
    profile_path = getProfilePath(bAutoCreatePath)
    from ic.engine import ic_user

    prj_name = ic_user.getPrjName()
    if prj_name:
        prj_profile_path = os.path.join(profile_path, prj_name)
    else:
        # Если в проект мы не вошли, то просто определяем папку профиля проекта
        # как папку профиля программы
        prj_profile_path = profile_path

    if not os.path.exists(prj_profile_path) and bAutoCreatePath:
        # Автоматическое создание пути
        try:
            os.makedirs(prj_profile_path)
        except OSError:
            log.fatal(u'Ошибка создания пути профиля проекта <%s>' % prj_profile_path)
    return prj_profile_path


def getProjectDir():
    """
    Папка проекта.
    :return: Папка проекта.
    """
    from ic.engine import ic_user
    return ic_user.getPrjDir()


def getRootProjectDir():
    """
    Корневая папка проекта, в которой находяться все папки подсистем проекта.
    :return: Корневая папка проекта, в которой находяться все папки подсистем проекта.
    """
    prj_dir = getProjectDir()
    return os.path.dirname(prj_dir)

