#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль функций работы с ресурсными файлами.
"""

import os
import os.path
import pickle

from ic.std.log import log

from . import textfunc

__version__ = (0, 1, 1, 2)

# Протокол хранения сериализованных объектов модулем cPickle
# ВНИМАНИЕ!!! PICKLE_PROTOCOL = 1,2 использовать нельзя - ресурсы не востанавливаются
PICKLE_PROTOCOL = 0

# Буфер транслированных ресурсных файлов
Buff_readAndEvalFile = {}


def loadResourceFile(filename, replace_dict={}, bRefresh=False, *arg, **kwarg):
    """
    Загрузить ресурс из файла. Функция читает файл и выполняет его.

    :type filename: C{string}
    :param filename: Имя ресурсного файла.
    :type replace_dict: C{dictionary}
    :param replace_dict: Словарь замен.
    :type bRefresh: C{bool}
    :param bRefresh: Признак того, что файл надо перечитать даже если он
        буферезирован.
    """
    obj = None
    filename = filename.strip()
    try:
        # Проверяем есть ли в буфферном файле такой объект, если есть, то его и возвращаем
        if not bRefresh and filename in Buff_readAndEvalFile:
            log.debug(u' '*3+u'[b] '+u'Возвращение файла <%s> из буфера' % filename)
            return Buff_readAndEvalFile[filename]

        nm = os.path.basename(filename)
        pt = nm.find('.')
        if pt >= 0:
            filepcl = os.path.join(os.path.dirname(filename), nm[:pt] + '_pkl' + nm[pt:])
        else:
            filepcl = os.path.join(os.path.dirname(filename), nm +'_pkl')

        # Проверяем нужно ли компилировать данную структуру по следующим признакам:
        # наличие скомпилированного файла, по времени последней модификации.
        try:
            if (os.path.isfile(filepcl) and not os.path.isfile(filename)) or \
                    (os.path.getmtime(filename) < os.path.getmtime(filepcl)):
                # Пытаеся прочитать сохраненную структуру если время последней
                # модификации текстового представления меньше, времени
                # последней модификации транслированного варианта.
                fpcl = None
                try:
                    fpcl = open(filepcl, 'rb')
                    obj = pickle.load(fpcl)
                    fpcl.close()
                    # Сохраняем объект в буфере
                    Buff_readAndEvalFile[filename] = obj
                    log.debug('\t[+] Загрузка из файла <%s>' % filepcl)
                    return obj
                except IOError:
                    log.error('\t[-] Ошибка открытия файла <%s>' % filepcl)
                except:
                    if fpcl:
                        fpcl.close()
        except:
            pass
        # Пытаемся прочитать cPickle, если не удается считаем, что в файле
        # хранится текст. Читаем его, выполняем, полученный объект сохраняем
        # на диске для последующего использования
        if os.path.isfile(filename):
            try:
                fpcl = open(filename, 'rb')
                obj = pickle.load(fpcl)
                fpcl.close()
                # Сохраняем объект в буфере
                Buff_readAndEvalFile[filename] = obj
                log.debug('\t[+] Загружен файл <%s> в PICKLE формате' % filename)
                return obj
            except Exception as msg:
                log.error('\t[*] Не PICKLE формат файла <%s>' % filename)

        # Открываем текстовое представление, если его нет, то создаем его
        f = open(filename, 'rt')
        txt = f.read().replace('\r\n', '\n')
        f.close()
        for key in replace_dict:
            txt = txt.replace(key, replace_dict[key])

        # Выполняем
        obj = eval(txt)
        # Сохраняем объект в буфере
        Buff_readAndEvalFile[filename] = obj

        # Сохраняем транслированный вариант
        fpcl = open(filepcl, 'wb')
        log.debug('Создание файла <%s> в PICKLE формате' % filepcl)
        pickle.dump(obj, fpcl)  # , PICKLE_PROTOCOL)
        fpcl.close()
    except IOError:
        log.error('\t[*] Ошибка открытия файла <%s>' % filename)
        obj = None
    except:
        log.error('\t[*] Ошибка загрузки файла <%s>' % filename)
        obj = None

    return obj


def loadResource(filename):
    """
    Получить ресурс в ресурсном файле.

    :param filename: Полное имя ресурсного файла.
    """
    # Сначала предположим что файл в формате Pickle.
    struct = loadResourcePickle(filename)
    if struct is None:
        # Но если он не в формате Pickle, то скорее всего в тексте.
        struct = loadResourceText(filename)
    if struct is None:
        # Но если не в тексте но ошибка!
        log.warning(u'Ошибка формата файла %s.' % filename)
        return None
    return struct


def loadResourcePickle(filename):
    """
    Получить ресурс из ресурсного файла в формате Pickle.

    :param filename: Полное имя ресурсного файла.
    """
    if os.path.isfile(filename):
        f = None
        try:
            f = open(filename, 'rb')
            struct = pickle.load(f)
            f.close()
            return struct
        except:
            if f:
                f.close()
            log.fatal(u'Ошибка чтения файла <%s>.' % filename)
    else:
        log.warning(u'Файл <%s> не найден.' % filename)
    return None


def loadResourceText(filename):
    """
    Получить ресурс из ресурсного файла в текстовом формате.

    :param filename: Полное имя ресурсного файла.
    """
    if os.path.isfile(filename):
        f = None
        try:
            f = open(filename, 'rt')
            txt = f.read().replace('\r\n', '\n')
            f.close()
            return eval(txt)
        except:
            if f:
                f.close()
            log.fatal(u'Ошибка чтения файла <%s>.' % filename)
    else:
        log.warning(u'Файл <%s> не найден.' % filename)
    return None


def saveResourcePickle(filename, resource):
    """
    Сохранить ресурс в файле в формате Pickle.

    :param filename: Полное имя ресурсного файла.
    :param resource: Словарно-списковая структура спецификации.
    :return: Возвращает результат выполнения операции True/False.
    """
    f = None
    try:
        # Если необходимые папки не созданы, то создать их
        dir_name = os.path.dirname(filename)
        try:
            os.makedirs(dir_name)
        except:
            pass

        f = open(filename, 'wb')
        pickle.dump(resource, f)
        f.close()
        log.info(u'Файл <%s> сохранен в формате Pickle.' % filename)
        return True
    except:
        if f:
            f.close()
        log.fatal(u'Ошибка сохраненения файла <%s> в формате Pickle.' % filename)

    return False


def saveResourceText(filename, resource):
    """
    Сохранить ресурс в файле в текстовом формате.

    :param filename: Полное имя ресурсного файла.
    :param resource: Словарно-списковая структура спецификации.
    :return: Возвращает результат выполнения операции True/False.
    """
    f = None
    try:
        # Если необходимые папки не созданы, то создать их
        dir_name = os.path.dirname(filename)
        try:
            os.makedirs(dir_name)
        except:
            pass

        f = open(filename, 'wt')
        text = textfunc.StructToTxt(resource)
        f.write(text)
        f.close()
        log.info(u'Файл <%s> сохранен в текстовом формате.' % filename)
        return True
    except:
        if f:
            f.close()
        log.fatal(u'Ошибка сохраненения файла <%s> в текстовом формате.' % filename)
    return False
