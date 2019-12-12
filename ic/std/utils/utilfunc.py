#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль сервисных функций программиста
"""

# --- Подключение библиотек ---
import sys
import inspect
import wx

from encodings.aliases import aliases

from ic.std.log import log
try:
    import win32api
    import win32con
    import regutil
except ImportError:
    log.error(u'Ошибка импорта PyWin32', bForcePrint=True)

try:
    import DevId
except:
    log.error(u'Ошибка импорта DevId', bForcePrint=True)


__version__ = (0, 1, 1, 2)


# --- Функции ---
def isSubClass(class1, class2):
    """
    Функция определяет является ли Class1_ базовым для class2. Проверка
    производится рекурсивно.
    :param class1: Объект класса.
    :param class2: Объект класса.
    :return: Возвращает результат отношения (1/0).
    """
    return issubclass(class2, class1)


def getAttrValue(attr_name, spc):
    """
    Получить нормированное значение свойства из спецификации.
    :param attr_name: Имя атрибута.
    :param spc: Спецификация.
    """
    try:
        # Нормализация по типам
        if isinstance(spc[attr_name], str):
            try:
                # Возможно это не строка
                value = eval(spc[attr_name])
            except:
                # Нет это все таки строка
                value = spc[attr_name]
            spc[attr_name] = value
        return spc[attr_name]
    except:
        log.fatal()
        return None


def getStrInQuotes(value):
    """
    Если value - строка то обрамляет ее одинарными кавычками, если нет,
    то просто преабразует в строку.
    """
    if isinstance(value, str):
        return '\'%s\'' % value
    else:
        return str(value)


def KeyToText(key, bShift=True, bAlt=True, bCtrl=True):
    """
    Функция преабразует код клавиши в текстовый эквивалент.
    :param key: код клавиши.
    :param bShift: флаг клавиши Shift.
    :param bAlt: флаг клавиши Alt.
    :param bCtrl: флаг клавиши Ctrl.
    :return: Возвращает текстовую строку, например 'Alt+X'.
    """
    result = ''
    if bShift:
        result += 'Shift+'
    if bAlt:
        result += 'Alt+'
    if bCtrl:
        result += 'Ctrl+'
    if key == wx.WXK_RETURN:
        result += 'Enter'
    elif key == wx.WXK_ESCAPE:
        result += 'Esc'
    elif key == wx.WXK_DELETE:
        result += 'Del'
    elif key == wx.WXK_TAB:
        result += 'Tab'
    elif key == wx.WXK_SPACE:
        result += 'Space'
    elif key == wx.WXK_END:
        result += 'End'
    elif key == wx.WXK_HOME:
        result += 'Home'
    elif key == wx.WXK_PAUSE:
        result += 'Pause'
    elif key == wx.WXK_LEFT:
        result += 'Left'
    elif key == wx.WXK_UP:
        result += 'Up'
    elif key == wx.WXK_RIGHT:
        result += 'Right'
    elif key == wx.WXK_DOWN:
        result += 'Down'
    elif key == wx.WXK_INSERT:
        result += 'Ins'
    elif key == wx.WXK_F1:
        result += 'F1'
    elif key == wx.WXK_F2:
        result += 'F2'
    elif key == wx.WXK_F3:
        result += 'F3'
    elif key == wx.WXK_F4:
        result += 'F4'
    elif key == wx.WXK_F5:
        result += 'F5'
    elif key == wx.WXK_F6:
        result += 'F6'
    elif key == wx.WXK_F7:
        result += 'F7'
    elif key == wx.WXK_F8:
        result += 'F8'
    elif key == wx.WXK_F9:
        result += 'F9'
    elif key == wx.WXK_F10:
        result += 'F10'
    elif key == wx.WXK_F11:
        result += 'F11'
    elif key == wx.WXK_F12:
        result += 'F12'
    elif key == wx.WXK_F13:
        result += 'F13'
    elif key == wx.WXK_F14:
        result += 'F14'
    elif key == wx.WXK_F15:
        result += 'F15'
    elif key == wx.WXK_F16:
        result += 'F16'
    elif key == wx.WXK_F17:
        result += 'F17'
    elif key == wx.WXK_F18:
        result += 'F18'
    elif key == wx.WXK_F19:
        result += 'F19'
    elif key == wx.WXK_F20:
        result += 'F20'
    elif key == wx.WXK_F21:
        result += 'F21'
    elif key == wx.WXK_F22:
        result += 'F22'
    elif key == wx.WXK_F23:
        result += 'F23'
    elif key == wx.WXK_F24:
        result += 'F24'
    else:
        if key < 256:
            result += chr(key)
            
    return result


def delKeyInDictTree(dictionary, key):
    """
    Функция удаляет из словаря рекурсивно все указанные ключи.
    :param dictionary: Непосредственно словарь или список.
    :param key: Ключ, который необходимо удалить.
    """
    # Если это у нас словарь, то ...
    if type(dictionary) is dict:
        try:
            # Сначала удаляем ключ на этом уровне.
            del dictionary[key]
        except:
            pass
        # Затем спускаемся на уровень ниже и обрабатываем
        for item in dictionary.values():
            delKeyInDictTree(item, key)
    # а если список, то перебираем элементы
    elif type(dictionary) is list:
        for item in dictionary:
            delKeyInDictTree(item, key)
    # а если не то и не другое, то ничего с ним не делать
    else:
        return


def setKeyInDictTree(dictionary, key, value):
    """
    Функция устанавливает значенеи ключа в словаре рекурсивно.
    :param dictionary: Непосредственно словарь или список.
    :param key: Ключ, который необходимо установить.
    :param value: Значение ключа.
    """
    # Если это у нас словарь, то ...
    if type(dictionary) is dict:
        try:
            # Сначала устанавливаем ключ на этом уровне.
            dictionary[key] = value
        except:
            pass
        # Затем спускаемся на уровень ниже и обрабатываем
        for item in dictionary.values():
            setKeyInDictTree(item, key, value)
    # а если список, то перебираем элементы
    elif type(dictionary) is list:
        for item in dictionary:
            setKeyInDictTree(item, key, value)
    # а если не то и не другое, то ничего с ним не делать
    else:
        return


def setObjToGlobal(objname, obj):
    """
    Поместить объект в глобальное пространство имен.
    :param objname: Имя объекта в глобальном пространстве имен.
    :param obj: Сам объект.
    """
    globals()[objname] = obj


def recodeText(text, txt_codepage, new_codepage):
    """
    Перекодировать из одной кодировки в другую.
    :param text: Строка.
    :param txt_codepage: Кодовая страница строки.
    :param new_codepage: Новая кодовая страница строки.
    """
    if new_codepage.upper() == 'UNICODE':
        # Кодировка в юникоде.
        return text

    if new_codepage.upper() == 'OCT' or new_codepage.upper() == 'HEX':
        # Закодировать строку в восьмеричном/шестнадцатеричном виде.
        return toOctHexText(text, new_codepage)

    string = u''
    if isinstance(text, str):
        string = str(text)   # , txt_codepage)

    return string.encode(new_codepage)


def toOctHexText(text, coding):
    """
    Закодировать строку в восьмеричном/шестнадцатеричном виде.
    Символы с кодом < 128 не кодируются.
    :param text:
    :param coding: Кодировка 'OCT'-восьмеричное представление.
                            'HEX'-шестнадцатеричное представление.
    :return: Возвращает закодированную строку.
    """
    try:
        if coding.upper() == 'OCT':
            fmt = '\\%o'
        elif coding.upper() == 'HEX':
            fmt = '\\x%x'
        else:
            # Ошибка аргументов
            log.warning(u'Argument error in toOctHexText.')
            return None
        # Перебор строки по символам
        ret_str = ''
        for char in text:
            code_char = ord(char)
            # Символы с кодом < 128 не кодируются.
            if code_char > 128:
                ret_str += fmt % code_char
            else:
                ret_str += char
        return ret_str
    except:
        log.fatal()
        return None


def recodeListStrings(cur_list, txt_codepage, new_codepage):
    """
    Перекодировать все строки в списке рекурсивно в другую кодировку.
    Перекодировка производится также внутри вложенных словарей и кортежей.
    :param cur_list: Сам список.
    :param txt_codepage: Кодовая страница строки.
    :param new_codepage: Новая кодовая страница строки.
    :return: Возвращает преобразованный список.
    """
    lst = []
    # Перебор всех элементов списка
    for i in range(len(cur_list)):
        if isinstance(cur_list[i], list):
            # Элемент - список
            value = recodeListStrings(cur_list[i], txt_codepage, new_codepage)
        elif isinstance(cur_list[i], dict):
            # Элемент списка - словарь
            value = recodeDictStrings(cur_list[i], txt_codepage, new_codepage)
        elif isinstance(cur_list[i], tuple):
            # Элемент списка - кортеж
            value = recodeTupleStrings(cur_list[i], txt_codepage, new_codepage)
        elif isinstance(cur_list[i], str):
            value = recodeText(cur_list[i], txt_codepage, new_codepage)
        else:
            value = cur_list[i]
        lst.append(value)
    return lst


def _isRUSText(text):
    """
    Строка с рускими буквами?
    """
    if isinstance(text, str):
        rus_chr = [c for c in text if ord(c) > 128]
        return bool(rus_chr)
    return False


def recodeDictStrings(dictionary, txt_codepage, new_codepage):
    """
    Перекодировать все строки в словаре рекурсивно в другую кодировку.
    Перекодировка производится также внутри вложенных словарей и кортежей.
    :param dictionary: Сам словарь.
    :param txt_codepage: Кодовая страница строки.
    :param new_codepage: Новая кодовая страница строки.
    :return: Возвращает преобразованный словарь.
    """
    keys_ = dictionary.keys()
    # Перебор всех ключей словаря
    for cur_key in keys_:
        value = dictionary[cur_key]
        # Нужно ключи конвертировать?
        if _isRUSText(cur_key):
            new_key = recodeText(cur_key, txt_codepage, new_codepage)
            del dictionary[cur_key]
        else:
            new_key = cur_key
            
        if isinstance(value, list):
            # Элемент - список
            dictionary[new_key] = recodeListStrings(value, txt_codepage, new_codepage)
        elif isinstance(value, dict):
            # Элемент - cловарь
            dictionary[new_key] = recodeDictStrings(value, txt_codepage, new_codepage)
        elif isinstance(value, tuple):
            # Элемент - кортеж
            dictionary[new_key] = recodeTupleStrings(value, txt_codepage, new_codepage)
        elif isinstance(value, str):
            dictionary[new_key] = recodeText(value, txt_codepage, new_codepage)

    return dictionary


def recodeTupleStrings(cur_tuple, txt_codepage, new_codepage):
    """
    Перекодировать все строки в кортеже рекурсивно в другую кодировку.
    Перекодировка производится также внутри вложенных словарей и кортежей.
    :param cur_tuple: Сам кортеж.
    :param txt_codepage: Кодовая страница строки.
    :param new_codepage: Новая кодовая страница строки.
    :return: Возвращает преобразованный кортеж.
    """
    # Перевести кортеж в список
    lst = list(cur_tuple)
    # и обработать как список
    list_ = recodeListStrings(lst, txt_codepage, new_codepage)
    # Обратно перекодировать
    return tuple(list_)


def recodeStructStrings(struct, txt_codepage, new_codepage):
    """
    Перекодировать все строки в структуре рекурсивно в другую кодировку.
    :param struct: Сруктура (список, словарь, кортеж).
    :param txt_codepage: Кодовая страница строки.
    :param new_codepage: Новая кодовая страница строки.
    :return: Возвращает преобразованную структру.
    """
    if isinstance(struct, list):
        # Список
        struct = recodeListStrings(struct, txt_codepage, new_codepage)
    elif isinstance(struct, dict):
        # Словарь
        struct = recodeDictStrings(struct, txt_codepage, new_codepage)
    elif isinstance(struct, tuple):
        # Кортеж
        struct = recodeTupleStrings(struct, txt_codepage, new_codepage)
    elif isinstance(struct, str):
        # Строка
        struct = recodeText(struct, txt_codepage, new_codepage)
    else:
        # Тип не определен
        struct = struct
    return struct


def list2str(cur_list, delimeter):
    """
    Конвертация списка в строку с разделением символом разделителя.
    :param cur_list: Список.
    :param delimeter: Символ разделителя.
    :return: Возвращает сформированную строку.
    """
    return delimeter.join(cur_list)


def getHDDSerialNo():
    """
    Определить серийный номер HDD.
    """
    try:
        hdd_info = DevId.GetHDDInfo()
        return hdd_info[2]
    except:
        # Ошибка определения серийного номера HDD.
        log.fatal(u'Ошибка определения серийного номера HDD')
        return ''


def getRegValue(reg_key, reg_value=None):
    """
    Взять информацию из реестра относительно данного проекта.
    :param reg_key: Ключ реестра.
    :param reg_value: Имя значения из реестра.
    """
    hkey = None
    try:
        hkey = win32api.RegOpenKey(win32con.HKEY_LOCAL_MACHINE, reg_key)
        value = win32api.RegQueryValueEx(hkey, reg_value)
        win32api.RegCloseKey(hkey)
        return value[0]
    except:
        if hkey:
            win32api.RegCloseKey(hkey)
        # Ошибка определения информации из реестра.
        log.fatal()
    return None


def findChildResByName(children_res, child_name):
    """
    Поиск ресурсного описания дочернего объекта по имени.
    :param children_res: Список ресурсов-описаний дечерних объектов.
    :return child_name: Имя искомого дочернего объекта.
    :return: Индекс ресурсного описания в списке, если
        описания с таким именем не найдено, то возвращается -1.
    """
    try:
        children_names = [child['name'] for child in children_res]
        return children_names.index(child_name)
    except ValueError:
        return -1


def getFuncListInModule(module=None):
    """
    Получить список имен функций в модуле.
    :param module: Объект модуля. 
        Для использования модуль д.б. импортирован.
    :return: Возвращает список кортежей:
        [(имя функции, описание функции, объект функции),...]
    """
    if module:
        return [(func_name, module.__dict__[func_name].__doc__,
                 module.__dict__[func_name]) for func_name in [f_name for f_name in module.__dict__.keys() if inspect.isfunction(module.__dict__[f_name])]]
    return None


def isOSWindowsPlatform():
    """
    Функция определения ОС.
    :return: True-если ОС-Windows и False во всех остальных случаях.
    """
    return bool(sys.platform[:3].lower() == 'win')


def get_encodings_list():
    """
    Список возможных кодировок.
    """
    try:
        encoding_list = [code for code in aliases.values()]
        encoding_list = [encoding for i, encoding in enumerate(encoding_list) if encoding not in encoding_list[:i]]
        result = encoding_list
        result.sort()
        return result
    except:
        return ['UTF-8', 'UTF-16', 'CP1251', 'CP866', 'KOI8-R']


def encode_unicode_struct(struct, dst_codepage='utf-8'):
    """
    Перекодировать все строки unicode структуры в  указанную кодировку.
    """
    return recodeStructStrings(struct, 'UNICODE', dst_codepage)
