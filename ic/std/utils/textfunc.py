#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль функций обработки текста.
"""

from ..log import log

__version__ = (0, 1, 1, 2)

# Наполнитель позиций при отображении вложенности пунктов в компоненте списка
PADDING = '    '

DEFAULT_ENCODING = 'utf-8'


def StructToTxt(struct, level=0):
    """
    Перевод словарно-списковой структуры в форматированный текст.
    @param struct : словарно-списковая структура.
    @param level: уровень вложенности (д.б. 0).
    """
    try:
        txt = ''
        obj_type = type(struct)
        if isinstance(obj_type, list):
            txt = txt +'\n' + level * PADDING + '[\n'
            for obj in struct:
                txt += level * PADDING
                txt += StructToTxt(obj, level + 1)
                txt += ',\n'
            if len(struct) != 0:
                txt = txt[:-2]
            txt = txt +'\n' + level * PADDING + ']'
        elif isinstance(obj_type, dict):
            txt = txt +'\n' + level * PADDING + '{\n'
            keys = struct.keys()
            values = struct.values()
            for key in keys:
                txt = txt + level * PADDING + '\'' + key + '\':'
                txt += StructToTxt(struct[key], level + 1)
                txt += ',\n'
            if len(keys) != 0:
                txt = txt[:-2]
            txt = txt +'\n' + level * PADDING + '}'
        elif isinstance(obj_type, str):
            # Появляется косяк с разделителем папок в именах путей
            # Проверка на кавычки
            txt = txt +'\'' + struct.replace('\'', '\\\'').replace('\'', '\\\'').replace('\r', '\\r').replace('\n', '\\n').replace('\t', '\\t') + '\''
        else:
            txt += str(struct)

        # Убрать первый перевод каретки
        if txt[0] == '\n' and (not level):
            txt = txt[1:]
        return txt
    except:
        log.fatal(u'Ошибка перевода структуры в форматированный текст. Уровень <%d>' % level)
    return u''


rusRegUpperDict = {'а': 'А', 'б': 'Б', 'в': 'В', 'г': 'Г', 'д': 'Д', 'е': 'Е', 'ё': 'Ё', 'ж': 'Ж',
                   'з': 'З', 'и': 'И', 'й': 'Й', 'к': 'К', 'л': 'Л', 'м': 'М', 'н': 'Н', 'о': 'О', 'п': 'П',
                   'р': 'Р', 'с': 'С', 'т': 'Т', 'у': 'У', 'ф': 'Ф', 'х': 'Х', 'ц': 'Ц', 'ч': 'Ч',
                   'ш': 'Ш', 'щ': 'Щ', 'ь': 'Ь', 'ы': 'Ы', 'ъ': 'Ъ', 'э': 'Э', 'ю': 'Ю', 'я': 'Я'}


def toUpper(text):
    """
    Тупой перевод к верхнему регистру русских букв.
    """
    pyUpper = text.upper()
    upper_str = list(pyUpper)
    upper_str = [rusRegUpperDict.setdefault(pyUpper[ch[0]], ch[1]) for ch in enumerate(upper_str)]
    return ''.join(upper_str)


rusRegLowerDict = {'А': 'а', 'Б': 'б', 'В': 'в', 'Г': 'г', 'Д': 'д', 'Е': 'е', 'Ё': 'ё', 'Ж': 'ж',
                   'З': 'з', 'И': 'и', 'Й': 'й', 'К': 'к', 'Л': 'л', 'М': 'м', 'Н': 'н', 'О': 'о', 'П': 'п',
                   'Р': 'р', 'С': 'с', 'Т': 'т', 'У': 'у', 'Ф': 'ф', 'Х': 'х', 'Ц': 'ц', 'Ч': 'ч',
                   'Ш': 'ш', 'Щ': 'щ', 'Ь': 'ь', 'Ы': 'ы', 'Ъ': 'ъ', 'Э': 'э', 'Ю': 'ю', 'Я': 'я'}

u_rusRegLowerDict = {u'А': u'а', u'Б': u'б', u'В': u'в', u'Г': u'г', u'Д': u'д', u'Е': u'е', u'Ё': u'ё', u'Ж': u'ж',
                     u'З': u'з', u'И': u'и', u'Й': u'й', u'К': u'к', u'Л': u'л', u'М': u'м', u'Н': u'н', u'О': u'о',
                     u'П': u'п', u'Р': u'р', u'С': u'с', u'Т': u'т', u'У': u'у', u'Ф': u'ф', u'Х': u'х', u'Ц': u'ц',
                     u'Ч': u'ч', u'Ш': u'ш', u'Щ': u'щ', u'Ь': u'ь', u'Ы': u'ы', u'Ъ': u'ъ', u'Э': u'э', u'Ю': u'ю',
                     u'Я': u'я'}

rusRegLowerLst = ['а', 'б', 'в', 'г', 'д', 'е', 'ё', 'ж',
                  'з', 'и', 'й', 'к', 'л', 'м', 'н', 'о', 'п',
                  'р', 'с', 'т', 'у', 'ф', 'х', 'ц', 'ч',
                  'ш', 'щ', 'ь', 'ы', 'ъ', 'э', 'ю', 'я']

u_rusRegLowerLst = [u'а', u'б', u'в', u'г', u'д', u'е', u'ё', u'ж',
                    u'з', u'и', u'й', u'к', u'л', u'м', u'н', u'о', u'п',
                    u'р', u'с', u'т', u'у', u'ф', u'х', u'ц', u'ч',
                    u'ш', u'щ', u'ь', u'ы', u'ъ', u'э', u'ю', u'я']

rusRegUpperLst = ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ё', 'Ж',
                  'З', 'И', 'Й', 'К', 'Л', 'М', 'Н', 'О', 'П',
                  'Р', 'С', 'Т', 'У', 'Ф', 'Х', 'Ц', 'Ч',
                  'Ш', 'Щ', 'Ь', 'Ы', 'Ъ', 'Э', 'Ю', 'Я']

u_rusRegUpperLst = [u'А', u'Б', u'В', u'Г', u'Д', u'Е', u'Ё', u'Ж',
                    u'З', u'И', u'Й', u'К', u'Л', u'М', u'Н', u'О', u'П',
                    u'Р', u'С', u'Т', u'У', u'Ф', u'Х', u'Ц', u'Ч',
                    u'Ш', u'Щ', u'Ь', u'Ы', u'Ъ', u'Э', u'Ю', u'Я']


def toLower(text):
    """
    Тупой перевод к нижнему регистру русских букв.
    """
    pyLower = text.lower()
    lower_str = list(pyLower)
    lower_str = [rusRegLowerDict.setdefault(pyLower[ch[0]], ch[1]) for ch in enumerate(lower_str)]
    return ''.join(lower_str)


def cmpLower(text1, text2):
    """
    Сравнивает два символа в нижнем регистре.
    """
    if text1 in rusRegLowerLst and text2 in rusRegLowerLst:
        p1 = rusRegLowerLst.index(text1)
        p2 = rusRegLowerLst.index(text2)

        if p1 > p2:
            return -1
        elif p1 < p2:
            return 1
        else:
            return 0
    else:
        if text1 > text2:
            return -1
        elif text1 < text2:
            return 1
        else:
            return 0


def cmpLowerU(str1, str2):
    """
    Сравнивает два символа в нижнем регистре.
    """
    for i in range(min(len(str1), len(str2))):
        s1 = str1[i]
        s2 = str2[i]
        if s1 in u_rusRegLowerLst and s2 in u_rusRegLowerLst:
            p1 = u_rusRegLowerLst.index(s1)
            p2 = u_rusRegLowerLst.index(s2)

            if p1 > p2:
                return -1
            elif p1 < p2:
                return 1
            else:
                return 0
        else:
            if s1 > s2:
                return -1
            elif s1 < s2:
                return 1
    if len(str1) > len(str2):
        return -1
    elif len(str1) < len(str2):
        return 1

    return 0


def text2unicode(text, txt_codepage=DEFAULT_ENCODING):
    """
    Перекодировка строки в юникод с проверкой типа входного аргумента.
    @param text: Строка.
    @param txt_codepage: Кодовая страница строки.
    @return: Строка в юникоде.
    """
    if isinstance(text, str):
        return text
    elif isinstance(text, bytes):
        return text.decode(txt_codepage)
    return str(text)


def isLATText(text):
    """
    Текст написан в латинице?
    """
    if isinstance(text, str):
        rus_chr = [c for c in text if ord(c) > 128]
        return not bool(rus_chr)
    # Это не строка
    return False


def isRUSText(text):
    """
    Строка с рускими буквами?
    """
    if isinstance(text, str):
        rus_chr = [c for c in text if ord(c) > 128]
        return bool(rus_chr)
    # Это не строка
    return False


def _rus2lat(text, translate_dict):
    """
    Перевод русских букв в латинские по словарю замен.
    """
    if isinstance(text, bytes):
        # Привести к юникоду
        text = text.decode(DEFAULT_ENCODING)

    txt_list = list(text)
    txt_list = [translate_dict.setdefault(ch, ch) for ch in txt_list]
    return ''.join(txt_list)


RUS2LATDict = {u'а': 'a', u'б': 'b', u'в': 'v', u'г': 'g', u'д': 'd', u'е': 'e', u'ё': 'yo', u'ж': 'j',
               u'з': 'z', u'и': 'idx', u'й': 'y', u'к': 'k', u'л': 'l', u'м': 'm', u'н': 'n', u'о': 'o', u'п': 'p',
               u'р': 'r', u'с': 'text', u'т': 't', u'у': 'u', u'ф': 'f', u'х': 'h', u'ц': 'c', u'ч': 'ch',
               u'ш': 'sh', u'щ': 'sch', u'ь': '', u'ы': 'y', u'ъ': '', u'э': 'e', u'ю': 'yu', u'я': 'ya',
               u'А': 'A', u'Б': 'B', u'В': 'V', u'Г': 'G', u'Д': 'D', u'Е': 'E', u'Ё': 'YO', u'Ж': 'J',
               u'З': 'Z', u'И': 'I', u'Й': 'Y', u'К': 'K', u'Л': 'L', u'М': 'M', u'Н': 'N', u'О': 'O', u'П': 'P',
               u'Р': 'R', u'С': 'S', u'Т': 'T', u'У': 'U', u'Ф': 'F', u'Х': 'H', u'Ц': 'C', u'Ч': 'CH',
               u'Ш': 'SH', u'Щ': 'SCH', u'Ь': '', u'Ы': 'Y', u'Ъ': '', u'Э': 'E', u'Ю': 'YU', u'Я': 'YA'}


def rus2lat(text):
    """
    Перевод русских букв в латинские.
    """
    return _rus2lat(text, RUS2LATDict)


RUS2LATKeyboardDict = {u'а': 'f', u'б': '_', u'в': 'd', u'г': 'u', u'д': 'l', u'е': 't', u'ё': '_', u'ж': '_',
                       u'з': 'p', u'и': 'b', u'й': 'q', u'к': 'r', u'л': 'k', u'м': 'v', u'н': 'y', u'о': 'j',
                       u'п': 'g', u'р': 'h', u'с': 'c', u'т': 'n', u'у': 'e', u'ф': 'a', u'х': '_', u'ц': 'w',
                       u'ч': 'x', u'ш': 'idx', u'щ': 'o', u'ь': 'm', u'ы': 'text', u'ъ': '_', u'э': '_', u'ю': '_',
                       u'я': 'z', u'А': 'F', u'Б': '_', u'В': 'D', u'Г': 'U', u'Д': 'L', u'Е': 'T', u'Ё': '_',
                       u'Ж': '_', u'З': 'P', u'И': 'B', u'Й': 'Q', u'К': 'R', u'Л': 'K', u'М': 'V', u'Н': 'Y',
                       u'О': 'J', u'П': 'G', u'Р': 'H', u'С': 'C', u'Т': 'N', u'У': 'E', u'Ф': 'A', u'Х': '_',
                       u'Ц': 'W', u'Ч': 'X', u'Ш': 'I', u'Щ': 'O', u'Ь': 'M', u'Ы': 'S', u'Ъ': '_', u'Э': '_',
                       u'Ю': '_', u'Я': 'Z'}


def rus2lat_keyboard(text):
    """
    Перевод русских букв в латинские по раскладке клавиатуры.
    """
    return _rus2lat(text, RUS2LATKeyboardDict)


rus_encodings = {'UTF-8':      'utf-8',
                 'CP1251':     'windows-1251',
                 'KOI8-R':     'koi8-r',
                 'IBM866':     'ibm866',
                 'ISO-8859-5': 'iso-8859-5',
                 'MAC':        'mac',
                 }


def get_codepage(text=None):
    """
    Определение кодировки текста.
    Пример вызова функции:
    rus_encodings[get_codepage(file('test.txt').read())]
    Есть альтернативный вариант определения кодировки (с помощью chardet):
    a = 'sdfds'
    import chardet
    chardet.detect(a)
    {'confidence': 1.0, 'encoding': 'ascii'}
    a = 'авыаыв'
    chardet.detect(a)
    {'confidence': 0.99, 'encoding': 'utf-8'}
    """
    uppercase = 1
    lowercase = 3
    utfupper = 5
    utflower = 7
    codepages = {}
    for enc in rus_encodings.keys():
        codepages[enc] = 0
    if text is not None and len(text) > 0:
        last_simb = 0
        for simb in text:
            simb_ord = ord(simb)

            # non-russian characters
            if simb_ord < 128 or simb_ord > 256:
                continue

            # UTF-8
            if last_simb == 208 and (143 < simb_ord < 176 or simb_ord == 129):
                codepages['UTF-8'] += (utfupper * 2)
            if (last_simb == 208 and (simb_ord == 145 or 175 < simb_ord < 192)) \
               or (last_simb == 209 and (127 < simb_ord < 144)):
                codepages['UTF-8'] += (utflower * 2)

            # CP1251
            if 223 < simb_ord < 256 or simb_ord == 184:
                codepages['CP1251'] += lowercase
            if 191 < simb_ord < 224 or simb_ord == 168:
                codepages['CP1251'] += uppercase

            # KOI8-R
            if 191 < simb_ord < 224 or simb_ord == 163:
                codepages['KOI8-R'] += lowercase
            if 222 < simb_ord < 256 or simb_ord == 179:
                codepages['KOI8-R'] += uppercase

            # IBM866
            if 159 < simb_ord < 176 or 223 < simb_ord < 241:
                codepages['IBM866'] += lowercase
            if 127 < simb_ord < 160 or simb_ord == 241:
                codepages['IBM866'] += uppercase

            # ISO-8859-5
            if 207 < simb_ord < 240 or simb_ord == 161:
                codepages['ISO-8859-5'] += lowercase
            if 175 < simb_ord < 208 or simb_ord == 241:
                codepages['ISO-8859-5'] += uppercase

            # MAC
            if 221 < simb_ord < 255:
                codepages['MAC'] += lowercase
            if 127 < simb_ord < 160:
                codepages['MAC'] += uppercase

            last_simb = simb_ord

        idx = ''
        max_cp = 0
        for item in codepages:
            if codepages[item] > max_cp:
                max_cp = codepages[item]
                idx = item
        return idx


def toUnicode(value, txt_codepage=DEFAULT_ENCODING):
    """
    Преобразовать любое значение в юникод.
    @param value: Значение.
    @param txt_codepage: Кодовая страница для строк.
    """
    if isinstance(value, str):
        return value
    elif isinstance(value, bytes):
        return value.decode(txt_codepage)
    return str(value)


def txt_find_words(txt, *words):
    """
    Поиск слов в тексте.
    Поиск ведется до первого нахождения одного из указанных слов.
    @param txt: Анализируемый текст.
    @param words: Искомые слова.
    @return: True (есть такие слова в тексте)/False (слова не найдены).
    """
    if not isinstance(txt, str):
        txt = toUnicode(txt)
    find = False
    for word in words:
        find = word in txt
        if find:
            break
    return find


def is_serial_symbol(txt, symbol):
    """
    Проверка на то что текст представляет из себя
    последовательность из одного конкретного символа.
    @param txt: Текст.
    @param symbol: Символ.
    @return: True/False.
    """
    if not txt:
        # Если это пустая строка то это
        # вообще не последовательность
        return False

    result = True
    for symb in txt:
        result = result and (symb == symbol)
    return result


def is_serial(txt):
    """
    Проверка на то что текст представляет из себя
    последовательность из одного символа.
    @param txt: Текст.
    @return: True/False.
    """
    return is_serial_symbol(txt, txt[0])


def is_serial_zero(txt):
    """
    Проверка на то что текст представляет из себя
    последовательность из одного символа '0'.
    @param txt: Текст.
    @return: True/False.
    """
    return is_serial_symbol(txt, '0')


def getNumEnding(number, endings):
    """
    Функция возвращает окончание для множественного числа слова на основании числа и
    массива окончаний.
    Функция взята с https://habrahabr.ru/post/105428/.
    @param number: Число на основе которого нужно сформировать окончание
    @param endings: Массив слов или окончаний для чисел(1, 4, 5),
        например ['яблоко', 'яблока', 'яблок']
    @return: Строку.
    """
    number %= 100
    if 11 <= number <= 19:
        sEnding = endings[2]
    else:
        i = number % 10
        if i == 1:
           sEnding = endings[0]
        elif i in (2, 3, 4):
            sEnding = endings[1]
        else:
            sEnding = endings[2]
    return sEnding
