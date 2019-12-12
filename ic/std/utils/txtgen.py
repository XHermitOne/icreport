#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Генератор текста по контексту. 
В качестве контекста может выступать любая словарная структура.
Простой аналог генератора страниц в Django.
Синтаксис такой же, что в дальнейшем позволит испльзовать 
Django в качестве генератора.
"""

import re
import os
import os.path

from ic.std.log import log
from . import textfunc

__version__ = (0, 1, 1, 1)

VAR_PATTERN = r'(\{\{.*?\}\})'

DEFAULT_ENCODING = 'utf-8'
FIND_REPLACEMENT_ERR = u'!!!Замена не определена в контексте!!!'

REPLACE_NAME_START = u'{{'
REPLACE_NAME_END = u'}}'


def gen(sTxt, dContext=None):
    """
    Генерация текста.

    :param sTxt: Tекст шаблона.
    :param dContext. Контекст.
        В качестве контекста может выступать любая словарная структура.
        По умолчанию контекст - локальное пространство имен модуля config.
    :return: Сгенерированный текст.
    """
    if dContext is None:
        dContext = {}

        try:
            from ic import config
        except ImportError:
            import config

        for name in config.__dict__.keys():
            dContext[name] = config.get_cfg_var(name)
    return auto_replace(sTxt, dContext)


def _getVarName(sPlace):
    """
    Определить имя переменной замены.
    """
    return sPlace.replace('{', u'').replace('}', u'').strip()


def get_raplace_names(sTxt):
    """
    Определить имена автозамен.

    :param sTxt: Редактируемый текст.
    :return: Список имен замен.
    """
    if sTxt is None:
        log.warning(u'Не определен текст для автозамен')
        return list()
    replaces = re.findall(VAR_PATTERN, sTxt)
    return [replace_name[len(REPLACE_NAME_START):-len(REPLACE_NAME_END)].strip() for replace_name in replaces]


def auto_replace(sTxt, dReplaces=None):
    """
    Запуск автозамен из контекста.

    :param sTxt: Редактируемый текст.
    :param dReplaces: Словарь замен, если None, то берется locals().
    :return: Возвращается отредактированный текст или None в
        случае возникновения ошибки.
    """
    if dReplaces is None:
        dReplaces = locals()
            
    if sTxt and (dReplaces is not None):

        replace_places = re.findall(VAR_PATTERN, sTxt)
        replaces = dict([(place, dReplaces.get(_getVarName(place), FIND_REPLACEMENT_ERR)) for place in replace_places])
            
        for place, value in replaces.items():
            place = textfunc.toUnicode(place, DEFAULT_ENCODING)
            value = textfunc.toUnicode(value, DEFAULT_ENCODING)
            try:
                log.debug(u'Автозамена %s -> <%s>' % (place, value))
            except UnicodeEncodeError:
                log.error(u'Ошибка отображения автозамены %s' % place)
            except UnicodeDecodeError:
                log.error(u'Ошибка отображения автозамены %s' % place)

            sTxt = sTxt.replace(place, value)
            
        return sTxt
    elif sTxt is None:
        log.warning(u'Не определен текст для автозамен')
    elif dReplaces is None:
        log.warning(u'Не определены замены для автозамен')
    return None


def is_genered(txt):
    """
    Проверка является ли текст генерируемым.
        Т.е. есть ли в нем необходимые замены.

    :param txt: Тест.
    :return: True - есть замены, False - замен нет.
    """
    if not isinstance(txt, str):
        return False

    if isinstance(txt, bytes):
        # Для корректной проверки необходимо преобразовать в Unicode
        txt = txt.decode(DEFAULT_ENCODING)
    return REPLACE_NAME_START in txt and REPLACE_NAME_END in txt


def gen_txt_file(sTxtTemplateFilename, sTxtOutputFilename, dContext=None, output_encoding=None):
    """
    Генерация текстового файла по шаблону.

    :param sTxtTemplateFilename: Шаблон - текстовый файл.
    :param sTxtOutputFilename: Наименование выходного текстового файла.
    :param dContext. Контекст.
        В качестве контекста может выступать любая словарная структура.
        По умолчанию контекст - локальное пространство имен модуля config.
    :param output_encoding: Кодовая страница результирующего файла.
        Если не определена, то кодовая страница остается такая же как и у шаблона.
    :return: True - генерация прошла успешно,
        False - ошибка генерации.
    """
    template_file = None
    output_file = None

    template_filename = os.path.abspath(sTxtTemplateFilename)
    if not os.path.exists(template_filename):
        log.warning(u'Файл шаблона для генерации текстового файла <%s> не найден' % template_filename)
        return False

    # Чтение шаблона из файла
    try:
        template_file = open(template_filename, 'rt')
        template_txt = template_file.read()
        template_file.close()
    except:
        if template_file:
            template_file.close()
        log.fatal(u'Ошибка чтения шаблона из файла <%s>' % template_filename)
        return False

    try:
        # Определить кодовую страницу текста
        template_encoding = textfunc.get_codepage(template_txt)
        log.debug(u'Кодовая страница шаблона <%s>' % template_encoding)

        # Шаблон необходимо проебразовать в юникод перед заполнением
        template_txt = textfunc.toUnicode(template_txt, template_encoding)
    except:
        log.fatal(u'Ошибка преобразования текста шаблона в Unicode')
        return False

    # Генерация текста по шаблону
    gen_txt = gen(template_txt, dContext)

    file_encoding = template_encoding if output_encoding is None else output_encoding

    # Запись текста в выходной результирующий файл
    output_filename = os.path.abspath(sTxtOutputFilename)
    try:
        output_path = os.path.dirname(output_filename)
        if not os.path.exists(output_path):
            log.info(u'Создание папки <%s>' % output_path)
            os.makedirs(output_path)

        output_file = open(output_filename, 'wt+', encoding=file_encoding)
        output_file.write(gen_txt)
        output_file.close()
        # Дополнительная проверка на существующий выходной файл
        return os.path.exists(output_filename)
    except:
        if output_file:
            output_file.close()
        log.fatal(u'Ошибка генерации текстового файла <%s> по шаблону.' % output_filename)
    return False


if __name__ == '__main__':
    txt = u'Тестовые замены {{ DEBUG_MODE }}'
    print(gen(txt))
