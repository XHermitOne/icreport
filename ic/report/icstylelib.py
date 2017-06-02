#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль библиотеки стилей для отчетов.
"""

import copy
import string

from ic.std.log import log
from ic.std.convert import xml2dict

from ic.report import icreptemplate

__version__ = (0, 0, 1, 1)

# Стиль
IC_REP_STYLE = {'font': None,   # Шрифт Структура типа ic.components.icfont.SPC_IC_FONT
                'color': None,  # Цвет
                # 'border':None, #Обрамление
                # 'align':None, #Расположение текста
                # 'format': None, #Формат ячейки
                }


class icRepStyleLib(icreptemplate.icReportTemplate):
    """
    Класс библиотеки стилей для отчетов.
    """
    def __init__(self):
        """
        Конструктор класса.
        """
        icreptemplate.icReportTemplate.__init__(self)
        self._style_lib = {}
        
    def convert(self, SrcData_):
        """
        Конвертация из XML представления библиотеки стилей в представление
            библиотеки стилей отчета.
        @param SrcData_: Исходные данные.
        """
        pass
        
    def get(self):
        """
        Получить сконвертированные данные библиотеки стилей отчета.
        """
        return self._style_lib


class icXMLRepStyleLib(icRepStyleLib):
    """
    Класс преобразования XML библиотеки стилей для отчетов.
    """
    def __init__(self):
        """
        Конструктор.
        """
        icRepStyleLib.__init__(self)

    def open(self, XMLFileName_):
        """
        Открыть XML файл.
        @param XMLFileName_: Имя XML файла библиотеки стилей отчета.
        """
        return xml2dict.XmlFile2Dict(XMLFileName_)

    def convert(self, XMLFileName_):
        """
        Конвертация из XML файла в представление библиотеки стилей отчета.
        @param XMLFileName_: Исходные данные.
        """
        xml_data = self.open(XMLFileName_)
        self._style_lib = self.data_convert(xml_data)
        return self._style_lib

    def data_convert(self, Data_):
        """
        Преобразование данных из одного представления в другое.
        """
        try:
            style_lib = {}
            
            # Определение основных структур
            workbook = Data_['children'][0]
            # Стили (в виде словаря)
            styles = {}
            styles_lst = [element for element in workbook['children'] if element['name'] == 'Styles']
            if styles_lst:
                styles = dict([(style['ID'], style) for style in  styles_lst[0]['children']])

            worksheets = [element for element in workbook['children'] if element['name'] == 'Worksheet']

            rep_worksheet = worksheets[0]
            rep_data_tab = rep_worksheet['children'][0]
            # Список строк
            rep_data_rows = [element for element in rep_data_tab['children'] if element['name'] == 'Row']

            # Заполнение
            style_lib = self._getStyles(rep_data_rows, styles)
            return style_lib
        except:
            # Вывести сообщение об ошибке в лог
            log.error(u'Ошибка преобразования данных библиотеки стилей отчета.')
            return None

    # Разрешенные символы в именах тегов стилей
    LATIN_CHARSET = string.ascii_uppercase+string.ascii_lowercase+string.digits+'_'

    def _isStyleTag(self, Value_):
        """
        Определить является ли значение тегом стиля.
        @param Value_: Строка-значение в ячейке.
        @return: True/False.
        """
        for symbol in Value_:
            if symbol not in self.LATIN_CHARSET:
                return False
        return True

    def _getStyles(self, Rows_, Styles_):
        """
        Определить словарь стилей.
        @param Rows_: Список строк в XML.
        @param Styles_: Словарь стилей в XML.
        """
        styles = {}
        for row in Rows_:
            for cell in row['children']:
                # Получить значение
                value = cell['children'][0]['value']
                # Обрабатывать только строковые значения
                if value and isinstance(value, str):
                    if self._isStyleTag(value):
                        style_id = cell['StyleID']
                        styles[value] = self._getStyle(style_id, Styles_)
        return styles
                            
    def _getStyle(self, StyleID_, Styles_):
        """
        Определить стиль по его идентификатору.
        @param StyleID_: Идентификатор стиля в XML описании.
        @param Styles_: Словарь стилей в XML.
        """
        style = copy.deepcopy(IC_REP_STYLE)
        
        if StyleID_ in Styles_:
            style['font'] = self._getFontStyle(Styles_[StyleID_])
            style['color'] = self._getColorStyle(Styles_[StyleID_])
            # style['border']=self._getBordersStyle(Styles_[StyleID_])
            # style['align']=self._getAlignStyle(Styles_[StyleID_])
            # style['format']=self._getFmtStyle(Styles_[StyleID_])
        return style
