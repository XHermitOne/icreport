#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль файла отчета.
    Файл отчета преставляет собой XML файл формата Excel xmlss.
    Документация по данному формату находится:
    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnexcl2k2/html/odc_xmlss.asp.
"""

# Подключение библиотек
import time
import copy
from xml.sax import saxutils

from ic.std.log import log
from ic.std.utils import textfunc

from ic.report import icrepgen

__version__ = (0, 0, 1, 2)

# Спецификации и структуры
# Спецификация стиля ячеек
SPC_IC_XML_STYLE = {'id': '',                                           # Идентификатор стиля
                    'align': {'align_txt': (0, 0), 'wrap_txt': False},  # Выравнивание
                    'font': None,                                       # Шрифт
                    'border': (0, 0, 0, 0),                             # Обрамление
                    'format': None,                                     # Формат ячейки
                    'color': {},                                        # Цвет
                    }


class icReportFile:
    """
    Класс файла отчета.
    """
    def __init__(self):
        """
        Конструктор класса.
        """
        pass
        
    def write(self, RepFileName_, RepData_):
        """
        Сохранить заполненный отчет в файле.
        @param RepFileName_: Имя файла отчета.
        @param RepData_: Данные отчета.
        @return: Функция возвращает имя созданного xml файла, 
            или None в случае ошибки.
        """
        pass


class icExcelXMLReportFile(icReportFile):
    """
    Файл *.XML отчета, в формат Excel XMLSS.
    """
    
    def __init__(self):
        """
        Конструктор.
        """
        icReportFile.__init__(self)
        
    def write(self, RepFileName_, RepData_):
        """
        Сохранить заполненный отчет в файле.
        @param RepFileName_: Имя файла отчета XML.
        @param RepData_: Данные отчета.
        @return: Функция возвращает имя созданного xml файла, 
            или None в случае ошибки.
        """
        xml_file = None
        try:
            # Начать запись
            xml_file = open(RepFileName_, 'w')
            xml_gen = icXMLSSGenerator(xml_file)
            xml_gen.startDocument()
            xml_gen.startBook()

            # Параметры страницы
            # xml_gen.savePageSetup(RepName_,Rep_)
        
            # Стили
            xml_gen.scanStyles(RepData_['sheet'])
            xml_gen.saveStyles()
        
            # Данные
            xml_gen.startSheet(RepData_['name'], RepData_)
            xml_gen.saveColumns(RepData_['sheet'])
            for i_row in xrange(len(RepData_['sheet'])):
                xml_gen.startRow(RepData_['sheet'][i_row])
                # Сбросить индекс ячейки
                xml_gen.cell_idx = 1
                for i_col in range(len(RepData_['sheet'][i_row])):
                    cell = RepData_['sheet'][i_row][i_col]
                    xml_gen.saveCell(i_row+1, i_col+1, cell, RepData_['sheet'])
                xml_gen.endRow()
            
            xml_gen.endSheet(RepData_)
       
            # Закончить запись
            xml_gen.endBook()
            xml_gen.endDocument()
            xml_file.close()
        
            return RepFileName_
        except:
            if xml_file:
                xml_file.close()
            log.error(u'Ошибка сохранения отчета <%s>.' % textfunc.toUnicode(RepFileName_))
            raise
        return None

    def write_book(self, RepFileName_, *RepSheetData_):
        """
        Сохранить список листов заполненного отчета в файле.
        @param RepFileName_: Имя файла отчета XML.
        @param RepSheetData_: Данные отчета, разобранные по листам.
        @return: Функция возвращает имя созданного xml файла, 
            или None в случае ошибки.
        """
        try:
            # Начать запись
            xml_file = None
            xml_file = open(RepFileName_, 'w')
            xml_gen = icXMLSSGenerator(xml_file)
            xml_gen.startDocument()
            xml_gen.startBook()
        
            for rep_sheet_data in RepSheetData_:
                # Стили
                xml_gen.scanStyles(rep_sheet_data['sheet'])
            xml_gen.saveStyles()
        
            for rep_sheet_data in RepSheetData_:
                # Данные
                xml_gen.startSheet(rep_sheet_data['name'], rep_sheet_data)
                xml_gen.saveColumns(rep_sheet_data['sheet'])
                for i_row in xrange(len(rep_sheet_data['sheet'])):
                    xml_gen.startRow(rep_sheet_data['sheet'][i_row])
                    # Сбросить индекс ячейки
                    xml_gen.cell_idx=1
                    for i_col in range(len(rep_sheet_data['sheet'][i_row])):
                        cell = rep_sheet_data['sheet'][i_row][i_col]
                        xml_gen.saveCell(i_row+1, i_col+1, cell, rep_sheet_data['sheet'])
                    xml_gen.endRow()
            
                xml_gen.endSheet(rep_sheet_data)
       
            #Закончить запись
            xml_gen.endBook()
            xml_gen.endDocument()
            xml_file.close()
        
            return RepFileName_
        except:
            if xml_file:
                xml_file.close()
            log.error(u'Ошибка сохранения отчета %s.' % RepFileName_)
            raise
        return None


class icXMLSSGenerator(saxutils.XMLGenerator):
    """
    Класс генератора конвертора отчетов в xml представление.
    """
    def __init__(self, out=None, encoding='utf-8'):
        """
        Конструктор.
        """
        saxutils.XMLGenerator.__init__(self, out, encoding)

        self._encoding = encoding
        
        # Отступ, определяющий вложение тегов
        self.break_line = ''
        
        # Стили ячеек
        self._styles = []
        
        # Текущий индекс ячейки в строке
        self.cell_idx = 0
        # Флаг установки индекса в строке
        self._idx_set = False

        # Время начала создания файла
        self.time_start = 0

    def startElementLevel(self, name, attrs):
        """
        Начало тега.
        @name: Имя тега.
        @attrs: Атрибуты тега (словарь).
        """
        # Дописать новый отступ
        self._write(unicode('\n'+self.break_line, self._encoding))

        saxutils.XMLGenerator.startElement(self, name, attrs)
        self.break_line += ' '

    def endElementLevel(self, name):
        """
        Конец тега.
        @name: Имя, закрываемого тега.
        """
        # Дописать новый отступ
        self._write(unicode('\n'+self.break_line, self._encoding))

        saxutils.XMLGenerator.endElement(self, name)

        if self.break_line:
            self.break_line = self.break_line[:-1]

    def startElement(self, name, attrs):
        """
        Начало тега.
        @name: Имя тега.
        @attrs: Атрибуты тега (словарь).
        """
        # Дописать новый отступ
        self._write(unicode('\n'+self.break_line, self._encoding))

        saxutils.XMLGenerator.startElement(self, name, attrs)

    def endElement(self, name):
        """
        Конец тега.
        @name: Имя, закрываемого тега.
        """
        saxutils.XMLGenerator.endElement(self, name)

        if self.break_line:
            self.break_line = self.break_line[:-1]

    _orientationRep2XML = {0: 'Portrait',
                           1: 'Landscape',
                           '0': 'Portrait',
                           '1': 'Landscape',
                           }

    def savePageSetup(self, Rep_):
        """
        Записать в xml файле параметры страницы.
        @param Rep_: Тело отчета.
        """
        self.startElementLevel('WorksheetOptions', {'xmlns': 'urn:schemas-microsoft-com:office:excel'})
        if 'page_setup' in Rep_:
            # Параметры страницы
            self.startElementLevel('PageSetup', {})

            # Ориентация листа
            if 'orientation' in Rep_['page_setup']:
                self.startElementLevel('Layout', 
                                       {'x:Orientation': self._orientationRep2XML[Rep_['page_setup']['orientation']],
                                        'x:StartPageNum': str(Rep_['page_setup'].setdefault('start_num', 1))})
                self.endElementLevel('Layout')

            # Поля
            if 'page_margins' in Rep_['page_setup']:
                self.startElementLevel('PageMargins', 
                                       {'x:Left': str(Rep_['page_setup']['page_margins'][0]),
                                        'x:Right': str(Rep_['page_setup']['page_margins'][1]),
                                        'x:Top': str(Rep_['page_setup']['page_margins'][2]),
                                        'x:Bottom': str(Rep_['page_setup']['page_margins'][3])})
                self.endElementLevel('PageMargins')
        
            # Обработка верхнего колонтитула
            if 'data' in Rep_['upper']:
                data = unicode(str(Rep_['upper']['data']), 'CP1251').encode('UTF-8')
                self.startElementLevel('Header', 
                                       {'x:Margin': str(Rep_['upper']['height']),
                                        'x:Data': data})
                self.endElementLevel('Header')
                
            # Обработка нижнего колонтитула
            if 'data' in Rep_['under']:
                data = unicode(str(Rep_['under']['data']), 'CP1251').encode('UTF-8')
                self.startElementLevel('Footer', 
                                       {'x:Margin': str(Rep_['under']['height']),
                                        'x:Data': data})
                self.endElementLevel('Footer')
                
            self.endElementLevel('PageSetup')

            # Параметры печати
            self.startElementLevel('Print', {})

            if 'paper_size' in Rep_['page_setup']:
                self.startElementLevel('PaperSizeIndex', {})
                self.characters(str(Rep_['page_setup']['paper_size']))
                self.endElementLevel('PaperSizeIndex')

            if 'scale' in Rep_['page_setup']:
                self.startElementLevel('Scale', {})
                self.characters(str(Rep_['page_setup']['scale']))
                self.endElementLevel('Scale')
        
            if 'resolution' in Rep_['page_setup']:
                self.startElementLevel('HorizontalResolution', {})
                self.characters(str(Rep_['page_setup']['resolution'][0]))
                self.endElementLevel('HorizontalResolution')
                self.startElementLevel('VerticalResolution', {})
                self.characters(str(Rep_['page_setup']['resolution'][1]))
                self.endElementLevel('VerticalResolution')

            if 'fit' in Rep_['page_setup']:
                self.startElementLevel('FitWidth', {})
                self.characters(str(Rep_['page_setup']['fit'][0]))
                self.endElementLevel('FitWidth')
                self.startElementLevel('FitHeight', {})
                self.characters(str(Rep_['page_setup']['fit'][1]))
                self.endElementLevel('FitHeight')
        
            self.endElementLevel('Print')

        self.endElementLevel('WorksheetOptions')

    def startBook(self):
        """
        Начало книги.
        """
        # Время начала создания фала
        self.time_start = time.time()

        # ВНИМАНИЕ! Неоходимо наличие следующих ключей
        # иначе некоторыетеги не будут пониматься библиотекой
        # и будет генерироваться ошибка чтения файла
        # <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
        # xmlns:o="urn:schemas-microsoft-com:office:office"
        # xmlns:x="urn:schemas-microsoft-com:office:excel"
        # xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">

        self.startElementLevel('Workbook', {'xmlns': 'urn:schemas-microsoft-com:office:spreadsheet',
                                            'xmlns:o': 'urn:schemas-microsoft-com:office:office',
                                            'xmlns:x': 'urn:schemas-microsoft-com:office:excel',
                                            'xmlns:ss': 'urn:schemas-microsoft-com:office:spreadsheet'})
            
    def endBook(self):
        """
        Конец книги.
        """
        self.endElementLevel('Workbook')

    def startSheet(self, RepName_, Rep_):
        """
        Теги начала страницы.
        @param RepName_: Имя отчета.
        @param Rep_: Тело отчета.
        """
        rep_name = unicode(str(RepName_), self._encoding)
        self.startElementLevel('Worksheet', {'ss:Name': rep_name})
        # Диапазон ячеек верхнего колонтитула
        try:
            if Rep_['upper']:
                refers_to = self._getUpperRangeStr(Rep_['upper'])

                self.startElementLevel('Names', {})
                self.startElementLevel('NamedRange', {'ss:Name': 'Print_Titles',
                                       'ss:RefersTo': refers_to})
                self.endElementLevel('NamedRange')
                self.endElementLevel('Names')
        except:
            log.error('Names SAVE <%s>' % Rep_['upper'])
            raise
        
        # Начало таблицы
        self.startElementLevel('Table', {})

    def _getUpperRangeStr(self, Upper_):
        """
        Представить диапазон ячеек верхнего колонтитула в виде строки.
        """
        return '=C%d:C%d,R%d:R%d' % (Upper_['col']+1, Upper_['col']+Upper_['col_size'],
                                     Upper_['row']+1, Upper_['row']+Upper_['row_size'])
        
    def endSheet(self, Rep_):
        """
        Теги начала страницы.
        @param Rep_: Тело отчета.
        """
        self.endElementLevel('Table')
        self.savePageSetup(Rep_)
        self.endElementLevel('Worksheet')
    
    def scanStyles(self, Sheet_):
        """
        Сканирование стилей на листе.
        """
        for i_row in xrange(len(Sheet_)):
            for i_col in range(len(Sheet_[i_row])):
                cell = Sheet_[i_row][i_col]
                if cell is not None:
                    self.setStyle(cell)
        return self._styles
        
    def setStyle(self, Cell_):
        """
        Определить стиль ячейки.
        @param Cell_: Атрибуты ячейки.
        @return: Возвращает индекс стиля в списке стилей.
        """
        cell_style_idx = self.getStyle(Cell_)
        if cell_style_idx is None:
            # Создать новый стиль
            new_idx = len(self._styles)
            cell_style = copy.deepcopy(SPC_IC_XML_STYLE)
            cell_style['align'] = Cell_['align']
            cell_style['font'] = Cell_['font']
            cell_style['border'] = Cell_['border']
            cell_style['format'] = Cell_['format']
            cell_style['color'] = Cell_['color']
            cell_style['id'] = 'x'+str(new_idx)
            # Прописать в ячейке идентификатор стиля
            Cell_['style_id'] = cell_style['id']
            self._styles.append(cell_style)
            return new_idx
        return cell_style_idx
      
    def getStyle(self, Cell_):
        """
        Определить стиль ячейки из уже имеющихся.
        @param Cell_: Атрибуты ячейки.
        @return: Возвращает индекс стиля в списке стилей.
        """
        # сначала поискать в списке стилей
        find_style = [style for style in self._styles if self._equalStyles(style, Cell_)]

        # Если такой стиль найден, то вернуть его
        if find_style:
            Cell_['style_id'] = find_style[0]['id']
            return self._styles.index(find_style[0])
        return None
        
    def _equalStyles(self, Style1_, Style2_):
        """
        Функция проверки равенства стилей.
        """
        return bool(self._equalAlign(Style1_['align'], Style2_['align']) and
                    self._equalFont(Style1_['font'], Style2_['font']) and
                    self._equalBorder(Style1_['border'], Style2_['border']) and
                    self._equalFormat(Style1_['format'], Style2_['format']) and
                    self._equalColor(Style1_['color'], Style2_['color']))

    def _equalAlign(self, Align1_, Align2_):
        """
        Равенство выравниваний.
        """
        return bool(Align1_ == Align2_)
        
    def _equalFont(self, Font1_, Font2_):
        """
        Равенство шрифтов.
        """
        return bool(Font1_ == Font2_)
        
    def _equalBorder(self, Border1_, Border2_):
        """
        Равенство обрамлений.
        """
        return bool(Border1_ == Border2_)
        
    def _equalFormat(self, Fmt1_, Fmt2_):
        """
        Равенство форматов.
        """
        return bool(Fmt1_ == Fmt2_)
        
    def _equalColor(self, Color1_, Color2_):
        """
        Равенство цветов.
        """
        return bool(Color1_ == Color2_)
        
    def saveStyles(self):
        """
        Записать стили.
        """
        self.startElementLevel('Styles', {})
        
        # Стиль по умолчанию
        self.startElementLevel('Style', {'ss:ID': 'Default', 'ss:Name': 'Normal'})
        self.startElement('Alignment', {'ss:Vertical': 'Bottom'})
        self.endElement('Alignment')    
        self.startElement('Borders', {})
        self.endElement('Borders')    
        self.startElement('Font', {'ss:FontName': 'Arial Cyr'})
        self.endElement('Font')    
        self.startElement('Interior', {})
        self.endElement('Interior')    
        self.startElement('NumberFormat', {})
        self.endElement('NumberFormat')    
        self.startElement('Protection', {})
        self.endElement('Protection')    
        self.endElementLevel('Style')
        
        # Дополнительные стили
        for style in self._styles:
            self.startElementLevel('Style', {'ss:ID': style['id']})
            # Выравнивание
            align = {}
            h_align = self._alignRep2XML[style['align']['align_txt'][icrepgen.IC_REP_ALIGN_HORIZ]]
            v_align = self._alignRep2XML[style['align']['align_txt'][icrepgen.IC_REP_ALIGN_VERT]]
            if h_align:
                align['ss:Horizontal'] = h_align
            if v_align:
                align['ss:Vertical'] = v_align
            # Перенос по словам
            if style['align']['wrap_txt']:
                align['ss:WrapText'] = '1'
            
            self.startElement('Alignment', align)
            self.endElement('Alignment')
            
            # Обрамление
            self.startElementLevel('Borders', {})
            for border_pos in range(4):
                border = self._borderRep2XML(style['border'], border_pos)
                if border:
                    border_element = dict([('ss:'+stl_name, border[stl_name]) for stl_name in
                                          [stl_name for stl_name in border.keys() if border[stl_name] is not None]])
                    self.startElement('Border', border_element)
                    self.endElement('Border')
            self.endElementLevel('Borders')
               
            # Шрифт
            font = {}
            font['ss:FontName'] = style['font']['name']
            font['ss:Size'] = str(int(style['font']['size']))
            if style['font']['style'] == 'bold' or style['font']['style'] == 'boldItalic':
                font['ss:Bold'] = '1'
            elif style['font']['style'] == 'italic' or style['font']['style'] == 'boldItalic':
                font['ss:Italic'] = '1'
            if style['color']:
                if 'text' in style['color'] and style['color']['text']:
                    font['ss:Color'] = self._getRGBColor(style['color']['text'])
                
            self.startElement('Font', font)
            self.endElement('Font')
            
            # Интерьер
            interior = {}
            if style['color']:
                if 'background' in style['color'] and style['color']['background']:
                    interior['ss:Color'] = self._getRGBColor(style['color']['background'])
                    interior['ss:Pattern'] = 'Solid'

            self.startElement('Interior', interior)
            self.endElement('Interior')

            # Формат
            fmt = {}
            if style['format']:
                fmt['ss:Format'] = self._getNumFmt(style['format'])
                self.startElement('NumberFormat', fmt)
                self.endElement('NumberFormat')

            self.endElementLevel('Style')
            
        self.endElementLevel('Styles')
        
    def _getRGBColor(self, Color_):
        """
        Преобразование цвета из (R,G,B) в #RRGGBB.
        """
        if type(Color_) in (list, tuple):
            return '#%02X%02X%02X' % (Color_[0], Color_[1], Color_[2])
        # ВНИМАНИЕ! Если цвет задается не RGB форматом, тогда оставить его без изменения
        return Color_

    def _getNumFmt(self, Fmt_):
        """
        Формат чисел.
        """
        if Fmt_[0] == icrepgen.REP_FMT_EXCEL:
            return Fmt_[1:]
        elif Fmt_[0] == icrepgen.REP_FMT_STR:
            return '@'
        elif Fmt_[0] == icrepgen.REP_FMT_NUM:
            return '0'
        elif Fmt_[0] == icrepgen.REP_FMT_FLOAT:
            return '0.'
        return '0'

    # Преобразование выравнивания из нашего представления
    # в xml представление.
    _alignRep2XML = {icrepgen.IC_HORIZ_ALIGN_LEFT: 'Left',
                     icrepgen.IC_HORIZ_ALIGN_CENTRE: 'Center',
                     icrepgen.IC_HORIZ_ALIGN_RIGHT: 'Right',
                     icrepgen.IC_VERT_ALIGN_TOP: 'Top',
                     icrepgen.IC_VERT_ALIGN_CENTRE: 'Center',
                     icrepgen.IC_VERT_ALIGN_BOTTOM: 'Bottom',
                     }
        
    def _borderRep2XML(self, Border_, Position_):
        """
        Преобразование обрамления из нашего представления
            в xml представление.
        """
        if Border_[Position_]:
            return {'Position': self._positionRep2XML.setdefault(Position_, 'Left'),
                    'Color': self._colorRep2XML(Border_[Position_].setdefault('color', None)),
                    'LineStyle': self._lineRep2XML.setdefault(Border_[Position_].setdefault('style', icrepgen.IC_REP_LINE_TRANSPARENT), 'Continuous'),
                    'Weight': str(Border_[Position_].setdefault('weight', 1)),
                    }

    # Преобразование позиции линии обрамления из нашего представления
    # в xml представление.
    _positionRep2XML = {icrepgen.IC_REP_BORDER_LEFT: 'Left',
                        icrepgen.IC_REP_BORDER_RIGHT: 'Right',
                        icrepgen.IC_REP_BORDER_TOP: 'Top',
                        icrepgen.IC_REP_BORDER_BOTTOM: 'Bottom',
                        }
        
    _lineRep2XML = {icrepgen.IC_REP_LINE_SOLID: 'Continuous',
                    icrepgen.IC_REP_LINE_SHORT_DASH: 'Dash',
                    icrepgen.IC_REP_LINE_DOT_DASH: 'DashDot',
                    icrepgen.IC_REP_LINE_DOT: 'Dot',
                    icrepgen.IC_REP_LINE_TRANSPARENT: None,
                    }
    
    def _colorRep2XML(self, Color_):
        """
        Преобразование цвета из нашего представления
            в xml представление.
        """
        return None

    def saveColumns(self, Sheet_):
        """
        Запись атрибутов колонок.
        """
        width_cols = self.getWidthColumns(Sheet_)
        for width_col in width_cols:
            # Если ширина колонки определена
            if width_col is not None:
                self.startElement('Column', {'ss:Width': str(width_col), 'ss:AutoFitWidth': '0'})
            else:
                self.startElement('Column', {'ss:AutoFitWidth': '0'})
            self.endElement('Column')
            
    def getColumnCount(self, Sheet_):
        """
        Определить количество колонок.
        """
        if Sheet_:
            return max([len(row) for row in Sheet_])
        return 0

    def getWidthColumns(self, Sheet_):
        """
        Ширины колонок.
        """
        col_count = self.getColumnCount(Sheet_)
        col_width = []
        # Выбрать строку по которой будыт выставяться ширины колонок
        row = [row for row in Sheet_ if len(row) == col_count][0] if Sheet_ else list()
        for cell in row:
            if cell:
                # log.debug('Column width <%s>' % cell['width'])
                col_width.append(cell['width'])
            else:
                # log.debug('Column default width')
                col_width.append(8.43)
        return col_width

    def getRowHeight(self, Row_):
        """
        Высота строки.
        """
        return min([cell['height'] for cell in [cell_ for cell_ in Row_ if type(cell_) == dict and 'height' in cell_]])
            
    def startRow(self, Row_):
        """
        Начало строки.
        """
        height_row = self.getRowHeight(Row_)
        self.startElementLevel('Row', {'ss:Height': str(height_row)})
        self._idx_set = False       # Сбросить флаг установки индекса
        self.cell_idx = 1
            
    def endRow(self):
        """
        Конец строки.
        """
        self.endElementLevel('Row')
        
    def _saveCellStyleID(self, Cell_):
        """
        Определить идентификатор стиля ячейки для записи.
        """
        if 'style_id' in Cell_:
            return Cell_['style_id']
        else:
            style_idx = self.getStyle(Cell_)
            if style_idx is not None:
                style_id = self._styles[style_idx]['id']
            else:
                style_id = 'Default'
            return style_id
        
    def saveCell(self, Row_, Col_, Cell_, Sheet_=None):
        """
        Записать ячейку.
        @param Row_: НОмер строки.
        @param Col_: Номер колонки.
        @param Cell_: Атрибуты ячейки.
        """
        if Cell_ is None:
            self._idx_set = False   # Сбросить флаг установки индекса
            self.cell_idx += 1
            return 

        if 'hidden' in Cell_ and Cell_['hidden']:
            self._idx_set = False   # Сбросить флаг установки индекса
            # ВНИМАНИЕ!!! Здесь надо увеличивать индекс на 1
            # потому что в Excel индексирование начинается с 1 !!!
            self.cell_idx += 1
            return 

        cell_attr = {}
        if self.cell_idx > 1:
            if not self._idx_set:
                cell_attr = {'ss:Index': str(self.cell_idx)}
                self._idx_set = True    # Установить флаг установки индекса

        # Объединение ячеек
        if Cell_['merge_col'] > 1:
            cell_attr['ss:MergeAcross'] = str(Cell_['merge_col']-1)
            # Обработать верхнюю строку области объединения
            self._setCellMergeAcross(Row_, Col_, Cell_['merge_col'], Sheet_)
            if Cell_['merge_row'] > 1:
                # Обработать дополнительную область объединения
                self._setCellMerge(Row_, Col_, Cell_['merge_col'], Cell_['merge_row'], Sheet_)
            self._idx_set = False   # Сбросить флаг установки индекса
        # ВНИМАНИЕ!!! Здесь надо увеличивать индекс на 1
        # потому что в Excel индексирование начинается с 1 !!!
        self.cell_idx = Col_+1
    
        if Cell_['merge_row'] > 1:
            cell_attr['ss:MergeDown'] = str(Cell_['merge_row']-1)
            # Обработать левый столбец области объединения
            self._setCellMergeDown(Row_, Col_, Cell_['merge_row'], Sheet_)

        # Стиль
        cell_attr['ss:StyleID'] = self._saveCellStyleID(Cell_)

        self.startElement('Cell', cell_attr)
        if Cell_['value'] is not None:
            self.startElement('Data', {'ss:Type': self._getCellType(Cell_['value'])})
            value = self._getCellValue(Cell_['value'])
            self.characters(value)
        
            self.endElement('Data')

        self.endElement('Cell')
        
    def _getCellValue(self, Value_):
        """
        Подготовить значение для записи в файл.
        """
        if self._getCellType(Value_) == 'Number':
            # Это число
            value = Value_.strip()
        else:
            # Это не число
            value = Value_

        if type(value) != type(u''):
            try:
                value = unicode(str(value), self._encoding)
            except:
                value = unicode(str(value), 'cp1251')
        if value:
            value = saxutils.escape(value)
            
        return value

    def _getCellType(self, Value_):
        """
        Тип ячейки.
        """
        try:
            # Это число
            float(Value_)
            return 'Number'
        except:
            # Это не число
            return 'String'

    def _setCellMergeAcross(self, Row_, Col_, MergeAcross_, Sheet_):
        """
        Сбросить все ячейки, которые попадают в горизонтальную зону объединения.
        @param Row_: НОмер строки.
        @param Col_: Номер колонки.
        @param MergeAcross_: Количество ячеек, объединенных с текущей.
        @param Sheet_: Структура листа.
        """
        for i in range(1, MergeAcross_):
            try:
                cell = Sheet_[Row_-1][Col_+i-1]
            except IndexError:
                continue
            if cell and (not cell['value']):
                Sheet_[Row_-1][Col_+i-1]['hidden'] = True
        return Sheet_

    def _setCellMergeDown(self, Row_, Col_, MergeDown_, Sheet_):
        """
        Сбросить все ячейки, которые попадают в вертикальную зону объединения.
        @param Row_: НОмер строки.
        @param Col_: Номер колонки.
        @param MergeDown_: Количество ячеек, объединенных с текущей.
        @param Sheet_: Структура листа.
        """
        for i in range(1, MergeDown_):
            try:
                cell = Sheet_[Row_+i-1][Col_-1]
            except IndexError:
                continue
            if cell and (not cell['value']):
                Sheet_[Row_+i-1][Col_-1]['hidden'] = True
        return Sheet_

    def _setCellMerge(self, Row_, Col_, MergeAcross_, MergeDown_, Sheet_):
        """
        Сбросить все ячейки, которые попадают в зону объединения.
        @param Row_: НОмер строки.
        @param Col_: Номер колонки.
        @param MergeAcross_: Количество ячеек, объединенных с текущей.
        @param MergeDown_: Количество ячеек, объединенных с текущей.
        @param Sheet_: Структура листа.
        """
        for x in range(1, MergeAcross_):
            for y in range(1, MergeDown_):
                try:
                    cell = Sheet_[Row_+y-1][Col_+x-1]
                except IndexError:
                    continue
                if cell is not None and (not cell['value']):
                    Sheet_[Row_+y-1][Col_+x-1]['hidden'] = True
        return Sheet_
