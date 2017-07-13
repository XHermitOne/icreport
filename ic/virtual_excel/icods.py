# !/usr/bin/env python
# -*- coding: utf-8 -*-

import os.path
import re

import config

try:
    # Если Virtual Excel работает в окружении icReport
    from ic.std.log import log
except ImportError:
    # Если Virtual Excel работает в окружении icServices
    from services.ic_std.log import log

log.init(config)
    
try:
    import odf.opendocument
    import odf.style
    import odf.number
    import odf.text
    import odf.table
except ImportError:
    log.error(u'ODFpy Import Error.')

__version__ = (0, 0, 2, 4)

DIMENSION_CORRECT = 35
DEFAULT_STYLE_ID = 'Default'

CM2PT_CORRECT = 25

SPREADSHEETML_CR = '&#10;'

LIMIT_ROWS_REPEATED = 1000000
LIMIT_COLUMNS_REPEATED = 100

ODS_LANDSCAPE_ORIENTATION = 'landscape'
ODS_PORTRAIT_ORIENTATION = 'portrait'
LANDSCAPE_ORIENTATION = ODS_LANDSCAPE_ORIENTATION.title()
PORTRAIT_ORIENTATION = ODS_PORTRAIT_ORIENTATION.title()

A4_PAPER_FORMAT = 'A4'
A3_PAPER_FORMAT = 'A3'

DEFAULT_ENCODE = 'utf-8'


class icODS(object):
    """
    Класс конвертации представления VirtualExcel в ODS файл.
    """
    def __init__(self):
        """
        Конструктор.
        """
        self._styles_ = {}
        self.ods_document = None

        # Внутренние данные в xmlss представлении
        self.xmlss_data = None
        
        # Стили числовых форматов в виде словаря
        self._number_styles_ = {}

        # Индекс генерации имен стилей
        self._style_name_idx = 0
        
    def Save(self, sFileName, dData=None):
        """
        Сохранить в ODS файл.
        @param sFileName: Имя ODS файла.
        @param dData: Словарь данных.
        @return: True/False.
        """
        if dData is None:
            log.warning('ODS. Not define saved data')

        self.ods_document = None
        self._styles_ = {}
        
        workbooks = dData.get('children', None)
        if not workbooks:
            workbook = {}
        else:
            workbook = workbooks[0]
        
        self.setWorkbook(workbook)
        
        if self.ods_document:
            if isinstance(sFileName, str):
                # ВНИМАНИЕ! Перед сохранением надо имя файла сделать
                # Юникодной иначе падает по ошибке в функции save
                sFileName = unicode(sFileName, DEFAULT_ENCODE)
            # Добавлять автоматически расширение-+
            # к имени файла (True - да)          |
            #                                    V
            self.ods_document.save(sFileName, addsuffix=False)
            self.ods_document = None
            
        return True
    
    def getChildrenByName(self, dData, sName):
        """
        Дочерние элементы по имени.
        @param dData: Словарь данных.
        @param sName: Имя дочернего элемента.
        """
        return [item for item in dData.get('children', []) if item['name'] == sName]
        
    def setWorkbook(self, dData):
        """
        Заполнить книгу.
        @param dData: Словарь данных.
        """
        self.ods_document = odf.opendocument.OpenDocumentSpreadsheet()
        
        if dData:
            styles = self.getChildrenByName(dData, 'Styles')
            if styles:
                self.setStyles(styles[0])

            sheets = self.getChildrenByName(dData, 'Worksheet')
            if sheets:
                for sheet in sheets:
                    ods_table = self.setWorksheet(sheet)
                    self.ods_document.spreadsheet.addElement(ods_table)

    def setStyles(self, dData):
        """
        Заполнить стили.
        @param dData: Словарь данных.
        """
        log.info('styles: <%s>' % dData)
        
        styles = dData.get('children', [])
        for style in styles:
            ods_style = self.setStyle(style)
            self.ods_document.automaticstyles.addElement(ods_style)

    def setFont(self, dData):
        """
        Заполнить шрифт стиля.
        @param dData: Словарь данных.
        """
        log.debug('Set FONT <%s>' % dData)
        font = {}
        font_name = dData.get('FontName', 'Arial')
        font_size = dData.get('Size', '10')
        font_bold = 'bold' if dData.get('Bold', '0') in (True, '1') else None
        font_italic = 'italic' if dData.get('Italic', '0') in (True, '1') else None

        through = 'solid' if dData.get('StrikeThrough', '0') in (True, '1') else None
        underline = 'solid' if dData.get('Underline', 'None') != 'None' else None

        font['fontfamily'] = font_name
        font['fontsize'] = font_size
        font['fontweight'] = font_bold
        font['fontstyle'] = font_italic

        if through:
            font['textlinethroughstyle'] = 'solid'
            font['textlinethroughtype'] = 'single'
        if underline:
            font['textunderlinestyle'] = 'solid'
            font['textunderlinewidth'] = 'auto'

        return font

    def _genNumberStyleName(self):
        """
        Генерация имени стиля формата числового представления.
        """
        # from services.ic_std.utils import uuid
        # return uuid.get_uuid()
        self._style_name_idx += 1
        return 's%d' % self._style_name_idx

    def setNumberFormat(self, dData):
        """
        Заполнить формат числового представления.
        @param dData: Словарь данных.
        """
        number_format = {}
        format = dData.get('Format', '0')

        # Не анализировать знак %
        format = format.replace('%', '')
        decimalplaces = len(format[format.find(',')+1:]) if format.find(',') >= 0 else 0
        minintegerdigits = len([i for i in list(format[:format.find(',')]) if i == '0']) if format.find(',') >= 0 else len([i for i in list(format) if i == '0'])
        grouping = 'true' if format.find(' ') >= 0 else 'false'
                    
        number_format['decimalplaces'] = str(decimalplaces)
        number_format['minintegerdigits'] = str(minintegerdigits)
        number_format['grouping'] = str(grouping)
        
        log.debug('set number format <%s>' % number_format)
        return number_format

    _lineStyles_SpreadsheetML2ODS = {None: 'solid',
                                     'Continuous': 'solid',
                                     'Double': 'double',
                                     'Dot': 'dotted',
                                     'Dash': 'dashed',
                                     'DashDot': 'dotted',
                                     'DashDotDot': 'dashed',
                                     'solid': 'solid'}
    _lineStyles_ODS2SpreadsheetML = {None: 'Continuous',
                                     'solid': 'Continuous',
                                     'double': 'Double',
                                     'dotted': 'Dot',
                                     'dashed': 'Dash'}

    def setBorders(self, dData):
        """
        Заполнить бордеры.
        @param dData: Словарь данных.
        """
        borders = {}
        for border in dData['children']:
            if border:
                border_weight = border.get('Weight', '1')
                color = border.get('Color', '#000000')
                line_style = self._lineStyles_SpreadsheetML2ODS.get(border.get('LineStyle', 'solid'), 'solid')
                if border['Position'] == 'Left':
                    borders['borderleft'] = '%spt %s %s' % (border_weight, line_style, color)
                elif border['Position'] == 'Right':
                    borders['borderright'] = '%spt %s %s' % (border_weight, line_style, color)
                elif border['Position'] == 'Top':
                    borders['bordertop'] = '%spt %s %s' % (border_weight, line_style, color)
                elif border['Position'] == 'Bottom':
                    borders['borderbottom'] = '%spt %s %s' % (border_weight, line_style, color)
                
        return borders
        
    _alignHorizStyle_SpreadsheetML2ODS = {'Left': 'start',
                                          'Right': 'end',
                                          'Center': 'center',
                                          'Justify': 'justify',
                                          }
    _alignVertStyle_SpreadsheetML2ODS = {'Top': 'top',
                                         'Bottom': 'bottom',
                                         'Center': 'middle',
                                         'Justify': 'justify',
                                         }

    def setAlignmentParagraph(self, dData):
        """
        Заполнить выравнивания текста стиля.
        @param dData: Словарь данных.
        """
        align = {}
        horiz = dData.get('Horizontal', None)
        vert = dData.get('Vertical', None)

        if horiz:
            align['textalign'] = self._alignHorizStyle_SpreadsheetML2ODS.get(horiz, 'start')
            
        if vert:
            align['verticalalign'] = self._alignVertStyle_SpreadsheetML2ODS.get(vert, 'top')

        return align

    def setAlignmentCell(self, dData):
        """
        Заполнить выравнивания текста стиля.
        @param dData: Словарь данных.
        """
        align = {}
        wrap_txt = dData.get('WrapText', 0)
        shrink_to_fit = dData.get('ShrinkToFit', 0)
        vert = dData.get('Vertical', None)
        
        if vert:
            align['verticalalign'] = self._alignVertStyle_SpreadsheetML2ODS.get(vert, 'top')
        
        if wrap_txt:
            align['wrapoption'] = 'wrap'
            
        if shrink_to_fit:
            align['shrinktofit'] = 'true'
        
        return align

    def setInteriorCell(self, dData):
        """
        Заполнить интерьер ячейки стиля.
        @param dData: Словарь данных.
        """
        interior = {}
        color = dData.get('Color', None)

        if color:
            interior['backgroundcolor'] = color

        return interior

    def setStyle(self, dData):
        """
        Заполнить стиль.
        @param dData: Словарь данных.
        """
        log.info('Set STYLE: <%s>' % dData)

        properties_args = {}
        number_format = self.getChildrenByName(dData, 'NumberFormat')
        if number_format:
            # Заполнениние формата числового представления
            number_properties = self.setNumberFormat(number_format[0])
            number_style_name = self._genNumberStyleName()
            properties_args['datastylename'] = number_style_name
            
            format = number_format[0].get('Format', '0')
            log.debug('Set NUMBER FORMAT <%s>' % format)
            if '%' in format:
                ods_number_style = odf.number.PercentageStyle(name=number_style_name)
                ods_number_style.addElement(odf.number.Number(**number_properties))
                ods_number_style.addElement(odf.number.Text(text='%'))
                self.ods_document.styles.addElement(ods_number_style)
            else:
                ods_number_style = odf.number.NumberStyle(name=number_style_name)
                ods_number_style.addElement(odf.number.Number(**number_properties))
                self.ods_document.automaticstyles.addElement(ods_number_style)
    
        style_id = dData['ID']
        properties_args['name'] = style_id
        properties_args['family'] = 'table-cell'
        ods_style = odf.style.Style(**properties_args)

        properties_args = {}
        fonts = self.getChildrenByName(dData, 'Font')
        # log.warning('Set font <%s>' % fonts)
        if fonts:
            # Заполнениние шрифта
            properties_args = self.setFont(fonts[0])
            log.debug('Font args: <%s>' % properties_args)

        if properties_args:
            ods_properties = odf.style.TextProperties(**properties_args)
            ods_style.addElement(ods_properties)
        
        properties_args = {}
        borders = self.getChildrenByName(dData, 'Borders')
        if borders:
            # Заполнение бордеров
            args = self.setBorders(borders[0])
            properties_args.update(args)
            log.debug('Border args: <%s>' % args)

        alignments = self.getChildrenByName(dData, 'Alignment')
        if alignments:
            # Заполнение выравнивания текста
            args = self.setAlignmentCell(alignments[0])
            properties_args.update(args)
            log.debug('Alignment Cell args: %s' % args)

        interiors = self.getChildrenByName(dData, 'Interior')
        if interiors:
            # Заполнение интерьера
            args = self.setInteriorCell(interiors[0])
            properties_args.update(args)
            log.debug('Interior Cell args: %s' % args)

        if properties_args:
            ods_properties = odf.style.TableCellProperties(**properties_args)
            ods_style.addElement(ods_properties)            
            
        properties_args = {}
        if alignments:
            # Заполнение выравнивания текста
            args = self.setAlignmentParagraph(alignments[0])
            properties_args.update(args)
            log.debug('Alignment Paragraph args: <%s>' % args)

        if properties_args:
            ods_properties = odf.style.ParagraphProperties(**properties_args)
            ods_style.addElement(ods_properties)            

        # Зарегистрировать стиль в кеше по имени
        self._styles_[style_id] = ods_style
        return ods_style

    def setWorksheet(self, dData):
        """
        Заполнить лист.
        @param dData: Словарь данных.
        """
        log.info('worksheet: <%s>' % dData)

        log.debug('setWorksheet: <%s:%s>' % (type(dData.get('Name', None)), dData.get('Name', None)))
        sheet_name = dData.get('Name', 'Лист')
        if type(sheet_name) != unicode:
            sheet_name = unicode(sheet_name, DEFAULT_ENCODE)
        ods_table = odf.table.Table(name=sheet_name)
        tables = self.getChildrenByName(dData, 'Table')
        if tables:
            self.setTable(tables[0], ods_table)
            
        # Установка параметров страницы
        worksheet_options = self.getChildrenByName(dData, 'WorksheetOptions')
        if worksheet_options:
            self.setWorksheetOptions(worksheet_options[0])

        # Установка разрывов страницы
        page_breaks = self.getChildrenByName(dData, 'PageBreaks')
        if page_breaks:
            self.setPageBreaks(page_breaks[0], ods_table)
        return ods_table

    def _set_row_break(self, iRow, ODSTable):
        """
        Установить разрыв по строке.
        @param iRow: Номер строки.
        @param ODSTable: Объект ODS таблицы.
        """
        if ODSTable:
            rows = ODSTable.getElementsByType(odf.table.TableRow)
            if rows:
                style_name = rows[iRow].getAttribute('stylename')
                style = self._styles_[style_name]
                if style:
                    row_properties = style.getElementsByType(odf.style.TableRowProperties)
                    if row_properties:
                        row_properties[0].setAttribute('breakbefore', 'page')

    def setPageBreaks(self, dData, ODSTable):
        """
        Установить разрывы страниц.
        @param dData: Словарь данных.
        @param ODSTable: Объект ODS таблицы.
        """
        row_breaks = dData['children'][0]['children']
        for row_break in row_breaks:
            i_row = row_break['children'][0]['value']
            log.debug('Row break: <%s>' % i_row)
            self._set_row_break(i_row, ODSTable)

    def setWorksheetOptions(self, dData):
        """
        Установить параметры страницы.
        @param dData: Словарь данных.
        """
        log.debug('WorksheetOptions: <%s>' % dData)
        page_setup = self.getChildrenByName(dData, 'PageSetup')
        print_setup = self.getChildrenByName(dData, 'Print')
        fit_to_page = self.getChildrenByName(dData, 'FitToPage')
        ods_properties = {'writingmode': 'lr-tb'}
        orientation = None
        if page_setup:
            
            layout = self.getChildrenByName(page_setup[0], 'Layout')
            orientation = layout[0].get('Orientation', None) if layout else None
            if orientation:
                if type(orientation) == unicode:
                    orientation = orientation.encode(DEFAULT_ENCODE)
                ods_properties['printorientation'] = orientation.lower()
                
            page_margins = self.getChildrenByName(page_setup[0], 'PageMargins')
            margin_top = page_margins[0].get('Top', None) if page_margins else None
            margin_bottom = page_margins[0].get('Bottom', None) if page_margins else None
            margin_left = page_margins[0].get('Left', None) if page_margins else None
            margin_right = page_margins[0].get('Right', None) if page_margins else None
            if margin_top:
                ods_properties['margintop'] = str(float(margin_top)*DIMENSION_CORRECT) 
            if margin_bottom:
                ods_properties['marginbottom'] = str(float(margin_bottom)*DIMENSION_CORRECT) 
            if margin_left:
                ods_properties['marginleft'] = str(float(margin_left)*DIMENSION_CORRECT) 
            if margin_right:
                ods_properties['marginright'] = str(float(margin_right)*DIMENSION_CORRECT)                 
        else:
            log.warning('WorksheetOptions PageSetup not define')
        if print_setup:
            paper_size_idx = self.getChildrenByName(print_setup[0], 'PaperSizeIndex')
            if paper_size_idx:
                width, height = self._getPageSizeByExcelIndex(paper_size_idx[0]['value'])
                if orientation == LANDSCAPE_ORIENTATION:
                    variable = height
                    height = width
                    width = variable
                # Преобразовать к строковому типу
                ods_properties['pagewidth'] = '%scm' % str(width)
                ods_properties['pageheight'] = '%scm' % str(height)
        else:
            log.warning('WorksheetOptions Print not define')

        if fit_to_page:
            ods_properties['scaletopages'] = '1'

        ods_pagelayout = odf.style.PageLayout(name='MyPageLayout')
        log.debug('[ODS] Page Layout Properties <%s>' % ods_properties)
        ods_pagelayoutproperties = odf.style.PageLayoutProperties(**ods_properties)
        ods_pagelayout.addElement(ods_pagelayoutproperties)
        if self.ods_document:
            self.ods_document.automaticstyles.addElement(ods_pagelayout)
                
            masterpage = odf.style.MasterPage(name=DEFAULT_STYLE_ID, pagelayoutname=ods_pagelayout)
            self.ods_document.masterstyles.addElement(masterpage)
        else:
            log.warning('Not define ODS document!')
        return ods_pagelayout

    def _getPageSizeByExcelIndex(self, iPaperSizeIndex):
        """
        Получить размер листа по его индуксу в Excel.
        @param iPaperSizeIndex: Индекс 9-A4 8-A3.
        @return: Кортеж (Ширина в см, Высота в см).
        """
        if type(iPaperSizeIndex) != int:
            iPaperSizeIndex = int(iPaperSizeIndex)
            
        if iPaperSizeIndex == 9:
            # A4
            return 21.0, 29.7
        elif iPaperSizeIndex == 8:
            # A3
            return 42.0, 29.7
        else:
            # По умолчанию A4
            return 21.0, 29.7

    def setTable(self, dData, ODSTable):
        """
        Заполнить таблицу.
        @param dData: Словарь данных.
        @param ODSTable: Объект таблицы ODS файла.
        """
        log.info('table: <%s>' % dData)

        # колонки
        i = 1
        columns = self.getChildrenByName(dData, 'Column')
        for column in columns:
            # Учет индекса колонки
            idx = int(column.get('Index', i))
            if idx > i:
                for ii in range(idx-i):
                    ods_column = odf.table.TableColumn()
                    ODSTable.addElement(ods_column)
                i = idx+1
            else:
                i += 1
            ods_column = self.setColumn(column)
            ODSTable.addElement(ods_column)
            
            span = column.get('Span', None)
            if span:
                i += int(span)
        
        # строки
        i = 1
        rows = self.getChildrenByName(dData, 'Row')
        for row in rows:
            # Учет индекса строки
            idx = int(row.get('Index', i))
            if idx > i:
                for ii in range(idx-i):
                    ods_row = odf.table.TableRow()
                    ODSTable.addElement(ods_row)
                i = idx+1
            else:
                i += 1
                
            ods_row = self.setRow(row)
            ODSTable.addElement(ods_row)

            span = row.get('Span', None)
            if span:
                i += int(span)

    def _genColumnStyleName(self):
        """
        Генерация имени стиля колонки.
        """
        from services.ic_std.utils import uuid
        return uuid.get_uuid()

    def _genRowStyleName(self):
        """
        Генерация имени стиля cтроки.
        """
        from services.ic_std.utils import uuid
        return uuid.get_uuid()

    def setColumn(self, dData):
        """
        Заполнить колонку.
        @param dData: Словарь данных.
        """
        log.info('column: <%s>' % dData)
        
        kwargs = {}
        
        width = dData.get('Width', None)

        if width:
            width = str(float(width)*DIMENSION_CORRECT)
            # Создать автоматические стили дбя ширин колонок
            ods_col_style = odf.style.Style(name=self._genColumnStyleName(), family='table-column')
            ods_col_properties = odf.style.TableColumnProperties(columnwidth=width, breakbefore='auto')
            ods_col_style.addElement(ods_col_properties)
            self.ods_document.automaticstyles.addElement(ods_col_style)
            
            kwargs['stylename'] = ods_col_style
        else:
            # Ширина колонки не определена
            ods_col_style = None

        cell_style = dData.get('StyleID', None)
        kwargs['defaultcellstylename'] = cell_style

        repeated = dData.get('Span', None)
        if repeated:
            repeated = str(int(repeated)+1)
            kwargs['numbercolumnsrepeated'] = repeated

        hidden = dData.get('Hidden', False)
        if hidden:
            kwargs['visibility'] = 'collapse'

        ods_column = odf.table.TableColumn(**kwargs)
        return ods_column

    def setRow(self, dData):
        """
        Заполнить строку.
        @param dData: Словарь данных.
        """
        log.info('row: <%s>' % dData)
        
        kwargs = dict()
        height = dData.get('Height', None)
        
        if height:
            height = str(float(height)*DIMENSION_CORRECT)
            # Создать автоматические стили дбя высот строк
            style_name = self._genRowStyleName()
            ods_row_style = odf.style.Style(name=style_name, family='table-row')
            ods_row_properties = odf.style.TableRowProperties(rowheight=height, breakbefore='auto')
            ods_row_style.addElement(ods_row_properties)
            self.ods_document.automaticstyles.addElement(ods_row_style)
            
            kwargs['stylename'] = ods_row_style
        else:
            # Высота строки не определена
            ods_row_style = None

        # repeated=dData.get('Span',None)
        # if repeated:
        #     self._row_repeated=int(repeated)+1
        #     self._row_repeated_style=ods_row_style
        #
        # if self._row_repeated>0:
        #     self._row_repeated=self._row_repeated-1
        #     if self._row_repeated_style:
        #         kwargs['stylename']=self._row_repeated_style

        hidden = dData.get('Hidden', False)
        if hidden:
            kwargs['visibility'] = 'collapse'

        # Зарегистрировать стиль
        if ods_row_style:
            self._styles_[style_name] = ods_row_style

        ods_row = odf.table.TableRow(**kwargs)
        
        # Ячейки
        i = 1
        cells = self.getChildrenByName(dData, 'Cell')
        for i_cell, cell in enumerate(cells):
            # Учет индекса ячейки
            idx = int(cell.get('Index', i))
            if idx > i:
                kwargs = dict()
                kwargs['numbercolumnsrepeated'] = (idx-i)

                style_id = self._find_prev_style(cells[:i_cell])
                if style_id:
                    kwargs['stylename'] = self._styles_.get(style_id, None)
                    
                ods_cell = odf.table.CoveredTableCell(**kwargs)
                ods_row.addElement(ods_cell)
                
                i = idx+1
            else:
                i += 1

            ods_cell = self.setCell(cell)
            if ods_cell:
                ods_row.addElement(ods_cell)

            # Учет объединенных ячеек
            merge = int(cell.get('MergeAcross', 0))
            if merge > 0:
                kwargs = dict()
                kwargs['numbercolumnsrepeated'] = merge

                style_id = self._find_prev_style(cells[:i_cell])
                if style_id:
                    kwargs['stylename'] = self._styles_.get(style_id, None)
                
                ods_cell = odf.table.CoveredTableCell(**kwargs)
                ods_row.addElement(ods_cell)
                i += merge
            
        return ods_row

    def _find_prev_style(self, lCells):
        """
        Поиск стиля определенного в предыдущей ячейке.
        @param lCells: Список предыдущих ячеек.
        @return: Идентификатор исокмого стиля или None если стиль не определен.
        """
        for cell in reversed(lCells):
            if 'StyleID' in cell:
                return cell.get('StyleID', None)
        return None
        
    def getCellValue(self, dData):
        """
        Получить значение ячейки.
        @param dData: Словарь данных.
        """
        type = self.getCellType(dData)
        
        if type != 'string':
            dates = self.getChildrenByName(dData, 'Data')
            value = ''
            if dates:
                value = unicode(dates[0].get('value', ''), DEFAULT_ENCODE)
            return value
        return None
        
    def getCellType(self, dData):
        """
        Тип значения ячейки.
        @param dData: Словарь данных.
        """
        dates = self.getChildrenByName(dData, 'Data')
        type = 'string'
        if dates:
            type = dates[0].get('Type', 'string').lower()
        
        if type == 'number':
            type = 'float'
        elif type == 'percentage':
            # ВНИМЕНИЕ! Здесь необходимо сделать проверку
            # на соответствие данных процентному типу
            str_value = dates[0].get('value', 'None')
            try:
                value = float(str_value)
            except:
                log.warning('Value <%s> is not [percentage] type' % str_value)
                type = 'string'
        return type            
    
    # Имена колонок Excel в формате A1
    COLS_A1 = None

    def _getColsA1(self):
        """
        Имена колонок Excel в формате A1.
        """
        if self.COLS_A1:
            return self.COLS_A1

        sAlf1 = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        sAlf2 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        self.COLS_A1 = []
        for s1 in sAlf1:
            for s2 in sAlf2:
                self.COLS_A1.append((s1+s2).strip())
                if len(self.COLS_A1) == 256:
                    return self.COLS_A1
        return self.COLS_A1
        
    R1_FORMAT = r'R[0-9]{1,5}'
    C1_FORMAT = r'C[0-9]{1,5}'

    def _getA1(self, sR1C1):
        """
        Преабразовать адрес из формата R1C1 в A1.
        """
        parse = re.findall(self.R1_FORMAT, sR1C1)
        row = 1
        if parse:
            row = int(parse[0][1:])
        parse = re.findall(self.C1_FORMAT, sR1C1)
        col = 1
        if parse:
            col = int(parse[0][1:])
        cols_a1 = self._getColsA1()
        return cols_a1[col-1]+str(row)

    ALPHA_FORMAT = r'[A-Z]{1,2}'
    DIGIT_FORMAT = r'[0-9]{1,5}'

    def _getR1C1(self, sA1):
        """
        Преабразовать адрес из формата A1 в R1C1.
        """
        parse = re.findall(self.DIGIT_FORMAT, sA1)
        row = 1
        if parse:
            row = int(parse[0])
        parse = re.findall(self.ALPHA_FORMAT, sA1)
        col = 1
        if parse:
            cols_a1 = self._getColsA1()
            try:
                col = cols_a1.index(parse[0]) + 1
            except:
                pass
        return 'R%dC%d' % (row, col)
        
    R1C1_FORMAT = r'R[0-9]{1,5}C[0-9]{1,5}'

    def _R1C1Fmt2A1Fmt(self, sFormula):
        """
        Перевод формулы из формата R1C1 в формат A1.
        @param sFormula: Формула в строковом представлении.
        @return: Строка транслированной формулы.
        """
        parse_all = re.findall(self.R1C1_FORMAT, sFormula)
        for replace_addr in parse_all:
            a1 = self._getA1(replace_addr)
            if self._is_sheetAddress(replace_addr, sFormula):
                a1 = '.'+a1
            sFormula = sFormula.replace(replace_addr, a1)
        return sFormula

    def _is_sheetAddress(self, sAddress, sFormula):
        """
        Адресация ячейки с указанием листа? Например Лист1.A1
        @param sAddress: Адресс ячейки.
        @param sFormula: Формула в строковом представлении.
        @return: True/False.
        """
        # if isinstance(sFormula, str):
        #    sFormula = unicode(sFormula, DEFAULT_ENCODE)
        if sAddress in sFormula:
            i = sFormula.index(sAddress)
            if i > 0:
                return sFormula[i-1].isalnum()
            else:
                return False
        return None

    A1_FORMAT = r'\.[A-Z]{1,2}[0-9]{1,5}'

    def _A1Fmt2R1C1Fmt(self, sFormula):
        """
        Перевод формулы из формата A1 в формат R1C1.
        @param sFormula: Формула в строковом представлении.
        @return: Строка транслированной формулы.
        """
        parse_all = re.findall(self.A1_FORMAT, sFormula)
        for replace_addr in parse_all:
            r1c1 = self._getR1C1(replace_addr)
            sFormula = sFormula.replace(replace_addr, r1c1)
        return sFormula
        
    def _translateR1C1Formula(self, sFormula):
        """
        Перевод формулы из формата R1C1 в формат ODS файла.
        @param sFormula: Формула в строковом представлении.
        @return: Строка транслированной формулы.
        """
        return self._R1C1Fmt2A1Fmt(sFormula)

    def _translateA1Formula(self, sFormula):
        """
        Перевод формулы из формата ODS(A1) в формат R1C1.
        @param sFormula: Формула в строковом представлении.
        @return: Строка транслированной формулы.
        """
        return self._A1Fmt2R1C1Fmt(sFormula)
        
    def setCell(self, dData):
        """
        Заполнить ячейку.
        @param dData: Словарь данных.
        """
        log.info('cell: <%s>' % dData)

        properties = {}
        ods_type = self.getCellType(dData)
        properties['valuetype'] = ods_type
        style_id = dData.get('StyleID', None)
        if style_id:
            ods_style = self._styles_.get(style_id, None)
            properties['stylename'] = ods_style
        
        merge_across = int(dData.get('MergeAcross', 0))
        if merge_across:
            merge_across = str(merge_across+1)
            properties['numbercolumnsspanned'] = merge_across
            
        merge_down = int(dData.get('MergeDown', 0))
        if merge_down:
            merge_down = str(merge_down+1)
            properties['numberrowsspanned'] = merge_down
            
        formula = dData.get('Formula', None)
        if formula:
            properties['formula'] = self._translateR1C1Formula(formula)
        else:
            value = self.getCellValue(dData)
            properties['value'] = value
        log.debug('TableCell properties <%s>' % properties)
        ods_cell = odf.table.TableCell(**properties)
            
        dates = self.getChildrenByName(dData,'Data')
        # Разбить на строки
        values = self.getDataValues(dData)
        log.info('values cell: <%s>' % values)
        for data in dates:
            for val in values:
                ods_data = self.setData(data, style_id, val)
                if ods_data:
                    ods_cell.addElement(ods_data)
            
        return ods_cell
        
    def getDataValues(self, dData):
        """
        Получить значение ячейки с разбитием по строкам.
        @param dData: Словарь данных.
        """
        dates = self.getChildrenByName(dData, 'Data')
        value = ''
        if dates:
            value = unicode(dates[0].get('value', ''), DEFAULT_ENCODE)
        return value.split(SPREADSHEETML_CR)
    
    def setData(self, dData, sStyleID=None, sValue=None):
        """
        Заполнить ячейку данными.
        @param dData: Словарь данных.
        @param sStyleID: Идентификатор стиля.
        @param sValue: Значение строки.
        """
        log.info('data: <%s> style: <%s>' % (dData, sStyleID))
        
        ods_style = None
        if sStyleID:
            ods_style = self._styles_.get(sStyleID, None)

        ods_data = None
        if sValue:
            # Просто текст
            if sStyleID and sStyleID != DEFAULT_STYLE_ID:
                ods_data = odf.text.P()
                style_span = odf.text.Span(stylename=ods_style, text=sValue)
                ods_data.addElement(style_span)
            else:
                ods_data = odf.text.P(text=sValue)
        return ods_data
        
    def Load(self, sFileName):
        """
        Загрузить из ODS файла.
        @param sFileName: Имя ODS файла.
        @return: Словарь данных или None в случае ошибки.
        """
        if not os.path.exists(sFileName):
            # Если файл не существует то верноть None
            log.warning('File <%s> not exists' % sFileName)
            return None
        else:
            try:
                return self._loadODS(sFileName)
            except:
                log.error('Open file <%s>' % sFileName)
                raise                
        
    def _loadODS(self, sFileName):
        """
        Загрузить из ODS файла.
        @param sFileName: Имя ODS файла.
        @return: Словарь данных или None в случае ошибки.
        """
        if isinstance(sFileName, str):
            # ВНИМАНИЕ! Перед загрузкой надо имя файла сделать
            # Юникодной иначе падает по ошибке в функции load
            sFileName = unicode(sFileName, DEFAULT_ENCODE)
        self.ods_document = odf.opendocument.load(sFileName)
        
        self.xmlss_data = {'name': 'Calc', 'children': []}
        ods_workbooks = self.ods_document.getElementsByType(odf.opendocument.Spreadsheet)
        if ods_workbooks:
            workbook_data = self.readWorkbook(ods_workbooks[0])
            self.xmlss_data['children'].append(workbook_data)
        return self.xmlss_data
    
    def readWorkbook(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о книге.
        @pararm ODSElement: ODS элемент соответствующий книги Excel.
        """
        data = {'name': 'Workbook', 'children': []}
        
        styles_data = self.readStyles()
        data['children'].append(styles_data)
        
        ods_tables = ODSElement.getElementsByType(odf.table.Table)
        if ods_tables:
            for ods_table in ods_tables:
                worksheet_data = self.readWorksheet(ods_table)
                data['children'].append(worksheet_data)
        
        return data

    def readNumberStyles(self, *ODSStyles):
        """
        Прочитать данные о стилях числовых форматов.
        @param ODSStyles: Список стилей.
        """
        if not ODSStyles:
            log.warning('Not define ODS styles for read Number styles')
            return {}
            
        result = {}
        for ods_styles in ODSStyles:
            num_styles = ods_styles.getElementsByType(odf.number.NumberStyle)
            percentage_styles = ods_styles.getElementsByType(odf.number.PercentageStyle)
            styles = num_styles + percentage_styles
            if styles:
                for style in styles:
                    result[style.getAttribute('name')] = style
    
        log.debug('NUMBER STYLES <%s>' % result)
        return result
        
    def readStyles(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о стилях.
        @pararm ODSElement: ODS элемент соответствующий стилям книги Excel.
        """
        data = {'name': 'Styles', 'children': []}
        ods_styles = self.ods_document.automaticstyles.getElementsByType(odf.style.Style) + \
            self.ods_document.styles.getElementsByType(odf.style.Style) + \
            self.ods_document.masterstyles.getElementsByType(odf.style.Style)

        # Стили числовых форматов
        self._number_styles_ = self.readNumberStyles(self.ods_document.automaticstyles,
                                                     self.ods_document.styles,
                                                     self.ods_document.masterstyles)
        
        # log.debug('STYLES <%s>' % ods_styles)
        
        for ods_style in ods_styles:
            style = self.readStyle(ods_style)
            data['children'].append(style)
                        
        return data

    def readStyle(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о стиле.
        @pararm ODSElement: ODS элемент соответствующий стилю Excel.
        """
        data = {'name': 'Style', 'children': []}
        id = ODSElement.getAttribute('name')
        data['ID'] = id
        
        data_style_name = ODSElement.getAttribute('datastylename')
        if data_style_name:
            number_style = self._number_styles_.get(data_style_name, None)
            if number_style:
                number_format_data = self.readNumberFormat(number_style)
                if number_format_data:
                    data['children'].append(number_format_data)

        # Чтение шрифта
        txt_properties = ODSElement.getElementsByType(odf.style.TextProperties)
        if txt_properties:
            font_data = self.readFont(txt_properties[0])
            if font_data:
                data['children'].append(font_data)

        # Чтение бордеров
        tab_cell_properties = ODSElement.getElementsByType(odf.style.TableCellProperties)
        if tab_cell_properties:
            borders_data = self.readBorders(tab_cell_properties[0])
            if borders_data:
                data['children'].append(borders_data)

        # Чтение интерьера
        if tab_cell_properties:
            interior_data = self.readInterior(tab_cell_properties[0])
            if interior_data:
                data['children'].append(interior_data)

        # Чтение выравнивания
        align_data = {}
        paragraph_properties = ODSElement.getElementsByType(odf.style.ParagraphProperties)
        if paragraph_properties:
            paragraph_align_data = self.readAlignmentParagraph(paragraph_properties[0])
            if paragraph_align_data:
                align_data.update(paragraph_align_data)

        if tab_cell_properties:
            cell_align_data = self.readAlignmentCell(tab_cell_properties[0])
            if cell_align_data:
                align_data.update(cell_align_data)
                
        if align_data:
            data['children'].append(align_data)
        
        log.debug('Read STYLE %s : %s : %s' % (id, txt_properties, tab_cell_properties))
        
        return data        
    
    def readNumberFormat(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о формате числового представления.
        @pararm ODSElement: ODS элемент соответствующий стилю числового представления.
        """
        if ODSElement is None:
            log.warning('Not define ODSElement <%s>' % ODSElement)
            return None
        
        numbers = ODSElement.getElementsByType(odf.number.Number)
        if not numbers:
            log.warning('Not define numbers in ODSElement <%s>' % ODSElement)
            return None
        else:
            number = numbers[0]

        decimalplaces_str = number.getAttribute('decimalplaces')
        decimalplaces = int(decimalplaces_str) if decimalplaces_str not in ('None', 'none', 'NONE', None) else 0
        minintegerdigits_str = number.getAttribute('minintegerdigits')
        minintegerdigits = int(minintegerdigits_str) if minintegerdigits_str not in ('None', 'none', 'NONE', None) else 0
        grouping = number.getAttribute('grouping')
        percentage = 'percentage-style' in ODSElement.tagName

        decimalplaces_format = ','+'0' * decimalplaces if decimalplaces else ''
        minintegerdigits_format = '0' * minintegerdigits
        grouping_format = '#' * (4 - minintegerdigits) if minintegerdigits<4 else ''
        percentage_format = '%' if percentage else ''
        
        if grouping and (grouping not in ('None', 'none', 'NONE', 'false')):
            format = list(grouping_format + minintegerdigits_format)
            format_result = []
            count = 3
            for i in range(len(format)-1, -1, -1):
                format_result = [format[i]] + format_result
                count = count - 1
                if not count:
                    format_result = [' '] + format_result
                    count = 3
            number_format = ''.join(format_result) + decimalplaces_format + percentage_format
        else:
            number_format = minintegerdigits_format + decimalplaces_format + percentage_format
            
        log.debug('NUMBER FORMAT %s' % (number_format))
        data = {'name': 'NumberFormat', 'children': [], 'Format': number_format}
        return data
        
    def readFont(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о шрифте.
        @pararm ODSElement: ODS элемент соответствующий ствойствам текста стиля.
        """
        name = ODSElement.getAttribute('fontname')
        name = name if name else ODSElement.getAttribute('fontfamily')
        size = ODSElement.getAttribute('fontsize')
        bold = ODSElement.getAttribute('fontweight')
        italic = ODSElement.getAttribute('fontstyle')
        # Зачеркивание
        through_style = ODSElement.getAttribute('textlinethroughstyle')
        # through_type = ODSElement.getAttribute('textlinethroughtype')
        # Подчеркивание
        underline_style = ODSElement.getAttribute('textunderlinestyle')
        # underline_width = ODSElement.getAttribute('textunderlinewidth')

        data = {'name': 'Font', 'children': []}
        if name and (name not in ('None', 'none', 'NONE')):
            data['FontName'] = name
            
        if size and (size not in ('None', 'none', 'NONE')):
            if size[-2:] == 'pt':
                size = size[:-2]
            data['Size'] = size
        if bold and (bold not in ('None', 'none', 'NONE', 'normal')):
            data['Bold'] = '1'
        if italic and (italic not in ('None', 'none', 'NONE', 'normal')):
            data['Italic'] = '1'

        if through_style and (through_style not in ('None', 'none', 'NONE', 'normal')):
            data['StrikeThrough'] = '1'
        if underline_style and (underline_style not in ('None', 'none', 'NONE', 'normal')):
            data['Underline'] = 'Single'

        log.debug('Read FONT: %s : %s : %s : %s : %s : %s' % (name, size, bold, italic, through_style, underline_style))
        
        return data
        
    def readInterior(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о бордерах.
        @pararm ODSElement: ODS элемент соответствующий ствойствам ячейки таблицы стиля.
        """
        data = {'name': 'Interior', 'children': []}

        color = ODSElement.getAttribute('backgroundcolor')

        log.debug('Read INTERIOR: color <%s>' % color)

        if color and (color not in ('None', 'none', 'NONE')):
            data['Color'] = color.strip()

        return data

    def readBorders(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о бордерах.
        @pararm ODSElement: ODS элемент соответствующий ствойствам ячейки таблицы стиля.
        """
        data = {'name': 'Borders', 'children': []}
        
        all_border = ODSElement.getAttribute('border')
        left = ODSElement.getAttribute('borderleft')
        right = ODSElement.getAttribute('borderright')
        top = ODSElement.getAttribute('bordertop')
        bottom = ODSElement.getAttribute('borderbottom')
        
        log.debug('BORDERS: border %s Left: %s Right: %s Top: %s Bottom: %s' % (all_border, left, right, top, bottom))
        
        if all_border and (all_border not in ('None', 'none', 'NONE')):
            border = self.parseBorder(all_border, 'Left')
            if border:
                data['children'].append(border)
            border = self.parseBorder(all_border, 'Right')
            if border:
                data['children'].append(border)
            border = self.parseBorder(all_border, 'Top')
            if border:
                data['children'].append(border)
            border = self.parseBorder(all_border, 'Bottom')
            if border:
                data['children'].append(border)
                
        if left and (left not in ('None', 'none', 'NONE')):
            border = self.parseBorder(left, 'Left')
            if border:
                data['children'].append(border)
            
        if right and (right not in ('None', 'none', 'NONE')):
            border = self.parseBorder(right, 'Right')
            if border:
                data['children'].append(border)
        
        if top and (top not in ('None', 'none', 'NONE')):
            border = self.parseBorder(top, 'Top')
            if border:
                data['children'].append(border)
            
        if bottom and (bottom not in ('None', 'none', 'NONE')):
            border = self.parseBorder(bottom, 'Bottom')
            if border:
                data['children'].append(border)
            
        return data
        
    def readAlignmentParagraph(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о выравнивании текста.
        @pararm ODSElement: ODS элемент соответствующий ствойствам параграфа.
        """
        data = {'name': 'Alignment', 'children': []}

        text_align = ODSElement.getAttribute('textalign')
        vert_align = ODSElement.getAttribute('verticalalign')
        
        if text_align == 'start':
            data['Horizontal'] = 'Left'
        elif text_align == 'end':
            data['Horizontal'] = 'Right'
        elif text_align == 'center':
            data['Horizontal'] = 'Center'
        elif text_align == 'justify':
            data['Horizontal'] = 'Justify'

        if vert_align == 'top':
            data['Vertical'] = 'Top'
        elif vert_align == 'bottom':
            data['Vertical'] = 'Bottom'
        elif vert_align == 'middle':
            data['Vertical'] = 'Center'
        elif vert_align == 'justify':
            data['Vertical'] = 'Justify'
            
        log.debug('ALIGNMENT PARAGRAPH: %s:%s' % (text_align, vert_align))
        
        return data
        
    def readAlignmentCell(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о выравнивании текста.
        @pararm ODSElement: ODS элемент соответствующий ствойствам ячейки.
        """
        data = {'name': 'Alignment', 'children': []}

        vert_align = ODSElement.getAttribute('verticalalign')
        wrap_txt = ODSElement.getAttribute('wrapoption')

        if vert_align == 'top':
            data['Vertical'] = 'Top'
        elif vert_align == 'bottom':
            data['Vertical'] = 'Bottom'
        elif vert_align == 'middle':
            data['Vertical'] = 'Center'
            
        if wrap_txt and wrap_txt == 'wrap':
            data['WrapText'] = '1'
            
        log.debug('ALIGNMENT CELL: %s' % vert_align)
        
        return data
        
    def parseBorder(self, sData, sPosition=None):
        """
        Распарсить бордер.
        @param sData: Строка данных в виде <1pt solid #000000>.
        @param sPosition: Позиция бордера.
        @return: Заполненный словарь бордера.
        """
        border = None
        if sData:
            border_data = self.parseBorderData(sData)
            if border_data:
                border = {'name': 'Border', 'Position': sPosition,
                          'Weight': border_data.get('weight', '1'),
                          'LineStyle': self._lineStyles_ODS2SpreadsheetML.get(border_data.get('line', None), 'Continuous'),
                          'Color': border_data.get('color', '000000')}
            log.debug('PARSE BORDER: %s : %s' % (sData, border))
            
        return border
            
    def parseBorderData(self, sData):
        """
        Распарсить данные бордера.
        @param sData: Строка данных в виде <1pt solid #000000>.
        @return: Словарь {'weight':1,'line':'solid','color':'#000000'}.
        """
        if sData in ('None', 'none', 'NONE'):
            return None
        
        weight_pattern_pt = r'([\.\d]+)pt'
        weight_pattern_cm = r'([\.\d]+)cm'
        line_style_pattern = r' [a-zA-Z]* '
        color_pattern = r'#......'
        
        result = {}

        weight_pt = re.findall(weight_pattern_pt, sData)
        weight_cm = re.findall(weight_pattern_cm, sData)
        if weight_pt:
            result['weight'] = weight_pt[0]
        elif weight_cm:
            # Ширина может задаватся в см поэтому нужно преобразовать в пт
            result['weight'] = unicode(float(weight_cm[0])*CM2PT_CORRECT)
            
        line_style = re.findall(line_style_pattern, sData)
        if line_style:
            result['line'] = line_style[0].strip()

        color = re.findall(color_pattern, sData)
        if color:
            result['color'] = color[0]
        
        return result

    def readWorksheet(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о листе.
        @param ODSElement: ODS элемент соответствующий листу.
        """
        data = {'name': 'Worksheet', 'children': []}
        name = ODSElement.getAttribute('name').encode(DEFAULT_ENCODE)
        
        log.debug('WORKSHEET: <%s : %s>' % (type(name), name))
        
        data['Name'] = name
        
        table = {'name': 'Table', 'children': []}
        
        # Колонки
        ods_columns = ODSElement.getElementsByType(odf.table.TableColumn)
        for ods_column in ods_columns:
            column_data = self.readColumn(ods_column)
            table['children'].append(column_data)
            
        # Строки
        ods_rows = ODSElement.getElementsByType(odf.table.TableRow)
        for i, ods_row in enumerate(ods_rows):
            row_data = self.readRow(ods_row, table, data, i)
            table['children'].append(row_data)
      
        data['children'].append(table)
        
        # Параметры страницы
        ods_pagelayouts = self.ods_document.automaticstyles.getElementsByType(odf.style.PageLayout)
        worksheet_options = self.readWorksheetOptions(ods_pagelayouts)
        if worksheet_options:
            data['children'].append(worksheet_options)
        
        return data
    
    def readWorksheetOptions(self, ODSPageLayouts):
        """
        Прочитать из ODS файла данные о параметрах страницы.
        @param ODSPageLayouts: Список найденных параметров страницы.
        """
        if not ODSPageLayouts:
            log.warning(u'Not define page layout')
            return None

        log.debug(u'Set default worksheet options')
        options = {'name': 'WorksheetOptions',
                   'children': [{'name': 'PageSetup',
                                 'children': [{'name': 'Layout'},
                                              {'name': 'PageMargins'},
                                              ]
                                 },
                                {'name': 'Print',
                                 'children': [{'name': 'PaperSizeIndex',
                                               'value': '9'}
                                              ]
                                 }
                                ]
                   }
        
        for pagelayout in ODSPageLayouts:
            properties = pagelayout.getElementsByType(odf.style.PageLayoutProperties)
            if properties:
                properties = properties[0]
                orientation = properties.getAttribute('printorientation')
                margin = properties.getAttribute('margin')
                margin_top = properties.getAttribute('margintop')
                margin_bottom = properties.getAttribute('marginbottom')
                margin_left = properties.getAttribute('marginleft')
                margin_right = properties.getAttribute('marginright')
                page_width = properties.getAttribute('pagewidth')
                page_height = properties.getAttribute('pageheight')
                fit_to_page = properties.getAttribute('scaletopages')
                scale_to = properties.getAttribute('scaleto')

                if orientation:
                    options['children'][0]['children'][0]['Orientation'] = orientation.title()
                if margin:
                    options['children'][0]['children'][1]['Top'] = self._dimension_ods2xml(margin)
                    options['children'][0]['children'][1]['Bottom'] = self._dimension_ods2xml(margin)
                    options['children'][0]['children'][1]['Left'] = self._dimension_ods2xml(margin)
                    options['children'][0]['children'][1]['Right'] = self._dimension_ods2xml(margin)
                if margin_top:
                    options['children'][0]['children'][1]['Top'] = self._dimension_ods2xml(margin_top)
                if margin_bottom:
                    options['children'][0]['children'][1]['Bottom'] = self._dimension_ods2xml(margin_bottom)
                if margin_left:
                    options['children'][0]['children'][1]['Left'] = self._dimension_ods2xml(margin_left)
                if margin_right:
                    options['children'][0]['children'][1]['Right'] = self._dimension_ods2xml(margin_right)
                if fit_to_page or (scale_to == (1, 1)):
                    options['children'].append({'name': 'FitToPage'})

                if page_width and page_height:
                    # Установить размер листа
                    options['children'][1]['children'][0]['value'] = self._getExcelPaperSizeIndex(page_width, page_height)
                # ВНИМАНИЕ! Обычно параметры печати указанные в начале и являются
                # параметрами по умолчанию. Поэтому пропускаем все остальные
                break
            else:
                log.debug(u'Not define worksheet options')
                continue
        return options

    def _getExcelPaperSizeIndex(self, fPageWidth, fPageHeight):
        """
        Определить по размеру листа его индекс в списке Excel.
        """
        paper_format = self._getPaperSizeFormat(fPageWidth, fPageHeight)
        
        if paper_format is None:
            # По умолчанию A4
            return '9'
        elif paper_format == A4_PAPER_FORMAT:
            return '9'
        elif paper_format == A3_PAPER_FORMAT:
            return '8'
        return None
        
    def _getPaperSizeFormat(self, fPageWidth, fPageHeight):
        """
        Определить по размеру листа его формат.
        @param fPageWidth: Ширина листа в см.
        @param fPageHeight: Высота листа в см.
        @return: A4 или A3.
        """
        if type(fPageWidth) in (str, unicode):
            page_width_txt = fPageWidth.replace('cm', '').replace('mm', '')
            fPageWidth = float(page_width_txt)
        if type(fPageHeight) in (str, unicode):
            page_height_txt = fPageHeight.replace('cm', '').replace('mm', '')
            fPageHeight = float(page_height_txt)

        # Если задается в сантиметрах
        if round(fPageWidth, 1) == 21.0 and round(fPageHeight, 1) == 29.7:
            return A4_PAPER_FORMAT
        elif round(fPageWidth, 1) == 29.7 and round(fPageHeight, 1) == 21.0:
            return A4_PAPER_FORMAT
        elif round(fPageWidth, 1) == 29.7 and round(fPageHeight, 1) == 42.0:
            return A3_PAPER_FORMAT
        elif round(fPageWidth, 1) == 42.0 and round(fPageHeight, 1) == 29.7:
            return A3_PAPER_FORMAT
        # Если задается в миллиметрах
        elif round(fPageWidth, 1) == 210.0 and round(fPageHeight, 1) == 297.0:
            return A4_PAPER_FORMAT
        elif round(fPageWidth, 1) == 297.0 and round(fPageHeight, 1) == 210.0:
            return A4_PAPER_FORMAT
        elif round(fPageWidth, 1) == 297.0 and round(fPageHeight, 1) == 420.0:
            return A3_PAPER_FORMAT
        elif round(fPageWidth, 1) == 420.0 and round(fPageHeight, 1) == 297.0:
            return A3_PAPER_FORMAT
        return None
        
    def readColumn(self, ODSElement=None):
        """
        Прочитать из ODS файла данные о колонке.
        @pararm ODSElement: ODS элемент соответствующий колонке.
        """
        data = {'name': 'Column', 'children': []}
        style_name = ODSElement.getAttribute('stylename')
        default_cell_style_name = ODSElement.getAttribute('defaultcellstylename')
        repeated = ODSElement.getAttribute('numbercolumnsrepeated')
        hidden = ODSElement.getAttribute('visibility')

        if style_name:
            # Определение ширины колонки
            column_width = None
            ods_styles = self.ods_document.automaticstyles.getElementsByType(odf.style.Style)
            find_style = [ods_style for ods_style in ods_styles if ods_style.getAttribute('name') == style_name]
            if find_style:
                ods_style = find_style[0]
                ods_column_properties = ods_style.getElementsByType(odf.style.TableColumnProperties)
                if ods_column_properties:
                    ods_column_property = ods_column_properties[0]
                    column_width = self._dimension_ods2xml(ods_column_property.getAttribute('columnwidth'))
                    if column_width:
                        data['Width'] = column_width
            log.debug('COLUMN: %s Width: %s Cell style: %s' % (style_name, column_width, default_cell_style_name))
        
        if default_cell_style_name and (default_cell_style_name not in ('Default', 'None', 'none', 'NONE')):
            data['StyleID'] = default_cell_style_name

        if repeated and repeated != 'None':
            repeated = str(int(repeated)-1)
            data['Span'] = repeated

        if hidden and hidden == 'collapse':
            data['Hidden'] = True

        return data
    
    def _dimension_ods2xml(self, sDimension):
        """
        Перевод размеров из представления ODS в XML.
        @param sDimension: Строковое представление размера.
        """
        if not sDimension:
            return None
        elif (len(sDimension) > 2) and (sDimension[-2:] == 'cm'):
            # Размер указан в сентиметрах?
            return str(float(sDimension[:-2]) * 28)
        elif (len(sDimension) > 2) and (sDimension[-2:] == 'mm'):
            # Размер указан в миллиметрах?
            return str(float(sDimension[:-2]) * 2.8)
        else:
            # Размер указан в точках
            return str(float(sDimension)/DIMENSION_CORRECT)

    def _add_page_break(self, dWorksheet, iRow):
        """
        Добавить разрыв страницы.
        @param dWorksheet: Словарь, описывающий лист.
        @param iRow: Номер строки.
        """
        find_page_breaks = [child for child in dWorksheet['children'] if child['name'] == 'PageBreaks']
        if find_page_breaks:
            data = find_page_breaks[0]
        else:
            data = {'name': 'PageBreaks', 'children': [{'name': 'RowBreaks', 'children': []}]}
            dWorksheet['children'].append(data)

        row_break = {'name': 'RowBreak', 'children': [{'name': 'Row', 'value': iRow}]}
        data['children'][0]['children'].append(row_break)
        return data

    def readRow(self, ODSElement=None, dTable=None, dWorksheet=None, iRow=-1):
        """
        Прочитать из ODS файла данные о строке.
        @pararm ODSElement: ODS элемент соответствующий строке.
        @param dTable: Словарь, описывающий таблицу.
        @param dWorksheet: Словарь, описывающий лист.
        @param iRow: Номер строки.
        """
        data = {'name': 'Row', 'children': []}
        style_name = ODSElement.getAttribute('stylename')
        repeated = ODSElement.getAttribute('numberrowsrepeated')
        hidden = ODSElement.getAttribute('visibility')

        if style_name:
            # Определение высоты строки
            row_height = None
            ods_styles = self.ods_document.automaticstyles.getElementsByType(odf.style.Style)
            find_style = [ods_style for ods_style in ods_styles if ods_style.getAttribute('name') == style_name]
            if find_style:
                ods_style = find_style[0]
                ods_row_properties = ods_style.getElementsByType(odf.style.TableRowProperties)
                if ods_row_properties:
                    ods_row_property = ods_row_properties[0]
                    row_height = self._dimension_ods2xml(ods_row_property.getAttribute('rowheight'))
                    if row_height:
                        data['Height'] = row_height
                    # Разрывы страниц
                    page_break = ods_row_property.getAttribute('breakbefore')
                    if page_break and page_break == 'page' and dWorksheet:
                        self._add_page_break(dWorksheet, iRow)
                    log.debug('ROW: %s \tHeight: %s \tPageBreak: %s \tStyle: %s' % (style_name, row_height, page_break, style_name))
        
        if repeated and (repeated not in ('None', 'none', 'NONE')):
            # Дополниетельное условие необходимо для
            # исключения ошибочной ситуации когда параметры строки
            # дублируются на все последующие строки
            # (в LibreOffice это сделано для дублирования стиля ячеек
            # построчно в конце листа)
            i_repeated = int(repeated)
            if i_repeated <= LIMIT_ROWS_REPEATED:
                repeated = i_repeated-1
                if dTable:
                    for i in range(repeated):
                        dTable['children'].append(self.readRow(ODSElement))
                else:
                    data['Span'] = str(repeated)

        if hidden and hidden == 'collapse':
            data['Hidden'] = True

        # Обработка ячеек
        ods_cells = ODSElement.childNodes
        
        i = 1
        set_idx = False
        for ods_cell in ods_cells:
            if ods_cell.qname[-1] == 'covered-table-cell':
                repeated = ods_cell.getAttribute('numbercolumnsrepeated')
                if repeated and (repeated not in ('None', 'none', 'NONE')):
                    # Учет индекса и пропущенных ячеек
                    i += int(repeated)  # +1
                    set_idx = True
                else:
                    # Стоит просто ячейка и ее надо учесть
                    i += 1
                    set_idx = True
            elif ods_cell.qname[-1] == 'table-cell':
                cell_data = self.readCell(ods_cell, i if set_idx else None)
                data['children'].append(cell_data)
                if set_idx:
                    set_idx = False

                repeated = ods_cell.getAttribute('numbercolumnsrepeated')
                if repeated and (repeated not in ('None', 'none', 'NONE')):
                    # Учет индекса и пропущенных ячеек
                    i_repeated = int(repeated)
                    if i_repeated < LIMIT_COLUMNS_REPEATED:
                        for ii in range(i_repeated-1):
                            cell_data = self.readCell(ods_cell)
                            data['children'].append(cell_data)
                            i += 1
                        # ВНИМАНИЕ! Здесь необходимо добавить 1 иначе таблицы/штампы могут "плыть"
                        i += 1
                    else:
                        i += i_repeated     # +1
                        set_idx = True
                else:
                    i += 1
                        
        return data
    
    def hasODSAttribute(self, ODSElement, sAttrName):
        """
        Имеется в ODS элементе атрибут с таким именем?
        @pararm ODSElement: ODS элемент.
        @param sAttrName: Имя атрибута.
        @return: True/False.
        """
        return sAttrName in [attr[-1].replace('-', '') for attr in ODSElement.attributes.keys()]
        
    def readCell(self, ODSElement=None, iIndex=None):
        """
        Прочитать из ODS файла данные о ячейке.
        @pararm ODSElement: ODS элемент соответствующий ячейке.
        @type iIndex: C{int}
        @param iIndex: Индекс ячейки, если необходимо указать.
        """
        data = {'name': 'Cell', 'children': []}
        if iIndex:
            data['Index'] = str(iIndex)

        style_name = ODSElement.getAttribute('stylename')
        if style_name:
            data['StyleID'] = style_name

        formula = ODSElement.getAttribute('formula')
        if formula:
            data['Formula'] = self._translateA1Formula(formula)
            
        merge_across = None
        if self.hasODSAttribute(ODSElement, 'numbercolumnsspanned'):
            numbercolumnsspanned = ODSElement.getAttribute('numbercolumnsspanned')
            if numbercolumnsspanned:
                merge_across = int(numbercolumnsspanned)-1
                data['MergeAcross'] = merge_across
            
        merge_down = None
        if self.hasODSAttribute(ODSElement, 'numberrowsspanned'):
            numberrowsspanned = ODSElement.getAttribute('numberrowsspanned')
            if numberrowsspanned:
                merge_down = int(numberrowsspanned)-1
                data['MergeDown'] = merge_down
        
        ods_data = ODSElement.getElementsByType(odf.text.P)
        if ods_data:
            value = ODSElement.getAttribute('value')
            valuetype = ODSElement.getAttribute('valuetype')
            cur_data = None
            for i, ods_txt in enumerate(ods_data):
                data_data = self.readData(ods_txt, value, valuetype)
                if not i:
                    cur_data = data_data
                else:
                    cur_data['value'] += SPREADSHEETML_CR+data_data['value']
            data['children'].append(cur_data)
        
        log.debug('CELL Style: %s MergeAcross: %s MergeDown: %s' % (style_name, merge_across, merge_down))
        
        return data
   
    def readData(self, ODSElement=None, sValue=None, sValueType=None):
        """
        Прочитать из ODS файла данные.
        @pararm ODSElement: ODS элемент соответствующий данным ячейки.
        @param sValue: Строковое значение ячейки.
        @param sValueType: Строковое представление типа значения ячейки.
        """
        data = {'name': 'Data', 'children': []}
        if sValue and sValue != 'None':
            data['value'] = sValue
        else:
            txt = u''.join([unicode(child) for child in ODSElement.childNodes])
            value = txt.encode(DEFAULT_ENCODE)
            data['value'] = value

        if sValueType:
            data['Type'] = str(sValueType).title()
        
        log.debug('DATA: %s' % sValueType)
        return data


def test_save(XMLFileName):
    """
    Функция тестирования.
    """
    import icexcel
    excel = icexcel.icVExcel()
    excel.Load(XMLFileName)
    data = excel.getData()
    
    ods = icODS()
    ods.Save('./testfiles/test.ods', data)


def test_load(ODSFileName):
    """
    Функция тестирования.
    """
    ods = icODS()
    ods.Load(ODSFileName)


def test_complex(SrcODSFileName, DestODSFileName):
    """
    Функция тестирования чтения/записи ODS файла.
    """
    ods = icODS()
    data = ods.Load(SrcODSFileName)
    
    import icexcel
    excel = icexcel.icVExcel()
    excel._data = data
    excel.SaveAs(DestODSFileName)


def test_1(SrcODSFileName, DestODSFileName):
    """
    Функция тестирования чтения/записи ODS файла.
    """
    ods = icODS()
    data = ods.Load(SrcODSFileName)
    
    if data:
        ods.Save(DestODSFileName, data)


def test_2(SrcODSFileName, DestXMLFileName, DestODSFileName):
    """
    Функция тестирования чтения/записи ODS файла.
    """
    ods = icODS()
    data = ods.Load(SrcODSFileName)

    if data:
        import icexcel
        excel = icexcel.icVExcel()
        excel._data = data
        cell = excel.getWorkbook().getWorksheetIdx().getTable().getCell(40, 61)
        log.debug('Cell: %s : %s' % (cell.getAddress(), cell.getRegion()))
        cell.setValue('123456')

        excel.SaveAs(DestXMLFileName)
        ods.Save(DestODSFileName, data)

if __name__ == '__main__':
    test_complex('./testfiles/test.ods', './testfiles/result.ods')
