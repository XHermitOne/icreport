#!/usr/bin/env python
# -*- coding: utf-8 -*-

import copy
import sys

import icprototype
import icrange
import iccell
import paper_size
import config

try:
    # Если Virtual Excel работает в окружении icReport
    from ic.std import icexceptions
except ImportError:
    # Если Virtual Excel работает в окружении icServices
    from services.ic_std import icexceptions

__version__ = (0, 0, 1, 3)


class icVWorksheet(icprototype.icVPrototype):
    """
    Лист.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'Worksheet', 'Name': 'default',
                            'children': [{'name': 'WorksheetOptions',
                                          'children': [{'name': 'PageSetup',
                                                        'children': [{'name': 'PageMargins',
                                                                      'Bottom': 0.984251969, 'Left': 0.787401575,
                                                                      'Top': 0.984251969, 'Right': 0.787401575,
                                                                      'children': []}]}]}]}

        # Таблица листа
        # ВНИМАНИЕ! Кэшируется для увеличения производительности
        self._table = None

    def _is_worksheets_name(self, Worksheets_, Name_):
        """
        Существуют листы с именем Name_?
        """
        for sheet in Worksheets_:
            if not isinstance(sheet['Name'], unicode):
                name = unicode(sheet['Name'], 'utf-8')
            else:
                name = sheet['Name']
            if name == Name_:
                return True
        return False

    def _create_new_name(self):
        """
        Создать новое имя для листа.
        """
        work_sheets = [element for element in self._parent.get_attributes()['children'] if element['name'] == 'Worksheet']
        i = 1
        new_name = u'Лист%d' % i
        while self._is_worksheets_name(work_sheets, new_name):
            i += 1
            new_name = u'Лист%d' % i
        return new_name

    def create(self):
        """
        Создать.
        """
        attrs = self._parent.get_attributes()
        self._attributes['Name'] = self._create_new_name()
        attrs['children'].append(self._attributes)
        return self._attributes

    def getName(self):
        """
        Имя листа.
        """
        return self._attributes['Name']

    def setName(self, Name_):
        """
        Установить имя листа.
        """
        self._attributes['Name'] = Name_

    def createTable(self):
        """
        Создать таблицу.
        """
        self._table = icVTable(self)
        attrs = self._table.create()
        return self._table

    def getTable(self):
        """
        Таблица.
        """
        if self._table:
            return self._table

        tab_attr = [element for element in self._attributes['children'] if element['name'] == 'Table']

        if tab_attr:
            self._table = icVTable(self)
            self._table.set_attributes(tab_attr[0])
        else:
            self.createTable()
        return self._table

    def getCell(self, row, col):
        """
        Получить ячейку.
        """
        return self.getTable().getCell(row, col)

    def getRange(self, row, col, height, width):
        """
        Диапазон ячеек.
        """
        new_range = icrange.icVRange(self)
        new_range.setAddress(row, col, height, width)
        return new_range

    def getUsedRange(self):
        """
        Весь диапазон таблицы.
        """
        used_height, used_width = self.getTable().getUsedSize()
        return self.getRange(1, 1, used_height, used_width)

    def clearWorksheet(self):
        """
        Очистка листа.
        """
        return self.getTable().clearTab()

    def getWorksheetOptions(self):
        """
        Параметры листа.
        """
        options_attr = [element for element in self._attributes['children'] if element['name'] == 'WorksheetOptions']
        if options_attr:
            options = icVWorksheetOptions(self)
            options.set_attributes(options_attr[0])
        else:
            options = self.createWorksheetOptions()
        return options

    def createWorksheetOptions(self):
        """
        Параметры листа.
        """
        options = icVWorksheetOptions(self)
        attrs = options.create()
        return options

    def getPrintNumberofCopies(self):
        """
        Количество копий припечати листа.
        """
        options = self.getWorksheetOptions()
        if options:
            print_section = options.getPrint()
            if print_section:
                n_copies = print_section.getNumberofCopies()
                return n_copies
        return None

    def setPrintNumberofCopies(self, NumberofCopies_=1):
        """
        Количество копий припечати листа.
        """
        options = self.getWorksheetOptions()
        if options:
            print_section = options.getPrint()
            if print_section:
                return print_section.setNumberofCopies(NumberofCopies_)
        return False

    def clone(self, NewName_):
        """
        Создать клон листа и добавить его в книгу.
        param NewName_: Новое имя листа.
        """
        new_attributes = copy.deepcopy(self._attributes)
        new_attributes['Name'] = NewName_

        new_worksheet = self._parent.createWorksheet()
        new_worksheet.update_attributes(new_attributes)
        return new_worksheet

    def delColumn(self, Idx_=-1):
        """
        Удалить колонку.
        """
        return self.getTable().delColumn(Idx_)

    def delRow(self, Idx_=-1):
        """
        Удалить строку.
        """
        return self.getTable().delRow(Idx_)

    def getColumns(self, iStartIDX=0, iStopIDX=None):
        """
        Список объектов колонок.
        @param iStartIDX: Индекс первой колонки. По умолчанию с первой.
        @param iStopIDX: Индекс последней колонки. По умолчанию до последней.
        """
        return self.getTable().getColumns(iStartIDX, iStopIDX)

    def getPageBreaks(self):
        """
        Разрывы страниц.
        """
        page_breaks_attr = [element for element in self._attributes['children'] if element['name'] == 'PageBreaks']
        if page_breaks_attr:
            page_breaks = icVPageBreaks(self)
            page_breaks.set_attributes(page_breaks_attr[0])
        else:
            page_breaks = self.createPageBreaks()
        return page_breaks

    def createPageBreaks(self):
        """
        Разрывы страниц.
        """
        page_breaks = icVPageBreaks(self)
        attrs = page_breaks.create()
        return page_breaks


class icVTable(icprototype.icVPrototype):
    """
    Таблица.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'Table', 'children': []}

        # Базисные строка и колонка
        self._basis_row = None
        self._basis_col = None

        # Словарь объединенных ячеек
        self._merge_cells = None

    def getUsedSize(self):
        """
        Используемый размер таблицы.
        """
        n_cols = self._maxColIdx()+1
        n_rows = self._maxRowIdx()+1
        return n_rows, n_cols

    def createColumn(self):
        """
        Создать колонку.
        """
        col = icrange.icVColumn(self)
        attrs = col.create()
        return col

    def getColumns(self, iStartIDX=0, iStopIDX=None):
        """
        Список объектов колонок.
        @param iStartIDX: Индекс первой колонки. По умолчанию с первой.
        @param iStopIDX: Индекс последней колонки. По умолчанию до последней.
        """
        col_count = self.getColumnCount()
        if iStopIDX is None:
            iStopIDX = col_count
        # Защита от не корректных входных данных
        if iStartIDX > iStopIDX:
            iStartIDX = iStopIDX
        return [self.getColumn(idx) for idx in range(iStartIDX, iStopIDX)]

    def getColumnsAttrs(self):
        """
        Список колонок. Данные.
        """
        return [element for element in self._attributes['children'] if element['name'] == 'Column']

    def getColumnCount(self):
        """
        Количество колонок.
        """
        return self._maxColIdx()+1

    def _reIndexCol(self, Col_, Index_, i_):
        """
        Переиндексирование колонки в таблице.
        """
        return self._getBasisCol()._reIndexElement('Column', Col_, Index_, i_)

    def _createColIdx(self, Idx_, i_):
        """
        Создать колонку с индексом.
        @param Idx_: Индекс Excel.
        @param i_: Индекс в списке children.
        """
        col = icrange.icVColumn(self)
        idx = 0
        for i, child in enumerate(self._attributes['children']):
            if child['name'] == 'Column':
                if idx >= i_:
                    attrs = col.create_idx(idx)

                    self._reIndexCol(col, Idx_, idx)
                    break
                idx += 1

        return col

    def getColumn(self, Idx_=-1):
        """
        Взять колонку по индексу.
        """
        col = None
        idxs, _i, col_data = self._findColIdxAttr(Idx_)
        if col_data is not None:
            col = icrange.icVColumn(self)
            col.set_attributes(col_data)
        elif _i >= 0 and col_data is None:
            for i in idxs:
                if Idx_ <= i:
                    return self._createColIdx(Idx_, _i)
            return self.createColumn()
        return col

    def createRow(self):
        """
        Создать строку.
        """
        row = icrange.icVRow(self)
        attrs = row.create()
        return row

    def cloneRow(self, bClearCell=True, iRow=-1):
        """
        Клонировать строку таблицы.
        @param bClearCell: Очистить значения в ячейках.
        @param iRow: Индекс(Начинаяется с 0) клонируемой ячейки. -1 - Последняя.
        @return: Возвращает объект клонированной строки. Если строк в таблице нет, то возвращает None.
        """
        if self._attributes['children']:
            row_attr = copy.deepcopy(self._attributes['children'][iRow])
            if bClearCell:
                row_attr['children'] = [dict(cell.items()+[('value', None)]) for cell in row_attr['children']]

            row = icrange.icVRow(self)
            row.set_attributes(row_attr)
            return row
        return None

    def _reIndexRow(self, Row_, Index_, i_):
        """
        Переиндексирование строки в таблице.
        """
        return self._getBasisRow()._reIndexElement('Row', Row_, Index_, i_)

    def _createRowIdx(self, Idx_, i_):
        """
        Создать строку с индексом.
        """
        row = icrange.icVRow(self)
        for i,child in enumerate(self._attributes['children']):
            if child['name'] == 'Row':
                if i >= i_:
                    attrs = row.create_idx(i)

                    self._reIndexRow(row, Idx_, i)
                    break

        return row

    def getRowsAttrs(self):
        """
        Список строк. Данные.
        """
        return [element for element in self._attributes['children'] if element['name'] == 'Row']

    def getRowCount(self):
        """
        Количество строк.
        """
        return self._maxRowIdx()+1

    def getRow(self, Idx_=-1):
        """
        Взять строку по индексу.
        """
        row = None
        idxs, _i, row_data = self._findRowIdxAttr(Idx_)
        if row_data is not None:
            row = icrange.icVRow(self)
            row.set_attributes(row_data)
        else:
            for i in idxs:
                if Idx_ <= i:
                    return self._createRowIdx(Idx_, _i)
            return self.createRow()
        return row

    def createCell(self, row, col):
        """
        Создать ячейку (row,col).
        """
        col_count = self.getColumnCount()
        if col > col_count:
            for i in range(col-col_count):
                self.createColumn()

        row_count = self.getRowCount()
        if row > row_count:
            for i in range(row-row_count):
                self.createRow()

        # Проверка на попадание в объединенную ячейку
        if self.isInMergeCell(row, col):
            sheet_name = self.get_parent_by_name('Worksheet').getName()
            err_txt = 'Getting cell (sheet: %s, row: %d, column: %d) into merge cell!' % (sheet_name, row, col)
            raise icexceptions.icMergeCellError((100, err_txt))

        cur_row = self.getRow(row)
        cell = cur_row.createCellIdx(col)
        return cell

    def getCell(self, row, col):
        """
        Получить ячейку (row,col).
        """
        # Если координаты недопустимы, тогда ошибка
        if row <= 0:
            raise IndexError
        if col <= 0:
            raise IndexError

        # Ограничение по индексам строк и колонок
        if row > 65535:
            return None
        if col > 256:
            return None

        col_count = self.getColumnCount()
        if col > col_count:
            for i in range(col-col_count):
                self.createColumn()

        row_count = self.getRowCount()
        if row > row_count:
            for i in range(row-row_count):
                self.createRow()

        # Проверка на попадание в объединенную ячейку
        if self.isInMergeCell(row, col):
            if config.DETECT_MERGE_CELL_ERROR:
                sheet_name = self.get_parent_by_name('Worksheet').getName()
                err_txt = 'Getting cell (sheet: %s, row: %d, column: %d) into merge cell!' % (sheet_name, row, col)
                raise icexceptions.icMergeCellError((100, err_txt))
            else:
                cell = self.getInMergeCell(row, col)
                return cell

        cur_row = self.getRow(row)
        cell = cur_row.getCellIdx(col)
        # Установить координаты ячейки
        cell._row_idx = row
        cell._col_idx = col
        return cell

    def clearTab(self):
        """
        Очистка таблицы.
        """
        return self.clear()

    def _findColIdxAttr(self, Idx_):
        """
        Найти атрибуты колонки в таблице по индексу.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        return self._getBasisCol()._findElementIdxAttr(Idx_, 'Column')

    def _findRowIdxAttr(self, Idx_):
        """
        Найти атрибуты строки в таблице по индексу.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        return self._getBasisRow()._findElementIdxAttr(Idx_, 'Row')

    def _maxColIdx(self):
        """
        Максимальный индекс колонок в таблице.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        return self._getBasisCol()._maxElementIdx(Elements_=self.getColumnsAttrs())

    def _maxRowIdx(self):
        """
        Максимальный индекс строк в таблице.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        return self._getBasisRow()._maxElementIdx(Elements_=self.getRowsAttrs())

    def setExpandedRowCount(self, ExpandedRowCount_=None):
        """
        Вычисление максимального количества строк таблицы.
        """
        if ExpandedRowCount_:
            self._attributes['ExpandedRowCount'] = ExpandedRowCount_
        else:
            if 'ExpandedRowCount' in self._attributes:
                cur_count = int(self._attributes['ExpandedRowCount'])
                calc_count = self._maxRowIdx()+1
                # Если расчетное количество больше текущего, то
                # генератор добавил строки и значение ExpandedRowCount
                # надо увеличить
                # Ограничение количества строк 65535
                self._attributes['ExpandedRowCount'] = min(max(calc_count, cur_count), 65535)

    def setExpandedColCount(self, ExpandedColCount_=None):
        """
        Вычисление максимального количества колонок в строке.
        """
        if ExpandedColCount_:
            self._attributes['ExpandedColumnCount'] = ExpandedColCount_
        else:
            if 'ExpandedColumnCount' in self._attributes:
                cur_count = int(self._attributes['ExpandedColumnCount'])
                calc_count = self._maxColIdx()+1
                # Если расчетное количество больше текущего, то
                # генератор добавил колонки и значение ExpandedColumnCount
                # надо увеличить
                # Ограничение количества колонок 256
                self._attributes['ExpandedColumnCount'] = min(max(calc_count, cur_count), 256)

    def paste(self, Paste_, To_=None):
        """
        Вставить копию атрибутов Past_ объекта внутрь текущего объекта
        по адресу To_. Если To_ None, тогда происходит замена.
        """
        if Paste_['name'] == 'Range':
            return self._pasteRange(Paste_, To_)
        else:
            print('ERROR: Error paste object attributes %s' % Paste_)
        return False

    def _pasteRange(self, Paste_, To_):
        """
        Вставить Range в таблицу по адресу ячейки.
        """
        if isinstance(To_, tuple) and len(To_) == 2:
            to_row, to_col = To_
            # Адресация ячеек задается как (row,col)
            for i_row in range(Paste_['height']):
                for i_col in range(Paste_['width']):
                    cell_attrs = Paste_['children'][i_row][i_col]
                    cell = self.getCell(to_row+i_row, to_col+i_col)
                    cell.set_attributes(cell_attrs)
            return True
        else:
            print('ERROR: Paste address error %s' % To_)
        return False

    def _getBasisRow(self):
        """
        Базисная строка, относительно которой происходит работа с индексами строк.
        """
        if self._basis_row is None:
            self._basis_row = icrange.icVRow(self)
        return self._basis_row

    def _getBasisCol(self):
        """
        Базисная колонка, относительно которой происходит работа с индексами колонок.
        """
        if self._basis_col is None:
            self._basis_col = icrange.icVColumn(self)
        return self._basis_col

    def getMergeCells(self):
        """
        Словарь объединенных ячеек. В качестве ключа - кортеж координаты ячейки.
        """
        merge_cells = {}
        rows = [element for element in self._attributes['children'] if element['name'] == 'Row']
        for i_row, row in enumerate(rows):
            i_col = 0
            for cell in row['children']:
                if cell['name'] == 'Cell':
                    if 'Index' in cell:
                        new_i_col = int(cell['Index'])
                        if new_i_col >= i_col:
                            i_col = new_i_col
                    else:
                        i_col += 1

                    if 'MergeAcross' in cell or 'MergeDown' in cell:
                        cur_row = self.getRow(i_row+1)
                        cell_obj = cur_row.getCellIdx(i_col)
                        # Установить координаты ячейки
                        cell_obj._row_idx = i_row+1
                        cell_obj._col_idx = i_col
                        merge_cells[cell_obj.getRegion()] = cell_obj
                    if 'MergeAcross' in cell:
                        # Учет объекдиненных ячеек ДЕЛАТЬ ОБЯЗАТЕЛЬНО!!!
                        # иначе не происходит учет предыдущих объединенных ячеек
                        i_col += int(cell['MergeAcross'])-1

        return merge_cells

    def isInMergeCell(self, Row_, Col_):
        """
        Попадает указанная ячейка в объединенную?
        """
        # Кеширование объединенных ячеек на случай попадания
        # в них при создании новой ячейки
        if self._merge_cells is None:
            self._merge_cells = self.getMergeCells()

        for cell in self._merge_cells.items():
            cell_region = cell[0]
            if (Row_ >= cell_region[0]) and (Row_ <= (cell_region[0]+cell_region[2])) and \
                    (Col_ >= cell_region[1]) and (Col_ <= (cell_region[1]+cell_region[3])):
                if Row_ != cell_region[0] or Col_ != cell_region[1]:
                    return True
        return False

    def getInMergeCell(self, Row_, Col_):
        """
        Получить объединенную ячейку на которую указывают координаты.
        """
        # Кеширование объединенных ячеек на случай попадания
        # в них при создании новой ячейки
        if self._merge_cells is None:
            self._merge_cells = self.getMergeCells()

        for cell in self._merge_cells.items():
            cell_region = cell[0]
            if (Row_ >= cell_region[0]) and (Row_<= (cell_region[0]+cell_region[2])) and \
                    (Col_ >= cell_region[1]) and (Col_ <= (cell_region[1]+cell_region[3])):
                if Row_ != cell_region[0] or Col_ != cell_region[1]:
                    return cell[1]
        return None

    def delColumn(self, Idx_=-1):
        """
        Удалить колонку.
        """
        col = self.getColumn(Idx_)
        if col:
            # Удалить колонку из таблицы
            result = col._delElementIdxAttr(Idx_-1, 'Column')
            # Кроме этого удалить ячейку, соответствующую текущей колонке
            for i_row in range(self.getRowCount()):
                row = self.getRow(i_row+1)
                if row:
                    row.delCell(Idx_)
            return result
        return False

    def delRow(self, Idx_=-1):
        """
        Удалить строку.
        """
        row = self.getRow(Idx_)

        if row:
            # Удалить строку из таблицы
            return row._delElementIdxAttr(Idx_-1, 'Row')
        return False


class icVWorksheetOptions(icprototype.icVPrototype):
    """
    Параметры листа.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'WorksheetOptions', 'children': []}

    def getPageSetup(self):
        """
        Параметры печати.
        """
        page_setup_attr = [element for element in self._attributes['children'] if element['name'] == 'PageSetup']
        if page_setup_attr:
            page_setup = icVPageSetup(self)
            page_setup.set_attributes(page_setup_attr[0])
        else:
            page_setup = self.createPageSetup()
        return page_setup

    def createPageSetup(self):
        """
        Параметры печати.
        """
        page_setup = icVPageSetup(self)
        attrs = page_setup.create()
        return page_setup

    def getPrint(self):
        """
        Параметры принтера.
        """
        print_attr = [element for element in self._attributes['children'] if element['name'] == 'Print']
        if print_attr:
            print_setup = icVPrint(self)
            print_setup.set_attributes(print_attr[0])
        else:
            print_setup = self.createPrint()
        return print_setup

    def createPrint(self):
        """
        Параметры принтера.
        """
        print_section = icVPrint(self)
        attrs = print_section.create()
        return print_section

    def isFitToPage(self):
        """
        Масштаб по размещению страниц.
        """
        fit_to_page = [element for element in self._attributes['children'] if element['name'] == 'FitToPage']
        return bool(fit_to_page)


class icVPageSetup(icprototype.icVPrototype):
    """
    Параметры печати.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'PageSetup', 'children': []}

    def getLayout(self):
        """
        Размещение листа.
        """
        layout = [element for element in self._attributes['children'] if element['name'] == 'Layout']
        if layout:
            return layout[0]
        return None

    def getOrientation(self):
        """
        Ориентация листа.
        """
        layout = self.getLayout()
        if layout:
            if 'Orientation' in layout:
                return layout['Orientation']
        # По умолчанию портретная ориентация
        return 'Portrait'

    def getCenter(self):
        """
        Центрирование по горизонтали/вертикали.
        """
        layout = self.getLayout()
        if layout:
            c_horiz = '0'
            if 'CenterHorizontal' in layout:
                c_horiz = layout['CenterHorizontal']
            c_vert = '0'
            if 'CenterVertical' in layout:
                c_vert = layout['CenterVertical']
            return bool(c_horiz == '1'), bool(c_vert == '1')

        # По умолчанию портретная ориентация
        return False, False

    def getPageMargins(self):
        """
        Поля.
        """
        margins = [element for element in self._attributes['children'] if element['name'] == 'PageMargins']
        if margins:
            return margins[0]
        return {}

    def getMargins(self):
        """
        Поля.
        """
        margins = self.getPageMargins()
        if margins:
            left_margin = 0
            if 'Left' in margins:
                left_margin = float(margins['Left'])

            top_margin = 0
            if 'Top' in margins:
                top_margin = float(margins['Top'])

            right_margin = 0
            if 'Right' in margins:
                right_margin = float(margins['Right'])

            bottom_margin = 0
            if 'Bottom' in margins:
                bottom_margin = float(margins['Bottom'])

            return left_margin, top_margin, right_margin, bottom_margin

        return 0, 0, 0, 0


class icVPrint(icprototype.icVPrototype):
    """
    Параметры принтера.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'Print', 'children': []}

    def getPaperSizeIndex(self):
        """
        Код размера бумаги.
        """
        paper_size_lst = [element for element in self._attributes['children'] if element['name'] == 'PaperSizeIndex']
        if paper_size_lst:
            return int(paper_size_lst[0]['value'])
        # По умолчанию размер A4
        return paper_size.xlPaperA4

    def getPaperSize(self):
        """
        Размер бумаги в 0.01 мм.
        """
        paper_size_i = self.getPaperSizeIndex()
        if paper_size_i > 0:
            return paper_size.XL_PAPER_SIZE.setdefault(paper_size_i, None)
        return None

    def getScale(self):
        """
        Масштаб бумаги.
        """
        scale = [element for element in self._attributes['children'] if element['name'] == 'Scale']
        if scale:
            return int(scale[0]['value'])
        # По умолчанию масштаб 100%
        return 100

    def getFitWidth(self):
        """
        Масштаб. Разместить не более чем на X стр. в ширину.
        """
        fit_width = [element for element in self._attributes['children'] if element['name'] == 'FitWidth']
        if fit_width:
            try:
                return int(fit_width[0]['value'])
            except:
                pass
        # По умолчанию 1
        return 1

    def getFitHeight(self):
        """
        Масштаб. Разместить не более чем на X стр. в высоту.
        """
        fit_height = [element for element in self._attributes['children'] if element['name'] == 'FitHeight']
        if fit_height:
            try:
                return int(fit_height[0]['value'])
            except:
                pass
        # По умолчанию 1
        return 1

    def getFit(self):
        """
        Масштаб в размещенных страницах.
        """
        return self.getFitWidth(), self.getFitHeight()

    def getNumberofCopies(self):
        """
        Количество копий листа.
        """
        n_copies = [element for element in self._attributes['children'] if element['name'] == 'NumberofCopies']
        if n_copies:
            try:
                return int(n_copies[0]['value'])
            except:
                pass
        # По умолчанию 1
        return 1

    def setNumberofCopies(self, NumberofCopies_=1):
        """
        Количество копий листа.
        """
        number_of_copies = min(max(int(NumberofCopies_), 1), 256)
        n_copies = {'name': 'NumberofCopies', 'value': number_of_copies}
        self._attributes['children'].append(n_copies)
        return n_copies


class icVPageBreaks(icprototype.icVPrototype):
    """
    Разрывы страниц.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'PageBreaks', 'children': [{'name': 'RowBreaks', 'children': []}]}

    def addRowBreak(self, iRow):
        """
        Добавить разрыв страницы по строке.
        @param iRow: Номер строки.
        """
        row_break = {'name': 'RowBreak', 'children': [{'name': 'Row', 'value': iRow}]}
        self._attributes['children'][0]['children'].append(row_break)
