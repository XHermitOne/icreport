#!/usr/bin/env python
# -*- coding: utf-8 -*-

import copy

import icprototype
import iccell

__version__ = (0, 0, 1, 2)

RANGE_ROW_IDX = 0
RANGE_COL_IDX = 1
RANGE_HEIGHT_IDX = 2
RANGE_WIDTH_IDX = 3


class icVRange(icprototype.icVPrototype):
    """
    Диапазон ячеек. Необходим для групповых операций над ячейками.
    """
    def __init__(self, parent, row=0, col=0, height=0, width=0, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)

        self.row = 1
        self.col = 1
        self.height = 1
        self.width = 1
        self._height_1 = 0
        self._width_1 = 0
        self._address = []
        self.setAddress(row, col, height, width)

        # Текущее смещение для функции относительной индексации ячеек
        self._cur_span_offset = (0, 0)

        # Базисная строка
        self._basis_row = None

    def setAddress(self, Row_, Col_, Height_, Width_):
        """
        Установить адрес.
        """
        self.row = max(Row_, 1)
        self.col = max(Col_, 1)
        self.height = max(Height_, 1)
        self.width = max(Width_, 1)
        self._height_1 = self.height-1
        self._width_1 = self.width-1
        self._address = [self.row, self.col, self.height, self.width]
        return self._address

    def setValues(self, Values_=None):
        """
        Установить значения в диапазоне.
        """
        for i_row, row in enumerate(Values_):
            if i_row < self.height:
                for i_col, col in enumerate(row):
                    if i_col < self.width:
                        cell = self._parent.getCell(self.row+i_row, self.col+i_col)
                        value = col
                        cell.setValue(value)

    def setStyle(self, alignment=None,
                 borders=None, font=None, interior=None,
                 number_format=None):
        """
        Установить стиль в диапазоне.
        """
        my_workbook = self.get_parent_by_name('Workbook')
        find_style = my_workbook.getStyles().findStyle(alignment, borders, font, interior, number_format)
        if find_style is None:
            style = my_workbook.getStyles().createStyle()
            style.setAttrs(alignment, borders, font, interior, number_format)
        else:
            style = find_style

        for i_row in range(self.height):
            for i_col in range(self.width):
                cell = self._parent.getCell(self.row+i_row, self.col+i_col)
                cell.setStyleID(style.getID())

    def updateStyle(self, alignment=None,
                    borders=None, font=None, interior=None,
                    number_format=None,
                    style_auto_create=True):
        """
        Установить стиль в диапазоне.
        """
        for i_row in range(self.height):
            for i_col in range(self.width):
                cell = self._parent.getCell(self.row+i_row, self.col+i_col)
                cell_style = cell.getStyle()
                # Если стиль ячейки не определен, то пропустить ее обработку
                if cell_style is None:
                    continue

                if not style_auto_create:
                    cell_style.updateAttrs(alignment, borders, font, interior, number_format)
                else:
                    cell_style_attrs = cell_style.getAttrs()
                    if alignment:
                        cell_style_attrs['alignment'] = alignment
                    if borders:
                        cell_style_attrs['borders'] = borders
                    if font:
                        cell_style_attrs['font'] = font
                    if interior:
                        cell_style_attrs['interior'] = interior
                    if number_format:
                        cell_style_attrs['number_format'] = number_format

                    my_workbook = self.get_parent_by_name('Workbook')
                    find_style = my_workbook.getStyles().findStyle(**cell_style_attrs)
                    if find_style is None:
                        style = my_workbook.getStyles().createStyle()
                        style.setAttrs(**cell_style_attrs)
                    else:
                        style = find_style

                    cell.setStyleID(style.getID())

    def _is_border_position(self, style, border_position):
        """
        Проверить есть ли в стиле в описании обрамления указанная граница.
        """
        return bool([border for border in style['children'] if border['Position'] == border_position])

    def _get_cell_border_attr_idx(self, old_borders, i_row, i_col,
                                  border_left=None, border_top=None, border_right=None, border_bottom=None):
        """
        Определить атрибуты стиля в диапазоне в зависимости
        от координат текущей ячейки.
        """
        if old_borders and 'children' in old_borders:
            attrs = {'borders': {'name': 'Borders', 'children': copy.deepcopy(old_borders['children'])}}
        else:
            attrs = {'borders': {'name': 'Borders', 'children': []}}

        if i_row == 0:
            # Ячейка верхней границы
            if border_top:
                if not self._is_border_position(attrs['borders'], 'Top'):
                    cur_border = border_top
                    cur_border['name'] = 'Border'
                    cur_border['Position'] = 'Top'
                    attrs['borders']['children'].append(cur_border)
            if i_col == 0:
                # Левая верхняя ячейка
                if border_left:
                    if not self._is_border_position(attrs['borders'], 'Left'):
                        cur_border = border_left
                        cur_border['name'] = 'Border'
                        cur_border['Position'] = 'Left'
                        attrs['borders']['children'].append(cur_border)
            if i_col == self._width_1:
                # Правая верхняя ячейка
                if border_right:
                    if not self._is_border_position(attrs['borders'], 'Right'):
                        cur_border = border_right
                        cur_border['name'] = 'Border'
                        cur_border['Position'] = 'Right'
                        attrs['borders']['children'].append(cur_border)

        if i_row == self._height_1:
            # Ячейка нижней границы
            if border_bottom:
                if not self._is_border_position(attrs['borders'], 'Bottom'):
                    cur_border = border_bottom
                    cur_border['name'] = 'Border'
                    cur_border['Position'] = 'Bottom'
                    attrs['borders']['children'].append(cur_border)
            if i_col == 0:
                # Левая нижняя ячейка
                if border_left:
                    if not self._is_border_position(attrs['borders'],'Left'):
                        cur_border = border_left
                        cur_border['name'] = 'Border'
                        cur_border['Position'] = 'Left'
                        attrs['borders']['children'].append(cur_border)
            if i_col == self._width_1:
                # Правая нижняя ячейка
                if border_right:
                    if not self._is_border_position(attrs['borders'], 'Right'):
                        cur_border = border_right
                        cur_border['name'] = 'Border'
                        cur_border['Position'] = 'Right'
                        attrs['borders']['children'].append(cur_border)

        if i_col == 0:
            # Ячейка левой границы
            if border_left:
                if not self._is_border_position(attrs['borders'], 'Left'):
                    cur_border = border_left
                    cur_border['name'] = 'Border'
                    cur_border['Position'] = 'Left'
                    attrs['borders']['children'].append(cur_border)
        if i_col == self._width_1:
            # Ячейка правой границы
            if border_right:
                if not self._is_border_position(attrs['borders'], 'Right'):
                    cur_border = border_right
                    cur_border['name'] = 'Border'
                    cur_border['Position'] = 'Right'
                    attrs['borders']['children'].append(cur_border)

        return attrs

    def setBorderOn(self, border_left=None,
                    border_top=None, border_right=None, border_bottom=None):
        """
        Обрамление диапазона ячеек.
        """
        my_workbook = self.get_parent_by_name('Workbook')

        for i_row in range(self.height):
            for i_col in range(self.width):
                # Обрабатывать только ячейки по периметру
                if (i_row == 0) or (i_col == 0) or \
                   (i_row == self._height_1) or (i_col == self._width_1):

                    # Ячейка
                    cell = self._parent.getCell(self.row+i_row, self.col+i_col)
                    # Идентификатор стиля ячейки
                    style_id = cell.getStyleID()
                    # Определение стиля
                    if style_id is not None:
                        style = my_workbook.getStyles().getStyle(style_id)
                        if style:
                            style_attrs = style.getAttrs()
                        else:
                            style_attrs = {'borders': {'name': 'Borders', 'children': []}}
                    else:
                        style_attrs = {'borders': {'name': 'Borders', 'children': []}}

                    if 'borders' not in style_attrs:
                        style_attrs['borders'] = {'name': 'Borders', 'children': []}

                    cur_style_borders = self._get_cell_border_attr_idx(style_attrs['borders'],
                                                                       i_row, i_col,
                                                                       border_left, border_top,
                                                                       border_right, border_bottom)
                    style_attrs['borders'] = cur_style_borders['borders']

                    # Установить стиль у ячейки
                    cell.setStyle(**style_attrs)
        return True

    def _limitOffset(self, RowOffset_, ColOffset_):
        """
        Ограничение смещения.
        """
        row_offset = max(min(RowOffset_, 0), self._height_1)
        col_offset = max(min(ColOffset_, 0), self._width_1)
        return row_offset, col_offset

    def spanColumn(self, Step_=1):
        """
        Взять ячейку в диапазоне по смещению колонки.
        """
        self._cur_span_offset = self._limitOffset(self._cur_span_offset[0],
                                                  self._cur_span_offset[1]+Step_)
        return self.getCellOffset(*self._cur_span_offset)

    def spanRow(self, Step_=1):
        """
        Взять ячейку в диапазоне по смещению строки.
        """
        self._cur_span_offset = self._limitOffset(self._cur_span_offset[0]+Step_,
                                                  self._cur_span_offset[1])
        return self.getCellOffset(*self._cur_span_offset)

    def getCellOffset(self, RowOffset_=0, ColOffset_=0):
        """
        Получить ячейку в диапазоне по относительным координатам диапазона.
        """
        # Ограничение смещения
        row_offset, col_offset = self._limitOffset(RowOffset_, ColOffset_)
        return self._parent.getCell(self.row+row_offset, self.col+col_offset)

    def _getBasisRow(self):
        """
        Базисная строка, относительно которой происходит работа с индексами строк.
        """
        if self._basis_row is None:
            self._basis_row = icVRow(self)
        return self._basis_row

    def copy(self):
        """
        Получить копию атрибутов объекта.
        """
        copy_result = {'name': 'Range',
                       'width': self.width,
                       'height': self.height,
                       'children': []}

        for i_row in range(self.height):
            cur_row = {'name': 'Row', 'children': []}
            copy_result['children'].append(cur_row)
            for i_col in range(self.width):
                cell = self._parent.getCell(self.row+i_row, self.col+i_col)
                cell_attrs = copy.deepcopy(cell.get_attributes())
                cur_row['children'].append(cell_attrs)
            # Переиндексировать все ячейки строки
            cell._reIndexAllElements(('Cell',))
        # Переиндексировать все строки диапазона ячеек
        self._getBasisRow()._reIndexAllElements(('Row',))

        return copy_result


class icVColumn(icprototype.icVIndexedPrototype, icVRange):
    """
    Колонка.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVIndexedPrototype.__init__(self, parent, *args, **kwargs)
        icVRange.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'Column', 'children': []}

    def setCaption(self, Caption_):
        """
        Заголовок колонки.
        """
        self._attributes['Caption'] = Caption_

    def setWidth(self, Width_):
        """
        Ширина колонки.
        """
        self._attributes['Width'] = str(Width_)

    def setCharacterWidth(self, CharacterCount_, Font_=None):
        """
        Установка ширины колонки по количеству символов в колонке.
        @param CharacterCount_: Количество символов в колонке.
        @param Font_: Указание шрифта, если не указано то
        берется Arial размера 10.
        """
        width = self._calcWidthByCharacter(CharacterCount_, Font_)
        self.setWidth(width)

    def _calcWidthByCharacter(self, CharacterCount_, Font_=None):
        """
        Функция преобразования количества символов в колонке в ширину.
        """
        return int(CharacterCount_*6)

    def setHidden(self, Hidden_=True):
        """
        Скрытие колонки.
        """
        if Hidden_:
            self._attributes['Hidden'] = str(int(Hidden_))
        else:
            if 'Hidden' in self._attributes:
                del self._attributes['Hidden']

    def setAutoFitWidth(self, bAutoFitWidth=True):
        """
        Установить авторазмер по ширине.
        @param bAutoFitWidth: Признак автообразмеривания колонки.
        """
        if bAutoFitWidth:
            self._attributes['AutoFitWidth'] = str(int(bAutoFitWidth))
        else:
            if 'AutoFitWidth' in self._attributes:
                del self._attributes['AutoFitWidth']


class icVRow(icprototype.icVIndexedPrototype, icVRange):
    """
    Строка.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVIndexedPrototype.__init__(self, parent, *args, **kwargs)
        icVRange.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'Row', 'children': []}

        # Базисная ячейка
        self._basis_cell = None

#     def setIndex(self,Index_):
#         """
#         Индекс строки в таблице.
#         """
#         self._attributes['Index']=str(Index_)

    def setHeight(self, Height_):
        """
        Высота строки.
        """
        self._attributes['Height'] = str(Height_)

    def setHidden(self, Hidden_=True):
        """
        Скрытие строки.
        """
        if Hidden_:
            self._attributes['Hidden'] = str(int(Hidden_))
        else:
            if 'Hidden' in self._attributes:
                del self._attributes['Hidden']

    def createCell(self):
        """
        Создать/Добавить в строку ячейку.
        """
        cell = iccell.icVCell(self)
        attrs = cell.create()
        return cell

    def insertCellIdx(self, Cell_, Idx_):
        """
        Вставить ячейку в строку по индексу.
        """
        indexes, cell_attr = self._findCellIdxAttr(Idx_)
        ins_i = max(0, len([i for i in indexes if i < Idx_]))

        if cell_attr is None:
            Cell_.setIndex(Idx_)
        self._attributes['children'].insert(ins_i, Cell_.get_attributes())
        return Cell_

    def createCellIdx(self, Idx_):
        """
        Создать/Добавить в строку ячейку.
        """
        cell = self.createCell()
        self._attributes['children'] = self._attributes['children'][:-1]

        # Переместить ячейку по индексу
        cell = self.insertCellIdx(cell, Idx_)
        return cell

    def _getBasisCell(self):
        """
        Базисная фчейка, относительно которой происходит работа с индексами ячеек.
        """
        if self._basis_cell is None:
            self._basis_cell = iccell.icVCell(self)
        return self._basis_cell

    def _findCellIdxAttr(self, Idx_):
        """
        Найти атрибуты ячеки в строке по индексу.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        return self._getBasisCell()._findElementIdxAttr(Idx_, 'Cell')

    def getCellIdx(self, Idx_):
        """
        Получить ячейку из строки по номеру.
        """
        indexes, cell_attr = self._findCellIdxAttr(Idx_)

        if cell_attr is None:
            cell = self.createCellIdx(Idx_)
        else:
            cell = iccell.icVCell(self)
            cell.set_attributes(cell_attr)
        return cell

    def delCell(self, Idx_):
        """
        Удалить ячейку из строки.
        """
        cell = self.getCellIdx(Idx_-1)
        if cell:
            # Удалить ячейку из строки
            return cell._delElementIdxAttr(Idx_-1, 'Cell')
        return False
