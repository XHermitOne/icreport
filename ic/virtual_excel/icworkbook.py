#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os.path

import icprototype

import icworksheet
import icstyle

__version__ = (0, 0, 1, 2)


class icVWorkbook(icprototype.icVPrototype):
    """
    Книга.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, parent, *args, **kwargs)
        self._attributes = {'name': 'Workbook', 'children': []}
        # Словарь листов по именам
        self._worksheet_dict = {}

        self.Name = None
        
        # Управление стилями
        self.styles = None

    def Load(self, Name_):
        """
        Загрузить.
        """
        self.Name=os.path.abspath(Name_.strip())
        return self._parent.Load(Name_)

    def Save(self):
        """
        Сохранить книгу.
        """
        return self._parent.Save()

    def SaveAs(self, Name_):
        """
        Сохранить как...
        """
        self.Name = os.path.abspath(Name_.strip())
        return self._parent.SaveAs(self.Name)
    
    def get_worksheet_dict(self):
        """
        Словарь листов по именам.
        """
        return self._worksheet_dict

    def init_worksheet_dict(self):
        """
        Инициализация словаря листов.
        """
        self._worksheet_dict = dict([(worksheet['Name'], worksheet) for worksheet in \
                                    [element for element in self._attributes['children'] if element['name'] == 'Worksheet']])
        return self._worksheet_dict
        
    worksheet_dict = property(get_worksheet_dict)

    def create(self):
        """
        Создать.
        """
        attrs = self._parent.get_attributes()
        # Т.к. в XML файле м.б. только одна книга, то здесь нельзя добавять
        # а только заменять описание книги
        attrs['children'] = [self._attributes]
        return self._attributes

    def createWorksheet(self):
        """
        Создать лист.
        """
        work_sheet = icworksheet.icVWorksheet(self)
        attrs = work_sheet.create()
        self.init_worksheet_dict()
        return work_sheet

    def _find_worksheet_attr_name(self, Worksheets_, Name_):
        """
        Найти лист в списке по имени.
        """
        work_sheet_attr = None

        if not isinstance(Name_, unicode):
            Name_ = unicode(Name_, 'utf-8')
            
        for sheet in Worksheets_:
            if not isinstance(sheet['Name'], unicode):
                name = unicode(sheet['Name'], 'utf-8')
            else:
                name = sheet['Name']
            if name == Name_:
                work_sheet_attr = sheet
                break
        return work_sheet_attr
            
    def findWorksheet(self, Name_):
        """
        Поиск листа по имени.
        """
        work_sheet = None
        if Name_ in self._worksheet_dict:
            work_sheet = icworksheet.icVWorksheet(self)
            work_sheet.set_attributes(self._worksheet_dict[Name_])
        else:
            # Попробовать поискать в списке
            find_worksheet = self._find_worksheet_attr_name([element for element in self._attributes['children'] \
                                                             if element['name'] == 'Worksheet'], Name_)
            if find_worksheet:
                work_sheet = icworksheet.icVWorksheet(self)
                work_sheet.set_attributes(find_worksheet)
                # Блин рассинхронизация произошла со словарем
                self.init_worksheet_dict()
        return work_sheet

    def getWorksheetIdx(self, Idx_=0):
        """
        Лист по индексу.
        """    
        work_sheets = [element for element in self._attributes['children'] if element['name'] == 'Worksheet']
        try:
            worksheet_attr = work_sheets[Idx_]
            work_sheet = icworksheet.icVWorksheet(self)
            work_sheet.set_attributes(worksheet_attr)
            return work_sheet
        except:
            print('Error getWorksheetIdx')
            raise
    
    def createStyles(self):
        """
        Создать стили.
        """
        styles = icstyle.icVStyles(self)
        attrs = styles.create()
        return styles

    def getStylesAttrs(self):
        """
        Стили.
        """
        styles = [element for element in self._attributes['children'] if element['name'] == 'Styles']
        if styles:
            return styles[0]
        return None

    def getStyles(self):
        """
        Объект стилей.
        """
        if self.styles:
            return self.styles
        
        styles_data = self.getStylesAttrs()
        if styles_data is None:
            self.styles = self.createStyles()
        else:
            self.styles = icstyle.icVStyles(self)
            self.styles.set_attributes(styles_data)
        return self.styles

    def getWorksheetNames(self):
        """
        Список имен листов книги.
        """
        return [work_sheet['Name'] for work_sheet in self._attributes['children'] if work_sheet['name'] == 'Worksheet']
