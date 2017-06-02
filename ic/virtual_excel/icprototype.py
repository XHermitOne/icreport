#!/usr/bin/env python
# -*- coding: utf-8 -*-

__version__ = (0, 0, 1, 2)

PROTOTYPE_ATTR_NAMES = ('name', 'children', 'crc', 'value')


class icVPrototype(object):
    """
    Прототип объектов Virtual Excel.
    """
    def __init__(self, parent=None, *args, **kwargs):
        """
        Конструктор.
        """
        self._parent = parent
        # Данные объекта
        self._attributes = {}

    def getApp(self):
        """
        Объект приложения.
        """
        if self._parent:
            return self._parent.getApp()
        return self

    def getData(self):
        """
        Данные.
        """
        if self._parent:
            return self._parent.getData()
        return None

    def get_attributes(self):
        """
        Данные об объекте.
        """
        return self._attributes

    def set_attributes(self, Data_={}):
        """
        Данные об объекте.
        """
        self._attributes = Data_
        return self._attributes

    def update_attributes(self, Data_={}):
        """
        Данные об объекте.
        """
        self._attributes.update(Data_)
        return self._attributes

    def create(self):
        """
        Создать.
        """
        attrs = self._parent.get_attributes()
        attrs['children'].append(self._attributes)
        return self._attributes

    def create_idx(self, idx):
        """
        Создать с индексом.
        """
        attrs = self._parent.get_attributes()
        attrs['children'].insert(idx, self._attributes)
        self._parent.set_attributes(attrs)
        return self._attributes

    def get_parent_by_name(self, Name_):
        """
        Поиск родительского объекта по имени.
        """
        if self._parent is None:
            return None
        elif 'name' in self._parent._attributes and self._parent._attributes['name'] == Name_:
            return self._parent
        else:
            return self._parent.get_parent_by_name(Name_)

    def clear(self):
        """
        Очистить.
        """
        if 'children' in self._attributes:
            self._attributes['children'] = []

    def copy(self):
        """
        Получить копию атрибутов объекта.
        """
        pass

    def paste(self, Paste_, To_=None):
        """
        Вставить копию атрибутов Past_ объекта внутрь текущего объекта
        по адресу To_. Если To_ None, тогда происходит замена.
        """
        pass

    def findChildAttrsByName(self, sName=None):
        """
        Поиск атрибутов дочернего элемента по имени.
        @param sName: Имя дочернего элемента.
        @return: Словарь атрибутов дочернего элемента или None, если не найден.
        """
        children = [child for child in self._attributes['children'] if child['name'] == sName]
        if children:
            return children[0]
        return None     
    
    def get_parent(self):
        return self._parent


class icVIndexedPrototype(icVPrototype):
    """
    Прототип индексируемого объекта.
    Необходим для реализации функций отслеживания и пересчета индексов.
    """
    def __init__(self, parent, *args, **kwargs):
        """
        Конструктор.
        """
        icVPrototype.__init__(self, parent, *args, **kwargs)

    def _maxElementIdx(self, ElementName_='', Elements_=None):
        """
        Максимальный индекс указанного элемента в родительском.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        cur_idx = -1
        if Elements_ is None:
            elements = [element for element in self._parent.get_attributes()['children']
                        if element['name'] == ElementName_]
        else:
            elements = Elements_
        if elements:
            for i, element_attr in enumerate(elements):
                if 'Index' in element_attr:
                    cur_idx = int(element_attr['Index'])-1
                else:
                    if 'Span' in element_attr:
                        # Несколько элементов с такими же атрибутами
                        cur_idx += int(element_attr['Span'])
                    else:
                        cur_idx += 1
        return cur_idx

    def _findElementIdxAttr(self, Idx_, ElementName_):
        """
        Найти атрибуты указанного элемента в родительском объекте по индексу.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        indexes = []
        cur_idx = 0
        ret_i = -1
        ret_attr = None
        flag = True
        for i, element_attr in enumerate(self._parent.get_attributes()['children']):
            if element_attr['name'] == ElementName_:
                if 'Index' in element_attr:
                    cur_idx = int(element_attr['Index'])
                else:
                    cur_idx += 1

                indexes.append(cur_idx)

                # Учет объединенных ячеек

                if Idx_ == cur_idx and flag:
                    ret_i = i
                    ret_attr = element_attr
                    flag = False
                elif Idx_ < cur_idx and flag:
                    ret_i = i
                    ret_attr = None
                    flag = False

        return indexes, ret_i, ret_attr

    def _reIndexElement(self, ElementName_, Element_, Index_, i_):
        """
        Переиндексирование элемента в родительском объекте.
        """
        if i_ > 0:
            # Предыдущие элементы
            prev_elements = [element for element in self._parent.get_attributes()['children'][:i_-1]
                             if element['name'] == ElementName_]
            if prev_elements:
                max_idx = self._maxElementIdx(ElementName_, prev_elements)
                if Index_ > (max_idx+1):
                    Element_.setIndex(Index_)
                return Element_
        if Index_ > 1:
            Element_.setIndex(Index_)
        return Element_

    def _reIndexAllElements(self, ElementNames_=(), OffsetIndex_=0):
        """
        Переиндексирование всех элементов в родительском объекте.
        """
        all_elements = []
        for i, element_attr in enumerate(self._parent.get_attributes()['children']):
            if element_attr['name'] in ElementNames_:
                if 'Index' in element_attr:
                    cur_idx = int(element_attr['Index'])
                else:
                    cur_idx += 1
                if 'Index' in element_attr:
                    element_attr['Index'] = cur_idx-OffsetIndex_
            all_elements.append(element_attr)
        return all_elements

    def getIndex(self):
        """
        Индекс индексируемого объекта в родительском объекте.
        """
        pass

    def setIndex(self, Index_):
        """
        Индекс объекта в родительском объекте.
        """
        self._attributes['Index'] = str(Index_)

    def _delElementIdxAttr(self, Idx_, ElementName_):
        """
        Удалить указанный элемент из родительском объекта по индексу.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        @return: Возвращает результат выполнения операции.
        """
        Idx_ += 1   # Проверка на совпадение индексов все равно делается в понятиях Excel т.е. начинается с 1
        cur_idx = 0
        children_count = len(self._parent.get_attributes()['children'])
        for i,element_attr in enumerate(self._parent.get_attributes()['children']):
            if element_attr['name'] == ElementName_:

                if 'Index' in element_attr:
                    cur_idx = int(element_attr['Index'])
                else:
                    cur_idx += 1

                if Idx_ == cur_idx:
                    del self._parent.get_attributes()['children'][i]
                    self._parent.get_attributes()['children'] = self._reIndexAfterDel(ElementName_, i)
                    return True

                elif Idx_ < cur_idx:
                    # Переиндексировать после удаления
                    self._reIndexAfterDel(ElementName_, i)
                    return True
        return False

    def _reIndexAfterDel(self, ElementName_, Index_):
        """
        Переиндексировать после удаления.
        """
        children = self._parent.get_attributes()['children']
        for element_attr in children[Index_:]:
            if element_attr['name'] == ElementName_:
                if 'Index' in element_attr:
                    element_attr['Index'] = int(element_attr['Index'])-1
        return children

    def _delElementIdxAttrChild(self, Idx_, ElementName_, IsReIndex_=True):
        """
        Удалить дочерний указанный элемент по индексу.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        @return: Возвращает результат выполнения операции.
        """
        Idx_ += 1   # Проверка на совпадение индексов все равно делается в понятиях Excel т.е. начинается с 1
        cur_idx = 0
        delta = 1
        children_count = len(self.get_attributes()['children'])
        for i, element_attr in enumerate(self.get_attributes()['children']):
            if element_attr['name'] == ElementName_:

                if 'Index' in element_attr:
                    cur_idx = int(element_attr['Index'])
                else:
                    cur_idx += 1

                if Idx_ == cur_idx:
                    element = self.get_attributes()['children'][i]
                    if 'MergeAcross' in element:
                        delta += int(element['MergeAcross'])

                    if not IsReIndex_:
                        self.get_attributes()['children'] = self._reIndexBeforeClearChild(ElementName_, i, delta)
                    del self.get_attributes()['children'][i]
                    if IsReIndex_:
                        self.get_attributes()['children'] = self._reIndexAfterDelChild(ElementName_, i, delta)
                    return True

                elif Idx_ < cur_idx:
                    # Переиндексировать после удаления
                    if IsReIndex_:
                        self._reIndexAfterDelChild(ElementName_, i, delta)
                    return True
        return False

    def _reIndexAfterDelChild(self, ElementName_, Index_, Delta_=1):
        """
        Переиндексировать дочерние элементы после удаления.
        """
        children = self.get_attributes()['children']
        for element_attr in children[Index_:]:
            if element_attr['name'] == ElementName_:
                if 'Index' in element_attr:
                    element_attr['Index'] = int(element_attr['Index'])-Delta_
        return children

    def _reIndexBeforeClearChild(self, ElementName_, Index_, Delta_=1):
        """
        Переиндексировать дочерние элементы до очистки элемента для объединения.
        """
        children = self.get_attributes()['children']
        for element_attr in children[Index_+1:]:
            if element_attr['name'] == ElementName_:
                if 'Index' not in element_attr:
                    element_attr['Index'] = self._maxElementIdx(ElementName_, children[:Index_+2])+Delta_
                break
        return children

    def _findElementIdxAttrChild(self, Idx_, ElementName_):
        """
        Найти атрибуты дочернего элемента по индексу.
        ВНИМАНИЕ! В этой функции индексация начинается с 0.
        """
        indexes = []
        cur_idx = 0
        ret_i = -1
        ret_attr = None
        flag = True
        for i, element_attr in enumerate(self.get_attributes()['children']):
            if element_attr['name'] == ElementName_:
                if 'Index' in element_attr:
                    cur_idx = int(element_attr['Index'])
                else:
                    cur_idx += 1

                # Учет объединенных ячеек

                indexes.append(cur_idx)

                if Idx_ == cur_idx and flag:
                    ret_i = i
                    ret_attr = element_attr
                    flag = False
                elif Idx_ < cur_idx and flag:
                    ret_i = i
                    ret_attr = None
                    flag = False

        return indexes, ret_attr
