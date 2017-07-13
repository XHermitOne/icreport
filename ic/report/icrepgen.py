#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль генератора отчетов.

В ячейках шаблона отчета можон ставить следующие теги:
["..."] - Обращение к полю таблицы запроса.

[&...&] - Обращение к переменной отчета.
Переменные отчета могут задаваться в таблице запроса
в виде словаря по ключу '__variables__'.

[@package.module.function()@] - Вызов функции прикладного программиста.

[=...=] - Исполнение блока кода.
Системные переменные блока кода:
    value - Значение записываемое в ячейку
    record - Словарь текущей записи таблицы запроса
Например:
    [=value=record['dt'].strftime('%B')=]
    [=cell['color']=dict(background=(128, 0 , 0)) if record['is_alarm'] else None; value=record['field_name']=]

[^...^] - Системные функции генератора.
Например: 
    [^N^] - Номер строки табличной части.
    [^SUM(record['...'])^] или [^SUM({имя поля})^] - Суммирование по полю.
    [^AVG(record['...'])^] или [^AVG({имя поля})^] - Вычисление среднего значения по полю.

[*...*] - Установка стиля генерации.
"""

# Подключение библиотек
import time
import re
import copy

from ic.std.log import log
from ic.std.utils import textfunc

__version__ = (0, 0, 1, 3)

# Константы
# Ключевые теги для обозначения:
# значения поля таблицы запроса
REP_FIELD_PATT = r'(\[\'.*?\'\])'
# функционала
REP_FUNC_PATT = r'(\[@.*?@\])'
# ламбда-выражения
REP_LAMBDA_PATT = r'(\[~.*?~\])'
# переменной
REP_VAR_PATT = r'(\[&.*?&\])'
# Блок кода
REP_EXEC_PATT = r'(\[=.*?=\])'
# Системные функции
REP_SYS_PATT = r'(\[\^.*?\^\])'
REP_SUM_FIELD_START = '{'   # Теги используются в системной функции
REP_SUM_FIELD_STOP = '}'    # суммирования SUM для обозначения значений полей
# Указание стиля из библиотеки стилей
REP_STYLE_PATT = r'(\[\*.*?\*\])'
# Указание родительского отчета
REP_SUBREPORT_PATT = r'(\[$.*?$\])'

# Список всех патернов используемых при разборе значений ячеек
ALL_PATTERNS = [REP_FIELD_PATT,
                REP_FUNC_PATT,
                REP_LAMBDA_PATT,
                REP_VAR_PATT,
                REP_EXEC_PATT,
                REP_SYS_PATT,
                REP_STYLE_PATT,
                REP_SUBREPORT_PATT,
                ]
    
# Спецификации и структуры
# Структура шаблона отчета
# Следующие ключи необходимы только для ICReportGenerator'a
IC_REP_TMPL = {'name': '',              # Имя отчета
               'description': '',       # Описание шаблона
               'variables': {},         # Переменные отчета
               'generator': None,       # Генератор
               'data_source': None,     # Указание источника данных/БД
               'query': None,           # Запрос отчета
               'style_lib': None,       # Библиотека стилей
               'header': {},            # Бэнд заголовка отчета (Координаты и размер)
               'footer': {},            # Бэнд подвала/примечания отчета (Координаты и размер)
               'detail': {},            # Бэнд области данных (Координаты и размер)
               'groups': [],            # Список бэндов групп (Координаты и размер)
               'upper': {},             # Бэнд верхнего колонтитула (Координаты и размер)
               'under': {},             # Бэнд нижнего колонтитула (Координаты и размер)
               'sheet': [],             # Лист ячеек отчета (Список строк описаний ячеек)
               'args': {},              # Аргументы для вывода отчета в акцесс
               'page_setup': None,      # Параметры страницы
               }

# Ориентация страницы
IC_REP_ORIENTATION_PORTRAIT = 0     # Книжная
IC_REP_ORIENTATION_LANDSCAPE = 1    # Альбомная

# Параметры страницы
IC_REP_PAGESETUP = {'orientation': IC_REP_ORIENTATION_PORTRAIT,     # Ориентация страницы
                    'start_num': 1,                                 # Начинать нумеровать со страницы...
                    'page_margins': (0, 0, 0, 0),                   # Поля
                    'scale': 100,                                   # Масштаб печати (в %)
                    'paper_size': 9,                                # Размер страницы - 9-A4
                    'resolution': (600, 600),                       # Плотность/качество печати
                    'fit': (1, 1),                                  # Параметры заполнения отчета на листах
                    }

# Структура данных бэнда
IC_REP_BAND = {'row': -1,       # Строка бэнда
               'col': -1,       # Колонка бэнда
               'row_size': -1,  # Размер бэнда по строкам
               'col_size': -1,  # Размер бэнда по колонкам
               }

# Форматы ячеек
REP_FMT_NONE = None     # Не устанавливать формат
REP_FMT_STR = 'S'       # Строковый/текстовый
REP_FMT_TIME = 'T'      # Время
REP_FMT_DATE = 'D'      # Дата
REP_FMT_NUM = 'N'       # Числовой
REP_FMT_FLOAT = 'F'     # Числовой с плавающей точкой
REP_FMT_MISC = 'M'      # Просто какойто формат
REP_FMT_EXCEL = 'X'     # Формат заданный Excel

# Структура ячейки отчета
IC_REP_CELL = {'merge_row': 0,      # Кол-во строк для объединеных ячеек
               'merge_col': 0,      # Кол-во волонок для объединеных ячеек
               'left': 0,           # Координата X
               'top': 0,            # Координата Y
               'width': 10,         # Ширина ячейки
               'height': 10,        # Высота ячейки
               'value': None,       # Текст ячейки
               'font': None,        # Шрифт Структура типа ic.components.icfont.SPC_IC_FONT
               'color': None,       # Цвет
               'border': None,      # Обрамление
               'align': None,       # Расположение текста
               'sum': None,         # Список сумм
               'visible': True,     # Видимость ячейки
               'format': None,      # Формат ячейки
               }

# Данные сумм
# ВНИМАНИЕ!!! На каждой итерации текущее значение суммы вычисляется, как value=value+eval(formul)
IC_REP_SUM = {'value': 0,       # Текущее значение суммы
              'formul': '0',    # Формула вычисления сумм
              }

# Цвет
IC_REP_COLOR = {'text': (0, 0, 0),      # Цвет текста
                'background': None,     # Цвет фона
                }

# Обрамление - кортеж из 4-х элементов
IC_REP_BORDER_LEFT = 0
IC_REP_BORDER_TOP = 1
IC_REP_BORDER_BOTTOM = 2
IC_REP_BORDER_RIGHT = 3
IC_REP_BORDER_LINE = {'color': (0, 0, 0),   # Цвет
                      'style': None,        # Стиль
                      'weight': 0,          # Толщина
                      }

# Стили линий обрамления отчета
IC_REP_LINE_SOLID = 0
IC_REP_LINE_SHORT_DASH = 1
IC_REP_LINE_DOT_DASH = 2
IC_REP_LINE_DOT = 3
IC_REP_LINE_TRANSPARENT = None

# Размещение
IC_REP_ALIGN = {'align_txt': (0, 0),    # кортеж из 2-х элементов
                'wrap_txt': False,      # Перенос текста по словам
                }
IC_REP_ALIGN_HORIZ = 0
IC_REP_ALIGN_VERT = 1

# Выравнивание текста
IC_HORIZ_ALIGN_LEFT = 0
IC_HORIZ_ALIGN_CENTRE = 1
IC_HORIZ_ALIGN_RIGHT = 2

IC_VERT_ALIGN_TOP = 3
IC_VERT_ALIGN_CENTRE = 4
IC_VERT_ALIGN_BOTTOM = 5

# Структура группы
IC_REP_GRP = {'header': {},         # Заголовок группы.
              'footer': {},         # Примечание группы.
              'field': None,        # Имя поля группы
              'old_rec': None,      # Старое значение записи таблицы запроса.
              }

DEFAULT_ENCODING = 'utf-8'


class icReportGenerator:
    """
    Класс генератора отчета.
    """
    def __init__(self):
        """
        Конструктор класса.
        """
        # Имя отчета
        self._RepName = None
        # Таблица запроса
        self._QueryTbl = None
        # Количество записей таблицы запроса
        self._QueryTblRecCount = -1
        # Текущая запись таблицы запроса
        self._CurRec = {}
        # Шаблон отчета
        self._Template = None
        # Описание листа шаблона отчета
        self._TemplateSheet = None
        # Выходной отчет
        self._Rep = None

        # Список групп
        self._RepGrp = []

        # Текущая координата Y для перераспределения координат ячеек
        self._cur_top = 0

        # Пространство имен отчета
        self._NameSpace = {}

        # Атрибуты ячейки по умолчанию
        # если None, то атрибуты не устанавливаются
        self.AttrDefault = None

        # Библиотека стилей
        self._StyleLib = None
        
        # Покоординатная замена значений ячеек
        self._CoordFill = None

        # Словарь форматов ячеек
        self._cellFmt = {}

    def generate(self, RepTemplate_, QueryTable_, NameSpace_=None, CoordFill_=None):
        """
        Генерация отчета.
        @param RepTemplate_: Структура шаблона отчета (см. спецификации).
        @param QueryTable_: Таблица запроса. 
            Словарь следующей структуры:
                {
                    '__name__':имя таблицы запроса,
                    '__fields__':[имена полей],
                    '__data__':[список списков значений],
                    '__sub__':{словарь данных подотчетов},
                }.
        @param NameSpace_: Пространство имен шаблона.
            Обычный словарь:
                {
                    'имя переменной': значение переменной, 
                }.
            ВНИМАНИЕ! Этот словарь может передаваться в таблице запроса
                ключ __variables__.
        @param CoordFill_: Координатное заполнение значений ячеек.
            Формат:
                {
                    (Row,Col): 'Значение',
                }.
            ВНИМАНИЕ! Этот словарь может передаваться в таблице запроса
                ключ __coord_fill__.
        @return: Заполненную структуру отчета.
        """
        try:
            # Покоординатная замена значений ячеек
            self._CoordFill = CoordFill_
            if QueryTable_ and '__coord_fill__' in QueryTable_:
                if self._CoordFill is None:
                    self._CoordFill = dict()
                self._CoordFill.update(QueryTable_['__coord_fill__'])

            # Инициализация списка групп
            self._RepGrp = list()

            # I. Определить все бэнды в шаблоне и ячейки сумм
            if isinstance(RepTemplate_, dict):
                self._Template = RepTemplate_
            else:
                # Вывести сообщение об ошибке в лог
                log.warning(u'Ошибка типа шаблона отчета <%s>.' % type(RepTemplate_))
                return None

            # Инициализация имени отчета
            if 'name' in QueryTable_ and QueryTable_['name']:
                # Если таблица запроса именована, то значит это имя готового отчета
                self._RepName = str(QueryTable_['name'])
            elif 'name' in self._Template:
                self._RepName = self._Template['name']
            
            # Заполнить пространство имен
            self._NameSpace = NameSpace_
            if self._NameSpace is None:
                self._NameSpace = dict()
            self._NameSpace.update(self._Template['variables'])
            if QueryTable_ and '__variables__' in QueryTable_:
                self._NameSpace.update(QueryTable_['__variables__'])
            log.debug(u'Переменные отчета: %s' % self._NameSpace.keys())

            # Библиотека стилей
            self._StyleLib = None
            if 'style_lib' in self._Template:
                self._StyleLib = self._Template['style_lib']
            
            self._TemplateSheet = self._Template['sheet']
            self._TemplateSheet = self._initSumCells(self._TemplateSheet)

            # II. Инициализация таблицы запроса
            self._QueryTbl = QueryTable_
            # Определить количество записей в таблице запроса
            self._QueryTblRecCount = 0
            if self._QueryTbl and '__data__' in self._QueryTbl:
                self._QueryTblRecCount = len(self._QueryTbl['__data__'])

            # Проинициализировать бенды групп
            for grp in self._Template['groups']:
                grp['old_rec'] = None

            time_start = time.time()
            log.info('REPORT <%s> GENERATE START' % self._RepName)

            # III. Вывод данных в отчет
            # Создать отчет
            self._Rep = copy.deepcopy(IC_REP_TMPL)
            self._Rep['name'] = self._RepName

            # Инициализация необходимых переменных
            field_idx = {}      # Индексы полей
            i = 0
            i_rec = 0
            # Перебор полей таблицы запроса
            if self._QueryTbl and '__fields__' in self._QueryTbl:
                for cur_field in self._QueryTbl['__fields__']:
                    field_idx[cur_field] = i
                    i += 1

            # Если записи в таблице запроса есть, то ...
            if self._QueryTblRecCount:
                # Проинициализировать текущую строку для использования
                # ее в заголовке отчета
                rec = self._QueryTbl['__data__'][i_rec]
                # Заполнить словарь текущей записи
                for field_name in field_idx.keys():
                    val = rec[field_idx[field_name]]
                    # Предгенерация значения данных ячейки
                    self._CurRec[field_name] = val
                # Прописать индекс текущей записи
                self._CurRec['ic_sys_num_rec'] = i_rec

            # Верхний колонтитул
            if self._Template['upper']:
                self._genUpper(self._Template['upper'])
            
            # Вывести в отчет заголовок
            self._genHeader(self._Template['header'])

            # Главный цикл
            # Перебор записей таблицы запроса
            while i_rec < self._QueryTblRecCount:
                # Обработка групп
                # Проверка смены группы в описании всех групп
                # и найти индекс самой общей смененной группы
                i_grp_out = -1      # индекс самой общей смененной группы
                # Флаг начала генерации (примечания групп не выводяться)
                start_gen = False
                for i_grp in range(len(self._Template['groups'])):
                    grp = self._Template['groups'][i_grp]
                    if grp['old_rec']:
                        # Проверить условие вывода примечания группы
                        if self._CurRec[grp['field']] != grp['old_rec'][grp['field']]:
                            i_grp_out = i_grp
                            break
                    else:
                        i_grp_out = 0
                        start_gen = True
                        break
                if i_grp_out != -1:
                    # Вывести примечания
                    if start_gen is False:
                        for i_grp in range(len(self._Template['groups'])-1, i_grp_out-1, -1):
                            grp = self._Template['groups'][i_grp]
                            self._genGrpFooter(grp)
                    # Вывести заголовки
                    for i_grp in range(i_grp_out, len(self._Template['groups'])):
                        grp = self._Template['groups'][i_grp]
                        grp['old_rec'] = copy.deepcopy(self._CurRec)
                        self._genGrpHeader(grp)
                    
                # Область данных
                self._genDetail(self._Template['detail'])

                # Увеличить суммы суммирующих ячеек
                self._sumIterate(self._TemplateSheet, self._CurRec)

                # Перейти на следующую запись
                i_rec += 1
                # Заполнить словарь текущей записи
                if i_rec < self._QueryTblRecCount:
                    rec = self._QueryTbl['__data__'][i_rec]
                    # Заполнить словарь текущей записи
                    for field_name in field_idx.keys():
                        val = rec[field_idx[field_name]]
                        # Предгенерация значения данных ячейки
                        self._CurRec[field_name] = val
                    # Прописать индекс текущей записи
                    self._CurRec['ic_sys_num_rec'] = i_rec

            # Вывести примечания после области данных
            for i_grp in range(len(self._Template['groups'])-1, -1, -1):
                grp = self._Template['groups'][i_grp]
                if grp['old_rec']:
                    self._genGrpFooter(grp)
                else:
                    break
            # Вывести в отчет примечание отчета
            self._genFooter(self._Template['footer'])
            # Нижний колонтитул
            if self._Template['under']:
                self._genUnder(self._Template['under'])

            # Параметры страницы
            self._Rep['page_setup'] = self._Template['page_setup']

            # Прогресс бар
            log.info('REPORT <%s> GENERATE STOP. Time <%d> sec.' % (self._RepName, time.time()-time_start))

            return self._Rep
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации отчета.')
            return None

    def _genHeader(self, Header_):
        """
        Сгенерировать заголовок отчета и перенести ее в выходной отчет.
        @param Header_: Бэнд заголовка.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            log.debug(u'Генерация заголовка')
            # Добавлять будем в конец отчета,
            # поэтому опреелить максимальную строчку
            max_row = len(self._Rep['sheet'])
            i_row = 0
            cur_height = 0
            # Перебрать все ячейки бэнда
            for row in range(Header_['row'], Header_['row'] + Header_['row_size']):
                for col in range(Header_['col'], Header_['col'] + Header_['col_size']):
                    if self._TemplateSheet[row][col]:
                        self._genCell(self._TemplateSheet, row, col,
                                      self._Rep, max_row+i_row, col, self._CurRec)
                        cur_height = self._TemplateSheet[row][col]['height']
                i_row += 1
                # Увеличить текущую координату Y
                self._cur_top += cur_height
            # Прописать область
            self._Rep['header'] = {'row': max_row,
                                   'col': Header_['col'],
                                   'row_size': i_row,
                                   'col_size': Header_['col_size'],
                                   }
            # Очистить сумы суммирующих ячеек
            self._TemplateSheet = self._clearSum(self._TemplateSheet, 0, len(self._TemplateSheet))
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации заголовка отчета <%s>.' % self._RepName)
            return False
            
    def _genFooter(self, Footer_):
        """
        Сгенерировать примечание отчета и перенести ее в выходной отчет.
        @param Footer_: Бэнд примечания.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            # Подвала отчета просто нет
            if not Footer_:
                return True

            # Добавлять будем в конец отчета,
            # поэтому опреелить максимальную строчку
            max_row = len(self._Rep['sheet'])
            i_row = 0       # Счетчик строк бэнда
            cur_height = 0
            # Перебрать все ячейки бэнда
            for row in range(Footer_['row'], Footer_['row']+Footer_['row_size']):
                for col in range(Footer_['col'], Footer_['col']+Footer_['col_size']):
                    if self._TemplateSheet[row][col]:
                        self._genCell(self._TemplateSheet, row, col,
                                      self._Rep, max_row+i_row, col, self._CurRec)
                        cur_height = self._TemplateSheet[row][col]['height']
                i_row += 1
                # Увеличить текущую координату Y
                self._cur_top += cur_height
            # Прописать область
            self._Rep['footer'] = {'row': max_row,
                                   'col': Footer_['col'],
                                   'row_size': i_row,
                                   'col_size': Footer_['col_size'],
                                   }
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации примечания отчета <%s>.' % self._RepName)
            return False

    def _genDetail(self, Detail_):
        """
        Сгенерировать область данных отчета и перенести ее в выходной отчет.
        @param Detail_: Бэнд области данных.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            # Добавлять будем в конец отчета,
            # поэтому опреелить максимальную строчку
            max_row = len(self._Rep['sheet'])
            i_row = 0   # Счетчик строк бэнда
            cur_height = 0
            # Перебрать все ячейки бэнда
            for row in range(Detail_['row'], Detail_['row']+Detail_['row_size']):
                for col in range(Detail_['col'], Detail_['col']+Detail_['col_size']):
                    if self._TemplateSheet[row][col]:
                        self._genCell(self._TemplateSheet, row, col,
                                      self._Rep, max_row+i_row, col, self._CurRec)
                        cur_height = self._TemplateSheet[row][col]['height']
                i_row += 1
                # Увеличить текущую координату Y
                self._cur_top += cur_height
            # Прописать область
            if self._Rep['detail'] == {}:
                self._Rep['detail'] = {'row': max_row,
                                       'col': Detail_['col'],
                                       'row_size': i_row,
                                       'col_size': Detail_['col_size'],
                                       }
            else:
                self._Rep['detail']['row_size'] += i_row

            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации области данных отчета <%s>.' % self._RepName)
            return False

    def _genGrpHeader(self, RepGrp_):
        """
        Генерация заголовка группы.
        @param RepGrp_: Словарь IC_REP_GRP, описывающий группу.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            band = RepGrp_['header']
            if not band:
                return False
            # Добавлять будем в конец отчета,
            # поэтому опреелить максимальную строчку
            max_row = len(self._Rep['sheet'])
            i_row = 0   # Счетчик строк бэнда
            cur_height = 0
            # Перебрать все ячейки бэнда
            for row in range(band['row'], band['row']+band['row_size']):
                for col in range(band['col'], band['col']+band['col_size']):
                    if self._TemplateSheet[row][col]:
                        self._genCell(self._TemplateSheet, row, col,
                                      self._Rep, max_row+i_row, col, self._CurRec)
                        cur_height = self._TemplateSheet[row][col]['height']
                i_row += 1
                # Увеличить текущую координату Y
                self._cur_top += cur_height
            # Очистить сумы суммирующих ячеек
            # ВНИМАНИЕ!!! Итоговых ячеек не бывает в заголовках. Поэтому я не обработываю их
            band = RepGrp_['footer']
            if band:
                self._TemplateSheet = self._clearSum(self._TemplateSheet, band['row'], band['row']+band['row_size'])
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации заголовка группы <%s> отчета <%s>.' % (RepGrp_['field'], self._RepName))
            return False

    def _genGrpFooter(self, RepGrp_):
        """
        Генерация примечания группы.
        @param RepGrp_: Словарь IC_REP_GRP, описывающий группу.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            band = RepGrp_['footer']
            if not band:
                return False
            # Добавлять будем в конец отчета,
            # поэтому опреелить максимальную строчку
            max_row = len(self._Rep['sheet'])
            i_row = 0   # Счетчик строк бэнда
            cur_height = 0
            # Перебрать все ячейки бэнда
            for row in range(band['row'], band['row']+band['row_size']):
                for col in range(band['col'], band['col']+band['col_size']):
                    if self._TemplateSheet[row][col]:
                        self._genCell(self._TemplateSheet, row, col,
                                      self._Rep, max_row+i_row, col, RepGrp_['old_rec'])
                        cur_height = self._TemplateSheet[row][col]['height']
                i_row += 1
                # Увеличить текущую координату Y
                self._cur_top += cur_height
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации примечания группы <%s> отчета <%s>.' % (RepGrp_['field'], self._RepName))
            return False

    def _genUpper(self, Upper_):
        """
        Сгенерировать верхний колонтитул/заголовок страницы отчета и перенести ее в выходной отчет.
        @param Upper_: Бэнд верхнего колонтитула.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            if 'row' not in Upper_ or 'col' not in Upper_ or \
               'row_size' not in Upper_ or 'col_size' not in Upper_:
                # Не надо обрабатывать строки
                self._Rep['upper'] = Upper_
                return True
                
            # Добавлять будем в конец отчета,
            # поэтому опреелить максимальную строчку
            max_row = len(self._Rep['sheet'])
            i_row = 0
            cur_height = 0
            # Перебрать все ячейки бэнда
            for row in range(Upper_['row'], Upper_['row']+Upper_['row_size']):
                for col in range(Upper_['col'], Upper_['col']+Upper_['col_size']):
                    if self._TemplateSheet[row][col]:
                        self._genCell(self._TemplateSheet, row, col,
                                      self._Rep, max_row+i_row, col, self._CurRec)
                        cur_height = self._TemplateSheet[row][col]['height']
                i_row += 1
                # Увеличить текущую координату Y
                self._cur_top += cur_height
            # Прописать область
            self._Rep['upper'] = copy.deepcopy(Upper_)
            self._Rep['upper']['row'] = max_row
            self._Rep['upper']['row_size'] = i_row
            
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации верхнего колонтитула отчета <%s>.' % self._RepName)
            return False

    def _genUnder(self, Under_):
        """
        Сгенерировать нижний колонтитул отчета и перенести ее в выходной отчет.
        @param Under_: Бэнд нижнего колонтитула.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            if 'row' not in Under_ or 'col' not in Under_ or \
               'row_size' not in Under_ or 'col_size' not in Under_:
                # Не надо обрабатывать строки
                self._Rep['under'] = Under_
                return True
                
            # Добавлять будем в конец отчета,
            # поэтому опреелить максимальную строчку
            max_row = len(self._Rep['sheet'])
            i_row = 0
            cur_height = 0
            # Перебрать все ячейки бэнда
            for row in range(Under_['row'], Under_['row']+Under_['row_size']):
                for col in range(Under_['col'], Under_['col']+Under_['col_size']):
                    if self._TemplateSheet[row][col]:
                        self._genCell(self._TemplateSheet, row, col,
                                      self._Rep, max_row+i_row, col, self._CurRec)
                        cur_height = self._TemplateSheet[row][col]['height']
                i_row += 1
                # Увеличить текущую координату Y
                self._cur_top += cur_height
            # Прописать область
            self._Rep['under'] = copy.deepcopy(Under_)
            self._Rep['under']['row'] = max_row
            self._Rep['under']['row_size'] = i_row
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации нижнего колонтитула отчета <%s>.' % self._RepName)
            return False
            
    def _genSubReport(self, SubReportName_, Row_):
        """
        Генерация под-отчета.
        @param SubReportName_: Имя под-отчета.
        @param Row_: Номер строки листа, после которой будет вставляться под-отчет.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            if '__sub__' in self._QueryTbl and self._QueryTbl['__sub__']:
                if SubReportName_ in self._QueryTbl['__sub__']:
                    # Если есть данные под-отчета, тогда запустить генерацию
                    report = self._QueryTbl['__sub__'][SubReportName_]['report']
                    if isinstance(report, str):
                        # Импорт для генерации под-отчетов
                        from ic.report import icreptemplate
                        # Под-отчет задан именем файла.
                        template = icreptemplate.icExcelXMLReportTemplate()

                        self._QueryTbl['__sub__'][SubReportName_]['report'] = template.read(report)
                    # Запуск генерации подотчета
                    rep_gen = icReportGenerator()
                    rep_result = rep_gen.generate(self._QueryTbl['__sub__'][SubReportName_]['report'],
                                                  self._QueryTbl['__sub__'][SubReportName_],
                                                  self._QueryTbl['__sub__'][SubReportName_]['__variables__'],
                                                  self._QueryTbl['__sub__'][SubReportName_]['__coord_fill__'])
                    # Вставить результат под-отчета после строки
                    self._Rep['sheet'] = self._Rep['sheet'][:Row_]+rep_result['sheet']+self._Rep['sheet'][Row_:]
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации под-отчета <%s> отчета <%s>.' % (SubReportName_, self._RepName))
            return False

    def _genCell(self, FromSheet_, FromRow_, FromCol_, ToRep_, ToRow_, ToCol_, record):
        """
        Генерация ячейки из шаблона в выходной отчет.
        @param FromSheet_: Из листа шаблона.
        @param FromRow_: Координаты ячейки шаблона. Строка.
        @param FromCol_: Координаты ячейки шаблона. Столбец.
        @param ToRep_: В отчет.
        @param ToRow_: Координаты ячейки отчета. Строка.
        @param ToCol_: Координаты ячейки отчета. Столбец.
        @param record: Запись.
        @return: Возвращает результат выполнения операции True/False.
        """
        try:
            cell = copy.deepcopy(FromSheet_[FromRow_][FromCol_])

            # Коррекция координат ячейки
            cell['top'] = self._cur_top
            # Генерация текста ячейки
            if self._CoordFill and (ToRow_, ToCol_) in self._CoordFill:
                # Координатные замены
                fill_val = str(self._CoordFill[(ToRow_, ToCol_)])
                cell['value'] = self._genTxt({'value': fill_val}, record, ToRow_, ToCol_)
            else:
                # Перенести все ячейки из шаблона в выходной отчет
                cell['value'] = self._genTxt(cell, record, ToRow_, ToCol_)

            # Установка атирибутов ячейки по умолчанию
            # Заполнение некоторых атрибутов ячейки по умолчанию
            if self.AttrDefault and isinstance(self.AttrDefault, dict):
                cell.update(self.AttrDefault)
                
            # Установить описание ячейки отчета.
            if len(ToRep_['sheet']) <= ToRow_:
                # Расширить строки
                for i_row in range(len(ToRep_['sheet']), ToRow_+1):
                    ToRep_['sheet'].append([])
            if len(ToRep_['sheet'][ToRow_]) <= ToCol_:
                # Расширить колонки
                for i_col in range(len(ToRep_['sheet'][ToRow_]), ToCol_+1):
                    ToRep_['sheet'][ToRow_].append(None)
            # Установить описание колонки
            if cell['visible']:
                ToRep_['sheet'][ToRow_][ToCol_] = cell
            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации ячейки шаблона <%s>.' % self._RepName)
            return False
        
    def _genTxt(self, cell, record=None, CellRow_=None, CellCol_=None):
        """
        Генерация текста.
        @param cell: Ячейка.
        @param record: Словарь, описывающий текущую запись таблицы запроса.
            Формат: { <имя поля> : <значение поля>, ...}
        @param CellRow_: Номер строки ячейки в результирующем отчете.
        @param CellCol_: Номер колонки ячейки в результирующем отчете.
        @return: Возвращает сгенерированное значение.
        """
        try:
            # Проверка на преобразование типов
            cell_val = cell['value']
            if cell_val is not None and type(cell_val) not in (str, unicode):
                cell_val = str(cell_val)
            if cell_val not in self._cellFmt:
                parsed_fmt = self.funcStrParse(cell_val)
                self._cellFmt[cell_val] = parsed_fmt
            else:
                parsed_fmt = self._cellFmt[cell_val]

            func_str = []   # Выходной список значений
            i_sum = 0
            # Перебрать строки функционала
            for cur_func in parsed_fmt['func']:

                # Функция
                if re.search(REP_FUNC_PATT, cur_func):
                    value = str(self._execFuncGen(cur_func[2:-2], locals()))
                    
                # Ламбда-выражение
                elif re.search(REP_LAMBDA_PATT, cur_func):
                    # ВНИМАНИЕ: Лямбда-выражение в шаблоне должно иметь
                    #     1 аргумент это словарь записи.
                    #     Например:
                    #         [~rec: rec['name']=='Петров'~]
                    lambda_func = eval('lambda '+cur_func[2:-2])
                    value = str(lambda_func(record))
                    
                # Переменная
                elif re.search(REP_VAR_PATT, cur_func):
                    var_name = cur_func[2:-2]
                    log.debug(u'Обработка переменной <%s> -- %s' % (var_name, var_name in self._NameSpace))
                    value = str(self._NameSpace.setdefault(var_name, None))
                    
                # Блок кода
                elif re.search(REP_EXEC_PATT, cur_func):
                    # ВНИМАНИЕ: В блоке кода доступны объекты cell и record.
                    #     Если надо вывести информацию, то ее надо выводить в
                    #     переменную value.
                    #     Например:
                    #         [=if record['name']=='Петров':
                    #             value='-'=]
                    value = ''
                    exec_func = cur_func[2:-2].strip()
                    try:
                        exec exec_func
                    except:
                        log.fatal(u'Ошибка выполнения блока кода <%s>' % textfunc.toUnicode(exec_func))
                    
                # Системная функция
                elif re.search(REP_SYS_PATT, cur_func):
                    # Функция суммирования
                    if cur_func[2:6].lower() == 'sum(':
                        value = str(cell['sum'][i_sum]['value'])
                        i_sum += 1  # Перейти к следующей сумме
                    # Функция вычисления среднего значения
                    elif cur_func[2:6].lower() == 'avg(':
                        if 'ic_sys_num_rec' not in record:
                            record['ic_sys_num_rec'] = 0
                        value = str(cell['sum'][i_sum]['value'] / (record['ic_sys_num_rec'] + 1))
                        i_sum += 1  # Перейти к следующей сумме
                    elif cur_func[2:-2].lower() == 'n':
                        if 'ic_sys_num_rec' not in record:
                            record['ic_sys_num_rec'] = 0
                        sys_num_rec = record['ic_sys_num_rec']
                        value = str(sys_num_rec + 1)
                    else:
                        # Вывести сообщение об ошибке в лог
                        log.warning(u'Неизвестная системная функция <%s> шаблона <%s>.' % (textfunc.toUnicode(cur_func),
                                                                                           self._RepName))
                        value = ''
                        
                # Стиль
                elif re.search(REP_STYLE_PATT, cur_func):
                    value = ''
                    style_name = cur_func[2:-2]
                    self._setStyleAttr(style_name)
                    
                # Под-отчеты
                elif re.search(REP_SUBREPORT_PATT, cur_func):
                    value = ''
                    subreport_name = cur_func[2:-2]
                    self._genSubReport(subreport_name, CellRow_)

                # Поле
                elif re.search(REP_FIELD_PATT, cur_func):
                    field_name = str((cur_func[2:-2]))
                    try:
                        value = record[field_name]
                    except KeyError:
                        log.warning(u'В строке (%s) поле <%s> не найдено' % (textfunc.toUnicode(record),
                                                                             textfunc.toUnicode(field_name)))
                        value = ''

                # ВНИМАНИЕ! В значении ячейки тоже могут быть управляющие коды
                value = self._genTxt({'value': value}, record)
                func_str.append(value)

            # Заполнение формата
            return self._valueFormat(parsed_fmt['fmt'], func_str)
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации текста ячейки <%s> шаблона <%s>.' % (textfunc.toUnicode(cell['value']),
                                                                              self._RepName))
            return None

    def _valueFormat(self, Fmt_, DataLst_):
        """
        Заполнение формата значения ячейки.
        @param Fmt_: Формат.
        @param DataLst_: Данные, которые нужно поместить в формат.
        @return: Возвращает строку, соответствующую формату.
        """
        if isinstance(Fmt_, str):
            Fmt_ = unicode(Fmt_, DEFAULT_ENCODING)

        # Заполнение формата
        if DataLst_ is []:
            if Fmt_:
                value = Fmt_
            else:
                return None
        elif DataLst_ == [None] and Fmt_ == '%s':
            return None
        # Обработка значения None
        elif bool(None in DataLst_):
            data_lst = [{None: ''}.setdefault(val, val) for val in DataLst_]
            value = Fmt_ % tuple(data_lst)
        else:
            value = Fmt_ % tuple(DataLst_)
        return value
        
    def _setStyleAttr(self, StyleName_):
        """
        Установить атрибуты по умолчанию ячеек по имени стиля из библиотеки стилей.
        @param StyleName_: Имя стиля из библиотеки стилей.
        """
        if self._StyleLib and StyleName_ in self._StyleLib:
            self.AttrDefault = self._StyleLib[StyleName_]
        else:
            self.AttrDefault = None
        
    def _getSum(self, Formul_):
        """
        Получить сумму по формуле.
        @param Formul_: Формула.
        @return: Возвращает строковое значение суммы.
        """
        return '0'

    def _initSumCells(self, Sheet_):
        """
        Выявление и инициализация ячеек с суммами.
        @param Sheet_: Описание листа отчета.
        @return: Возвращает описание листа с корректным описанием ячеек с суммами.
            В результате ошибки возвращает старое описание листа.
        """
        try:
            sheet = Sheet_
            # Просмотр и коррекция каждой ячейки листа
            for row in range(len(sheet)):
                for col in range(len(sheet[row])):
                    if sheet[row][col]:
                        sheet[row][col] = self._initSumCell(sheet[row][col])
            return sheet
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка инициализации суммирующих ячеек шаблона <%s>.' % self._RepName)
            return Sheet_

    def _initSumCell(self, cell):
        """
        Инициализация суммарной ячейки.
        @param cell: Описание ячейки.
        @return: Возвращает скоррекстированное описание ячейки.
            В случае ошибки возвращает старое описание ячейки.
        """
        try:
            cell = cell
            # Проверка на преобразование типов
            cell_val = cell['value']
            if cell_val is not None and not isinstance(cell_val, str):
                cell_val = str(cell_val)
            parsed_fmt = self.funcStrParse(cell_val, [REP_SYS_PATT])
            # Перебрать строки функционала
            for cur_func in parsed_fmt['func']:
                # Системная функция
                if re.search(REP_SYS_PATT, cur_func):
                    # Функция суммирования
                    if cur_func[2:6].lower() in ('sum(', 'avg('):
                        # Если данные суммирующей ячейки не инициализированы, то
                        if cell['sum'] is None:
                            cell['sum'] = []
                        # Проинициализировать данные суммарной ячейки
                        cell['sum'].append(copy.deepcopy(IC_REP_SUM))
                        cell['sum'][-1]['formul'] = cur_func[6:-3].replace(REP_SUM_FIELD_START, 'record[\'').replace(REP_SUM_FIELD_STOP, '\']')
            return cell
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка инициализации ячейки <%s>.' % cell)
            return cell

    def _sumIterate(self, Sheet_, record):
        """
        Итерация суммирования.
        @param Sheet_: Описание листа отчета.
        @param record: Запись, на которой вызывается итерация.
        @return: Возвращает описание листа с корректным описанием ячеек с суммами.
            В результате ошибки возвращает старое описание листа.
        """
        try:
            sheet = Sheet_
            # Просмотр и коррекция каждой ячейки листа
            for row in range(len(sheet)):
                for col in range(len(sheet[row])):
                    # Если ячейка определена, то ...
                    if sheet[row][col]:
                        # Если ячейка суммирующая,
                        # то выполнить операцию суммирования
                        if sheet[row][col]['sum'] is not None and sheet[row][col]['sum'] is not []:
                            for cur_sum in sheet[row][col]['sum']:
                                try:
                                    value = eval(cur_sum['formul'], globals(), locals())
                                except:
                                    log.warning(u'Ошибка выполнения формулы для подсчета сумм <%s>.' % cur_sum)
                                    value = 0.0
                                try:
                                    if value is None:
                                        value = 0.0
                                    else:
                                        value = float(value)
                                    cur_sum['value'] += value
                                except:
                                    log.warning(u'Ошибка итерации сумм <%s>+<%s>' % (cur_sum['value'], value))

            return sheet
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка итерации сумм суммирующих ячеек шаблона <%s>.' % self._RepName)
            return Sheet_

    def _clearSum(self, Sheet_, RowStart_, RowStop_):
        """
        Обнуление сумм.
        @param Sheet_: Описание листа отчета.
        @param RowStart_: Начало бэнда обнуления.
        @param RowStop_: Конец бэнда обнуления.
        @return: Возвращает описание листа с корректным описанием ячеек с суммами.
            В результате ошибки возвращает старое описание листа.
        """
        try:
            sheet = Sheet_
            # Просмотр и коррекция каждой ячейки листа
            for row in range(RowStart_, RowStop_):
                for col in range(len(sheet[row])):
                    # Если ячейка определена, то ...
                    if sheet[row][col]:
                        # Если ячейка суммирующая, то выполнить операцию обнуления
                        if sheet[row][col]['sum'] is not None and sheet[row][col]['sum'] is not []:
                            for cur_sum in sheet[row][col]['sum']:
                                cur_sum['value'] = 0
            return sheet
        except:
            # Вывести сообщение об ошибке в лог
            log.error(u'Ошибка обнуления сумм суммирующих ячеек шаблона <%s>.' % self._RepName)
            return Sheet_
        
    # Функции-свойства
    def getCurRec(self):
        return self._CurRec

    def funcStrParse(self, Str_, Patterns_=ALL_PATTERNS):
        """
        Разобрать строку на формат и исполняемый код.
        @param Str_: Разбираемая строка.
        @param Patterns_: Cписок строк патернов тегов обозначения
            начала и конца функционала.
        @return: Возвращает словарь следующей структуры:
            {
            'fmt': Формат строки без строк исполняемого кода вместо него стоит %s;
            'func': Список строк исполняемого кода.
            }
            В случае ошибки возвращает None.
        """
        try:
            # Инициализация структуры
            ret = {'fmt': '', 'func': []}

            # Проверка аргументов
            if not Str_:
                return ret

            # Заполнение патерна
            pattern = r''
            for cur_sep in Patterns_:
                pattern += cur_sep
                if cur_sep != Patterns_[-1]:
                    pattern += r'|'
                    
            # Разбор строки на обычные строки и строки функционала
            parsed_str = [x for x in re.split(pattern, Str_) if x is not None]
            # Перебор тегов
            for i_parse in range(len(parsed_str)):
                # Перебор патернов функционала
                func_find = False
                for cur_patt in Patterns_:
                    # Какой-то функционал
                    if re.search(cur_patt, parsed_str[i_parse]):
                        ret['func'].append(parsed_str[i_parse])
                        # И добавить в формат %s
                        ret['fmt'] += '%s'
                        func_find = True
                        break
                # Обычная строка
                if func_find is False:
                    ret['fmt'] += parsed_str[i_parse]
            return ret
        except:
            log.fatal(u'Ошибка формата <%s> ячейки шаблона <%s>.' % (Str_, self._RepName))
            return None

    def _execFuncGen(self, Func_, Locals_):
        """
        Выполнить функцию при генерации.
        """
        re_import = not ic_mode.isRuntimeMode()
        return ic_exec.execFuncStr(Func_, Locals_, re_import)
