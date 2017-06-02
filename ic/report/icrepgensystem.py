#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль функций системы генератора отчетов.
В качестве ключей таблицы запроса могут быть:
    '__variables__': Словарь переменных отчета,
    '__coord_fill__': Словарь координатных замен,
    '__sql__': SQL выражения указания таблицы запроса,
        ВНИМАНИЕ! SQL выражение задается без сигнатуры SQL_SIGNATURE.
            Просто в виде SQL выражения типа <SELECT>,
    '__fields__': Список описаний полей таблицы запроса.
    '__data__': Список записей таблицы запроса,

Все эти ключи обрабатываются в процессе генерации отчета.
"""

# Подключение библиотек
import os
import os.path
import re
import shutil
import sqlalchemy

from ic.std.utils import execfunc
from ic.std.log import log
from ic.std.utils import filefunc
from ic.std.utils import res
from ic.std.dlg import dlg

from ic.report import icrepgen
from ic.report import icreptemplate

__version__ = (0, 0, 2, 3)

# Константы подсистемы
DEFAULT_REP_TMPL_FILE = os.path.dirname(__file__)+'/new_report_template.ods'

OFFICE_OPEN_CMD_FORMAT = 'libreoffice %s'

ODS_TEMPLATE_EXT = '.ods'
XML_TEMPLATE_EXT = '.xml'
DEFAULT_TEMPLATE_EXT = ODS_TEMPLATE_EXT
DEFAULT_REPORT_TEMPLATE_EXT = '.rprt'

DEFAULT_REPORT_DIR = os.path.abspath(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))+'/reports/')

# Сигнатуры значений бендов отчета (Длина сигнатуры д.б. 4 символа)
DB_URL_SIGNATURE = 'URL:'
SQL_SIGNATURE = 'SQL:'
CODE_SIGNATURE = 'PRG:'
PY_SIGNATURE = 'PY:'


class icReportGeneratorSystem:
    """
    Класс системы генерации отчетов. Абстрактный класс.
    """

    def __init__(self, report=None, ParentForm_=None):
        """
        Конструктор класса.
        @param report: Шаблон отчета.
        @param ParentForm_: Родительская форма, необходима для вывода сообщений.
        """
        # Шаблон отчета
        self._Rep = report
        # Таблица запроса
        self._QueryTab = None

        # Родительская форма, необходима для вывода сообщений.
        self._ParentForm = ParentForm_

        # Предварительный просмотр
        self.PrintPreview = None

    def getProfileDir(self):
        """
        Папка профиля программы.
        """
        return filefunc.getProfilePath()

    def getReportDescription(self):
        """
        Описание отчета.
        @return: Строку описание отчета или его имя если описание не
            определено.
        """
        description = u''
        if self._Rep:
            description = self._Rep.get('description', self._Rep.get('name', u''))
        return description

    def setParentForm(self, ParentForm_):
        """
        Установить родительскую форму для определения папки отчетов.
        """
        self._ParentForm = ParentForm_
        
    def getParentForm(self):
        """
        Родительская форма.
        """
        return self._ParentForm
        
    def reloadRepData(self, RepTmplFileName_=None):
        """
        Перегрузить данные отчета.
        @param RepTmplFileName_: Имя файла шаблона отчета.
        """
        self._Rep = filefunc.loadResourceFile(RepTmplFileName_, bRefresh=True)
        
    def setRepData(self, report):
        """
        Установить данные отчета.
        """
        self._Rep = report

    def selectAction(self, report=None, *args, **kwargs):
        """
        Запуск генерации отчета с последующим выбором действия.
        @param report: Полное описание шаблона отчета.
        """
        return

    def Preview(self, report=None, *args, **kwargs):
        """
        Предварительный просмотр.
        @param report: Полное описание шаблона отчета.
        """
        return

    def Print(self, report=None, *args, **kwargs):
        """
        Печать.
        @param report: Полное описание шаблона отчета.
        """
        return 

    def PageSetup(self):
        """
        Установка параметров страницы.
        """
        return

    def Convert(self, report=None, ToFile_=None, *args, **kwargs):
        """
        Конвертирование результатов отчета.
        @param report: Полное описание шаблона отчета.
        @param ToFile_: Имя файла, куда необходимо сохранить отчет.
        """
        return

    def Export(self, report=None, ToFile_=None, *args, **kwargs):
        """
        Вывод результатов отчета во внешнюю программу.
        @param report: Полное описание шаблона отчета.
        @param ToFile_: Имя файла, куда необходимо сохранить отчет.
        """
        return self.Convert(report, ToFile_)

    def New(self, dst_path=None):
        """
        Создание нового отчета.
        @param dst_path: Результирующая папка, в которую будет помещен новый файл.
        """
        return self.NewByOffice(dst_path)
        
    def NewByOffice(self, dst_path=None):
        """
        Создание нового отчета средствами LibreOffice Calc.
        @param dst_path: Результирующая папка, в которую будет помещен новый файл.
        """
        try:
            src_filename = DEFAULT_REP_TMPL_FILE
            new_filename = dlg.getTextInputDlg(self._ParentForm,
                                               u'Создание нового файла',
                                               u'Введите имя файла шаблона отчета')
            if os.path.splitext(new_filename)[1] != '.ods':
                new_filename += '.ods'

            if dst_path is None:
                # Необходимо определить результирующий путь
                dst_path = dlg.getDirDlg(self._ParentForm,
                                         u'Папка хранения')
                if not dst_path:
                    dst_path = os.getcwd()

            dst_filename = os.path.join(dst_path, new_filename)
            if os.path.exists(dst_filename):
                if dlg.getAskBox(u'Заменить существующий файл?'):
                    shutil.copyfile(src_filename, dst_filename)
            else:
                shutil.copyfile(src_filename, dst_filename)

            cmd = OFFICE_OPEN_CMD_FORMAT % dst_filename
            log.debug('Command <%s>' % cmd)
            os.system(cmd)

            return True
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'New report template by LibreOffice Calc')

    def Edit(self, report=None):
        """
        Редактирование отчета.
        @param report: Полное описание шаблона отчета.
        """
        return

    def Update(self, RepTemplateFileName_=None):
        """
        Обновить шаблон отчета в системе генератора отчетов.
        @param RepTemplateFileName_: Имя файла шаблона отчета.
            Если None, то должен производиться запрос на выбор этого файла.
        """
        if RepTemplateFileName_ is None:
            filename = dlg.getFileDlg(self._ParentForm, u'Выберите шаблон отчета:',
                                      u'Microsoft Excel 2003 XML (*.xml)|*.xml|Электронные таблицы ODF (*.ods)|*.ods',
                                      self.getReportDir())
        else:
            filename = os.path.abspath(os.path.normpath(RepTemplateFileName_))

        if os.path.isfile(filename):

            # Конвертация
            log.debug(u'Начало конвертации <%s>' % filename)
            template = None
            if os.path.exists(os.path.splitext(filename)[0] + DEFAULT_TEMPLATE_EXT):
                tmpl_filename = os.path.splitext(filename)[0] + DEFAULT_TEMPLATE_EXT
                template = icreptemplate.icODSReportTemplate()
            elif os.path.exists(os.path.splitext(filename)[0] + ODS_TEMPLATE_EXT):
                tmpl_filename = os.path.splitext(filename)[0] + ODS_TEMPLATE_EXT
                template = icreptemplate.icODSReportTemplate()
            elif os.path.exists(os.path.splitext(filename)[0] + XML_TEMPLATE_EXT):
                tmpl_filename = os.path.splitext(filename)[0] + XML_TEMPLATE_EXT
                template = icreptemplate.icExcelXMLReportTemplate()
            else:
                log.warning(u'Not find report template for <%s>' % filename)
            if template:
                rep_template = template.read(tmpl_filename)
                new_filename = os.path.splitext(filename)[0]+DEFAULT_REPORT_TEMPLATE_EXT
                res.saveResourcePickle(new_filename, rep_template)
            log.info(u'Конец конвертации')
   
    def OpenModule(self,RepTemplateFileName_=None):
        """
        Открыть модуль отчета в редакторе.
        """
        if RepTemplateFileName_ is None:
            log.warning(u'Не определен файл модуля отчета')
        # Определить файл *.xml
        module_file = os.path.abspath(os.path.splitext(RepTemplateFileName_)[0]+'.py')
        if os.path.exists(module_file):
            try:
                self._ParentForm.GetParent().ide.OpenFile(module_file)
            except:
                log.fatal(u'Ошибка открытия модуля <%s>' % module_file)
        else:
            dlg.getMsgBox(u'Файл модуля отчета <%s> не найден.' % module_file)
        
    def generate(self, report=None, db_url=None, sql=None, stylelib=None, vars=None, *args, **kwargs):
        """
        Запустить генератор отчета.
        @param report: Шаблон отчета.
        @param db_url: Connection string в виде url. Например
            postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
        @param sql: Запрос SQL.
        @param stylelib: Библиотека стилей.
        @param vars: Словарь переменных отчета.
        @return: Возвращает сгенерированный отчет или None в случае ошибки.
        """
        return None

    def generateReport(self, report=None, *args, **kwargs):
        """
        Запустить генератор отчета.
        @param report: Шаблон отчета.
        @return: Возвращает сгенерированный отчет или None в случае ошибки.
        """
        return None

    def InitRepTemplate(self, report, QueryTab_=None):
        """
        Прочитать данные о шаблоне отчета.
        @param report: Полное описание шаблона отчета.
        @param QueryTab_: Таблица запроса.
        """
        # 1. Прочитать структру отчета
        self._Rep = report
        # 2. Таблица запроса
        self._QueryTab = QueryTab_

        # 3. Скорректировать шаблон для нормальной обработки генератором
        self._Rep = self.RepSQLObj2SQLite(self._Rep, res.loadResourceFile(res.icGetTabResFileName()))

        # 5. Коррекция параметров БД и запроса
        # self._Rep['data_source'] = ic_exec.ExecuteMethod(self._Rep['data_source'], self)
        # self._Rep['query'] = ic_exec.ExecuteMethod(self._Rep['query'], self)

        # !!! ВНИМАНИЕ!!!
        # --- Проверка случая когда функция  запроса возвращает SQL запрос ---
        # if type(self._Rep['query'])==type(''):
        #     self._Rep['query']=ic.db.tabrestr.icQueryTxtSQLObj2SQLite(self._Rep['query'],\
        #         ic.utils.util.readAndEvalFile(ic.utils.resource.icGetTabResFileName()))

    def _getSQLQueryTable(self, report, db_url=None, sql=None):
        """
        Получить таблицу запроса.
        @param report: Шаблон отчета.
        @param db_url: Connection string в виде url. Например
            postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
        @param sql: Текст SQL запроса.
        @return: Функция возвращает словарь -
            ТАБЛИЦА ЗАПРОСА ПРЕДСТАВЛЯЕТСЯ В ВИДЕ СЛОВАРЯ 
            {'__fields__':имена полей таблицы,'__data__':данные таблицы}
        """
        result = None
        # Инициализация
        db_connection = None
        try:
            if not db_url:
                data_source = report['data_source']

                if not data_source:
                    # Учет случая когда источник данных не определен
                    log.warning(u'Не определен источник данных в отчете')
                    return {'__fields__': list(), '__data__': list()}

                signature = data_source[:4].upper()
                if signature != DB_URL_SIGNATURE:
                    log.warning('Not support DB type <%s>' % signature)
                    return result
                # БД задается с помощью стандартного DB URL
                db_url = data_source[4:].lower().strip()

            log.info(u'DB connection <%s>' % db_url)
            # Установить связь с БД
            db_connection = sqlalchemy.create_engine(db_url)
            # Освободить БД
            # db_connection.dispose()
            log.info(u'DB SQL <%s>' % unicode(sql, 'utf-8'))
            sql_result = db_connection.execute(sql)
            rows = sql_result.fetchall()
            cols = rows[0].keys() if rows else []

            # Закрыть связь
            db_connection.dispose()
            db_connection = None

            # ТАБЛИЦА ЗАПРОСА ПРЕДСТАВЛЯЕТСЯ В ВИДЕ СЛОВАРЯ
            # {'__fields__':имена полей таблицы,'__data__':данные таблицы} !!!
            result = {'__fields__': cols, '__data__': list(rows)}
            return result
        except:
            if db_connection:
                # Закрыть связь
                db_connection.dispose()

            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка определения таблицы SQL запроса <%s>.' % sql)
            log.error(u'''ВНИМАНИЕ! Если возникает ошибка в модуле:
        ----------------------------------------------------------------------------------------------
        File "/usr/lib/python2.7/dist-packages/sqlalchemy/engine/default.py", line 324, in do_execute
            cursor.execute(statement, parameters)
        TypeError: 'dict' object does not support indexing
        ----------------------------------------------------------------------------------------------        
        Это означает что SQLAlchemy не может распарсить SQL выражение. 
        Необходимо вместо <%> использовать <%%> в SQL выражении. 
                    ''')

        return None

    def _isQueryFunc(self, Query_):
        """
        Определить представлен запрос в виде функции?
        @param Query_: Текст запроса.
        @return: True/False.
        """
        return Query_ and isinstance(Query_, str) and Query_.startswith(PY_SIGNATURE)

    def _execQueryFunc(self, Query_, vars=None):
        """
        Получить запрос из функции.
        @param Query_: Текст запроса.
        @param vars: Внешние переменные.
        @return: Возвращает запрос в разрешенном формате.
        """
        # Убрать сигнатуру определения функции
        func = Query_.replace(PY_SIGNATURE, '').strip()
        var_names = vars.keys() if vars else None
        log.debug(u'Выполнение функции: <%s>. Дополнительные переменные %s' % (func, var_names))
        return execfunc.exec_code(func, name_space=locals(), kwargs=vars)

    def _isEmptyQueryTbl(self, query_tbl):
        """
        Проверка пустой таблицы запроса.
        @param query_tbl: Словарь таблицы запроса.
        @return: True - пустая таблица запроса.
            False - есть данные.
        """
        if not query_tbl:
            return True
        # Есть переменные?
        elif isinstance(query_tbl, dict) and '__variables__' in query_tbl and query_tbl['__variables__']:
            return False
        # Есть координатные замены?
        elif isinstance(query_tbl, dict) and '__coord_fill__' in query_tbl and query_tbl['__coord_fill__']:
            return False
        # Есть данные табличной части?
        elif isinstance(query_tbl, dict) and '__data__' in query_tbl and query_tbl['__data__']:
            return False
        return True

    def getQueryTbl(self, report, db_url=None, sql=None, *args, **kwargs):
        """
        Получить таблицу запроса.
        @param report: Шаблон отчета.
        @param db_url: Connection string в виде url. Например
            postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
        @param sql: Запрос SQL.
            ВНИМАНИЕ! В начале SQL запроса должна стоять сигнатура:
            <SQL:> - Текст SQL запроса
            <PY:> -  Запрос задается функцией Python
            <PRG:> - Запрос задается внешней функцией Python.
        @return: Функция возвращает словарь -
            ТАБЛИЦА ЗАПРОСА ПРЕДСТАВЛЯЕТСЯ В ВИДЕ СЛОВАРЯ 
            {'__fields__':описания полей таблицы,'__data__':данные таблицы}
        """
        query = None
        try:
            if sql:
                query = sql
            else:
                # !!! ВНИМАНИЕ!!!
                # Проверка случая когда функция  запроса возвращает таблицу
                if isinstance(report['query'], dict):
                    return report['query']
                elif self._isQueryFunc(report['query']):
                    variables = kwargs.get('variables', None)
                    query = self._execQueryFunc(report['query'], vars=variables)

                    # Если метод возвращает уже сгенерированную таблицу запроса,
                    # то просто вернуть ее
                    if isinstance(query, dict):
                        if '__sql__' in query:
                            dataset_dict = self._getSQLQueryTable(report=report,
                                                                  db_url=db_url,
                                                                  sql=query['__sql__'])
                            if isinstance(dataset_dict, dict):
                                query.update(dataset_dict)
                        return query
                else:
                    query = report['query']
                # Обработка когда запрос вообще не определен
                if query is None:
                    log.warning(u'Запрос отчета не определен')
                    return None

                if query.startswith(SQL_SIGNATURE):
                    # Запрос задается SQL выражением
                    query = query.replace(SQL_SIGNATURE, u'').strip()
                elif query.startswith(CODE_SIGNATURE):
                    # Запрос задается функцией Python
                    query = execfunc.exec_code(query.replace(CODE_SIGNATURE, u'').strip())
                elif query.startswith(PY_SIGNATURE):
                    # Запрос задается функцией Python
                    query = execfunc.exec_code(query.replace(PY_SIGNATURE, u'').strip())
                else:
                    log.warning(u'Не указана сигнатура в запросе <%s>' % query)
                    return None

            query_tbl = None
            if not self._QueryTab:
                # Таблица запроса определена в виде SQL
                query_tbl = self._getSQLQueryTable(report, db_url=db_url, sql=query)
            else:
                # Если таблица запроса указана конкретно, то обработать ее
                # Указано имя таблицы запроса
                if isinstance(self._QueryTab, str):
                    if self._QueryTab[:4].upper() == SQL_SIGNATURE:
                        # Обработка обычного SQL запроса
                        query_tbl = self._getSQLQueryTable(report, sql=self._QueryTab[4:].strip())
                    else:
                        log.warning(u'Not support query type <%s>' % self._QueryTab)
                # Таблица уже просто определена как DataSet
                elif isinstance(self._QueryTab, dict):
                    query_tbl = self._QueryTab
                else:
                    log.warning(u'Not support query type <%s>' % type(self._QueryTab))
            return query_tbl
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка определения таблицы запроса <%s>.' % query)
        return None

    def RepSQLObj2SQLite(self, report, ResTab_):
        """
        Преобразование имен в шаблоне отчета в контексте SQLObject в имена в
            контексте sqlite.
        @param report: Щаблон отчета.
        @param ResTab_: Ресурсное описание таблиц.
        """
        try:
            rep = report
            # Для корректной обработки имен полей и таблиц они д.б.
            # отсортированны по убыванию длин имен классов данных
            data_class_names = ResTab_.keys()
            data_class_names.sort()
            data_class_names.reverse()

            # Обработка шаблона отчета
            # 1. Верхний и нижний колонтитулы
            # Перебор ячеек
            for row in rep['upper']:
                for cell in row:
                    if cell:
                        # Перебор классов
                        for data_class_name in data_class_names:
                            cell['value'] = ic.db.tabrestr.icNamesSQLObj2SQLite(cell['value'], data_class_name,
                                                                                ResTab_[data_class_name]['scheme'])
            # Перебор ячеек
            for row in rep['under']:
                for cell in row:
                    if cell:
                        # Перебор классов
                        for data_class_name in data_class_names:
                            cell['value'] = ic.db.tabrestr.icNamesSQLObj2SQLite(cell['value'], data_class_name,
                                                                                ResTab_[data_class_name]['scheme'])
            # 2. Лист шаблона
            # Перебор ячеек
            for row in rep['sheet']:
                for cell in row:
                    if cell:
                        # Перебор классов
                        for data_class_name in data_class_names:
                            cell['value'] = ic.db.tabrestr.icNamesSQLObj2SQLite(cell['value'], data_class_name,
                                                                                ResTab_[data_class_name]['scheme'])
            # 3. Описание групп
            for grp in rep['groups']:
                for data_class_name in data_class_names:
                    grp['field'] = ic.db.tabrestr.icNamesSQLObj2SQLite(grp['field'], data_class_name,
                                                                       ResTab_[data_class_name]['scheme'])

            return rep
        except:
            return report

    def PreviewResult(self, report_data=None):
        """
        Предварительный просмотр.
        @param report_data: Сгенерированный отчет.
        """
        return

    def PrintResult(self, report_data=None):
        """
        Печать.
        @param report_data: Сгенерированный отчет.
        """
        return

    def ConvertResult(self, report_data=None, to_file=None):
        """
        Конвертирование результатов отчета.
        @param report_data: Сгенерированный отчет.
        @param to_file: Имя результирующего файла.
        """
        return

    def save(self, report_data=None):
        """
        Сохранить результаты генерации в файл
        @param report_data: Сгенерированный отчет.
        @return: Имя сохраненного файла или None, если сохранения не произошло.
        """
        return None
