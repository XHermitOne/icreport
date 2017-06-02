#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль функций запуска генератора отчетов.
"""

import os
import os.path
import wx

# Подключение библиотек
from ic.std.log import log
from ic.std.utils import res

from ic.report import icreportbrowser
from ic.report import report_generator
from ic.report import icstylelib


__version__ = (0, 0, 1, 1)

DEFAULT_REPORT_FILE_EXT = '.rprt'

# Функции управления


def getReportResourceFilename(report_filename='', report_dir=''):
    """
    Получить полное имя файла шаблона отчета.
    @param report_filename: Имя файла отчета в кратком виде.
    @param report_dir: Папка отчетов.
    @return: Полное имя файла отчета.
    """
    # Проверить расширение
    if not report_filename.endswith(DEFAULT_REPORT_FILE_EXT):
        report_filename = os.path.splitext(report_filename)[0]+DEFAULT_REPORT_FILE_EXT

    if os.path.exists(report_filename):
        # Проверить может быть задано абсолютное имя файла
        filename = report_filename
    else:
        # Задано скорее всего относительное имя файла
        # относительно папки отчетов
        filename = os.path.join(report_dir, report_filename)
        if not os.path.exists(filename):
            # Нет такого файла
            log.warning(u'Файл шаблона отчета <%s> не найден' % filename)
            filename = None
    return filename


def loadStyleLib(stylelib_filename=None):
    """
    Загрузить библиотеку стилей из файла.
    @param stylelib_filename: Файл библиотеки стилей.
    @return: Библиотека стилей.
    """
    # Загрузить библлиотеку стилей из файла
    stylelib = None
    if stylelib_filename:
        stylelib_filename = os.path.abspath(stylelib_filename)
        if os.path.exists(stylelib_filename):
            xml_stylelib = icstylelib.icXMLRepStyleLib()
            stylelib = xml_stylelib.convert(stylelib_filename)
    return stylelib


def ReportBrowser(parent_form=None, report_dir='', mode=icreportbrowser.IC_REPORT_EDITOR_MODE):
    """
    Запуск браузера отчетов.
    @param parent_form: Родительская форма, если не указана, 
        то создается новое приложение.
    @param report_dir: Директорий, где хранятся отчеты.
    @return: Возвращает результат выполнения операции True/False.
    """
    try:
        app = None
        if parent_form is None:
            app = wx.PySimpleApp()

        # Иначе вывести окно выбора отчета
        rep_gen_dlg = icreportbrowser.icReportBrowserDialog(parent_form, mode,
                                                            report_dir=report_dir)
        rep_gen_dlg.ShowModal()

        return True
    except:
        log.fatal(u'Report browser')


def ReportEditor(parent_form=None, report_dir=''):
    """
    Запуск редактора отчетов. Редактор - режим работы браузера.
    @param parent_form: Родительская форма, если не указана, 
        то создается новое приложение.
    @param report_dir: Директорий, где хранятся отчеты.
    @return: Возвращает результат выполнения операции True/False.
    """
    return ReportBrowser(parent_form, report_dir, icreportbrowser.IC_REPORT_EDITOR_MODE)


def ReportViewer(parent_form=None, report_dir=''):
    """
    Запуск просмотрщика отчетов. Просмотрщик - режим работы браузера.
    @param parent_form: Родительская форма, если не указана, 
        то создается новое приложение.
    @param report_dir: Директорий, где хранятся отчеты.
    @return: Возвращает результат выполнения операции True/False.
    """
    return ReportBrowser(parent_form, report_dir, icreportbrowser.IC_REPORT_VIEWER_MODE)


def DoReport(report_filename='', report_dir='', parent_form=None):
    """
    Функция запускает генератор отчетов.
    @param report_filename: Файл отчета.
    @param report_dir: Директорий, где хранятся отчеты.
    @param parent_form: Родительская форма, если не указана, 
        то создается новое приложение.
    @return: Возвращает результат выполнения операции True/False.
    """
    try:
        if not parent_form:
            # Создать объект приложения
            # !!! ВНИМАНИЕ: Чтобы закрывалась консоль питона надо
            # поставить вместо нуля единицу
            #                            |
            #                            V
            application = icRepGenSysApp(0)
            application.SetReportDir(report_dir)
            if not report_filename:
                # Иначе вывести окно выбора отчета
                application.RunReportGenerator()
            else:
                # Если определен отчет, то запустить на выполнение
                ic_repgen_sys = rg_sys.icRepGenSystem()
                ic_repgen_sys.icGenerateReport(report_filename)
            # Запуск главного цикла программы
            application.MainLoop() 
        else:
            # Если форма родительская определена, то приложение не создавать
            if not report_filename:
                # Иначе вывести окно выбора отчета
                rep_gen_form = icRepGenSysDialog(parent_form)
                # rep_gen_form.SetReportDir(report_dir)
                rep_gen_form.ShowModal()
            else:
                # Если определен отчет, то запустить на выполнение
                ic_repgen_sys = rg_sys.icRepGenSystem(parent_form=parent_form)
                ic_repgen_sys.icGenerateReport(report_filename)
        return True
    except:
        log.fatal(u'Do report')


def DoReportExcel(report_filename, QueryTable_=None, ToXLSFile_=None, XLSSheetName_=None, parent_form=None):
    """
    Функция запускает генератор отчетов для конвертации в Excel.
    @param report_filename: Файл отчета.
    @param QueryTable_: Таблица запроса.
    @param ToXLSFile_: Имя файла, куда необходимо сохранить отчет.
    @param XLSSheetName_: Имя листа.
    @param parent_form: Родительская форма, необходима для вывода сообщений.
    @return: Возвращает результат выполнения операции True/False.
    """
    try:
        i_rg_sys.icGetRepGenSys(report_filename,
                                QueryTable_, parent_form).Export(None, ToXLSFile_, XLSSheetName_)
        return True
    except:
        log.fatal(u'Ошибка вывода отчета %s в Microsoft Excel.' % report_filename)


def ReportPrint(parent_form=None, report_filename='', report_dir='',
                db_url=None, sql=None, command=None,
                stylelib_filename=None, variables=None):
    """
    Функция запускает генератор отчетов и вывод на печать.
    @param parent_form: Родительская форма, если не указана,
        то создается новое приложение.
    @param report_filename: Файл отчета.
    @param report_dir: Директорий, где хранятся отчеты.
    @param db_url: Connection string в виде url. Например
        postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
    @param sql: Запрос SQL.
    @param command: Комманда после генерации. print/preview/export.
    @param stylelib_filename: Файл библиотеки стилей.
    @param variables: Словарь переменных для заполнения отчета.
    @return: Возвращает результат выполнения операции True/False.
    """
    report_filename = getReportResourceFilename(report_filename, report_dir)
    try:
        if not report_filename:
            return ReportViewer(parent_form, report_dir)
        else:
            stylelib = loadStyleLib(stylelib_filename)
            # Если определен отчет, то запустить на выполнение
            repgen_system = report_generator.getReportGeneratorSystem(report_filename, parent_form)
            return repgen_system.Print(res.loadResourceFile(report_filename),
                                       stylelib=stylelib,
                                       variables=variables)
        return False
    except:
        log.fatal(u'Report print <%s>' % report_filename)


def ReportPreview(parent_form=None, report_filename='', report_dir='',
                  db_url=None, sql=None, command=None,
                  stylelib_filename=None, variables=None):
    """
    Функция запускает генератор отчетов и вывод на экран предварительного просмотра.
    @param parent_form: Родительская форма, если не указана,
        то создается новое приложение.
    @param report_filename: Файл отчета.
    @param report_dir: Директорий, где хранятся отчеты.
    @param db_url: Connection string в виде url. Например
        postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
    @param sql: Запрос SQL.
    @param command: Комманда после генерации. print/preview/export.
    @param stylelib_filename: Файл библиотеки стилей.
    @param variables: Словарь переменных для заполнения отчета.
    @return: Возвращает результат выполнения операции True/False.
    """
    report_filename = getReportResourceFilename(report_filename, report_dir)
    try:
        if not report_filename:
            return ReportViewer(parent_form, report_dir)
        else:
            stylelib = loadStyleLib(stylelib_filename)
            # Если определен отчет, то запустить на выполнение
            repgen_system = report_generator.getReportGeneratorSystem(report_filename, parent_form)
            return repgen_system.Preview(res.loadResourceFile(report_filename),
                                         stylelib=stylelib,
                                         variables=variables)
        return False
    except:
        log.fatal(u'Report preview <%s>' % report_filename)


def ReportExport(parent_form=None, report_filename='', report_dir='',
                 db_url=None, sql=None, command=None,
                 stylelib_filename=None, variables=None):
    """
    Функция запускает генератор отчетов и вывод в Office.
    @param parent_form: Родительская форма, если не указана,
        то создается новое приложение.
    @param report_filename: Файл отчета.
    @param report_dir: Директорий, где хранятся отчеты.
    @param db_url: Connection string в виде url. Например
        postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
    @param sql: Запрос SQL.
    @param command: Комманда после генерации. print/preview/export.
    @param stylelib_filename: Файл библиотеки стилей.
    @param variables: Словарь переменных для заполнения отчета.
    @return: Возвращает результат выполнения операции True/False.
    """
    report_filename = getReportResourceFilename(report_filename, report_dir)
    try:
        if not report_filename:
            return ReportViewer(parent_form, report_dir)
        else:
            stylelib = loadStyleLib(stylelib_filename)
            # Если определен отчет, то запустить на выполнение
            repgen_system = report_generator.getReportGeneratorSystem(report_filename, parent_form)
            return repgen_system.Convert(res.loadResourceFile(report_filename),
                                         stylelib=stylelib,
                                         variables=variables)
        return False
    except:
        log.fatal(u'Report export <%s>' % report_filename)


def ReportSelect(parent_form=None, report_filename='', report_dir='',
                 db_url=None, sql=None, command=None,
                 stylelib_filename=None, variables=None):
    """
    Функция запускает генератор отчетов с последующим выбором действия.
    @param parent_form: Родительская форма, если не указана,
        то создается новое приложение.
    @param report_filename: Файл отчета.
    @param report_dir: Директорий, где хранятся отчеты.
    @param db_url: Connection string в виде url. Например
        postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
    @param sql: Запрос SQL.
    @param command: Комманда после генерации. print/preview/export.
    @param stylelib_filename: Файл библиотеки стилей.
    @param variables: Словарь переменных для заполнения отчета.
    @return: Возвращает результат выполнения операции True/False.
    """
    report_filename = getReportResourceFilename(report_filename, report_dir)
    try:
        if not report_filename:
            return ReportViewer(parent_form, report_dir)
        else:
            stylelib = loadStyleLib(stylelib_filename)
            # Если определен отчет, то запустить на выполнение
            repgen_system = report_generator.getReportGeneratorSystem(report_filename, parent_form)
            return repgen_system.selectAction(res.loadResourceFile(report_filename),
                                              stylelib=stylelib,
                                              variables=variables)
        return False
    except:
        log.fatal(u'Report export <%s>' % report_filename)


# Комманды пост обработки сгенерированног отчета
DO_COMMAND_PRINT = 'print'
DO_COMMAND_PREVIEW = 'preview'
DO_COMMAND_EXPORT = 'export'
DO_COMMAND_SELECT = 'select'


def doReport(parent_form=None, report_filename='', report_dir='', db_url='', sql='', command=None,
             stylelib_filename=None, variables=None):
    """
    Функция запускает генератор отчетов.
    @param parent_form: Родительская форма, если не указана,
        то создается новое приложение.
    @param report_filename: Файл отчета.
    @param report_dir: Директорий, где хранятся отчеты.
    @param db_url: Connection string в виде url. Например
        postgresql+psycopg2://postgres:postgres@10.0.0.3:5432/realization.
    @param sql: Запрос SQL.
    @param command: Комманда после генерации. print/preview/export.
    @param stylelib_filename: Файл библиотеки стилей.
    @param variables: Словарь переменных для заполнения отчета.
    @return: Возвращает результат выполнения операции True/False.
    """
    try:
        app = None
        if parent_form is None:
            app = wx.PySimpleApp()

        if not report_filename:
            return ReportViewer(parent_form, report_dir)
        else:
            # Если определен отчет, то запустить на выполнение
            repgen_system = report_generator.getReportGeneratorSystem(report_filename, parent_form)
            stylelib = loadStyleLib(stylelib_filename)

            data = repgen_system.generate(res.loadResourceFile(report_filename), db_url, sql,
                                          stylelib=stylelib, vars=variables)

            if command:
                command = command.lower()
                if command == DO_COMMAND_PRINT:
                    repgen_system.PrintResult(data)
                elif command == DO_COMMAND_PREVIEW:
                    repgen_system.PreviewResult(data)
                elif command == DO_COMMAND_EXPORT:
                    repgen_system.ConvertResult(data)
                elif command == DO_COMMAND_SELECT:
                    repgen_system.doSelectAction(data)
                else:
                    log.warning(u'Not define command Report System <%s>' % command)
            else:
                repgen_system.save(data)
        return True
    except:
        log.fatal(u'Do report <%s>' % report_filename)


if __name__ == '__main__':
    DoReport()
