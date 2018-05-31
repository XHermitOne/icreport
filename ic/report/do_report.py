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
from ic.std.utils import textfunc
from ic.std.utils import filefunc

from ic.report import icreportbrowser
from ic.report import report_generator
from ic.report import icstylelib


__version__ = (0, 0, 3, 1)

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
    rprt_filename = report_filename
    if not rprt_filename.endswith(DEFAULT_REPORT_FILE_EXT):
        rprt_filename = os.path.splitext(rprt_filename)[0]+DEFAULT_REPORT_FILE_EXT

    # Проверить актуальность шаблона
    full_src_filename = getPathFilename(report_filename, report_dir)
    full_rprt_filename = getPathFilename(rprt_filename, report_dir)
    if isNewReportTemplateFile(full_src_filename, full_rprt_filename):
        # Если исходный шаблон изменен позже чем рабочий файл шаблона <rprt>
        # то необходимо сделать изменения
        updateReportTemplateFile(full_src_filename, full_rprt_filename)

    if os.path.exists(rprt_filename):
        # Проверить может быть задано абсолютное имя файла
        filename = rprt_filename
    else:
        # Задано скорее всего относительное имя файла
        # относительно папки отчетов
        filename = full_rprt_filename
        if not os.path.exists(filename):
            # Нет такого файла
            log.warning(u'Файл шаблона отчета <%s> не найден' % textfunc.toUnicode(filename))
            filename = createReportResourceFile(filename)
    log.debug(u'Полное имя файла шаблона <%s>' % textfunc.toUnicode(filename))
    return filename


def getPathFilename(filename='', report_dir=''):
    """
    Получить полное имя файла отчета.
    @param report_filename: Имя файла отчета в кратком виде.
    @param report_dir: Папка отчетов.
    @return: Полное имя файла отчета.
    """
    return filefunc.normal_path(os.path.join(report_dir, filename))


def isNewReportTemplateFile(src_filename, rprt_filename):
    """
    Проверить актуальность шаблона.
    Если исходный шаблон изменен позже чем рабочий файл шаблона *.rprt
    то необходимо сделать изменения.
    @return: True-внесены измененияв исходный шаблон/False-изменений нет.
    """
    src_modify_dt = filefunc.file_modify_dt(src_filename)
    rprt_modify_dt = filefunc.file_modify_dt(rprt_filename)
    if src_modify_dt and rprt_modify_dt:
        return src_modify_dt > rprt_modify_dt
    return False


def updateReportTemplateFile(src_filename, rprt_filename):
    """
    Произвести обновления шаблона отчета.
    @param src_filename: Имя файла шаблона источника.
    @param rprt_filename: Имя результирующего файла шаблона *.rprt.
    @return: Скорректированное имя созданного файла шаблона или None в случае ошибки.
    """
    # Удаляем результирующий файл
    filefunc.remove_file(rprt_filename)
    # Удаляем все промежуточные файлы
    for ext in report_generator.SRC_REPORT_EXT:
        src_ext = os.path.splitext(src_filename)[1].lower()
        if src_ext != ext:
            filename = os.path.splitext(src_filename)[0] + ext
            if os.path.exists(filename):
                filefunc.remove_file(filename)
    # Пересоздаем шаблон
    return createReportResourceFile(rprt_filename)


def createReportResourceFile(template_filename):
    """
    Создать ресурсный файл шаблона по имени запрашиваемого.
    @param template_filename: Имя запрашиваемого файла шаблона.
    @return: Скорректированное имя созданного файла шаблона или None в случае ошибки.
    """
    # Коррекция имени файла с учетом русских букв в имени файла
    dir_name = os.path.dirname(template_filename)
    base_filename = os.path.basename(template_filename).replace(' ', '_')
    base_filename = textfunc.rus2lat(base_filename) if textfunc.isRUSText(base_filename) else base_filename
    norm_tmpl_filename = os.path.join(dir_name, base_filename)

    log.info(u'Создание нового файла шаблона <%s>' % norm_tmpl_filename)
    # Последовательно проверяем какой файл можно взять за основу для шаблона
    for ext in report_generator.SRC_REPORT_EXT:
        src_filename = os.path.splitext(template_filename)[0] + ext
        unicode_src_filename = textfunc.toUnicode(src_filename)
        if os.path.exists(src_filename):
            # Да такой файл есть и он может выступать
            # в качестве источника для шаблона
            log.info(u'Найден источник шаблона отчета <%s>' % unicode_src_filename)
            try:
                rep_generator = report_generator.createReportGeneratorSystem(ext)
                return rep_generator.Update(src_filename)
            except:
                log.fatal(u'Ошибка конвертации шаблона отчета <%s> -> <%s>' % (unicode_src_filename, norm_tmpl_filename))
            return None

    log.warning(u'Не найдены источники шаблонов отчета в папке <%s> для <%s>' % (dir_name,
                                                                                 textfunc.toUnicode(os.path.basename(template_filename))))
    return None


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
