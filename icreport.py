#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
icReport - Программа запуска/обзора отчетов.

Параметры коммандной строки:
    
    python icreport.py <Параметры запуска>
    
Параметры запуска:

    [Помощь и отладка]
        --help|-h|-?        Напечатать строки помощи
        --version|-v        Напечатать версию программы
        --debug|-d          Режим отладки
        --log|-l            Режим журналирования

   [Режимы запуска]
        --viewer|-V         Режим выбора отчета и отправки его на печать
        --editor|-E         Режим с поддержкой вызова редактора отчета
        --print=|-p=        Режим запуска формирования отчета с выводом на печать
        --preview=|-P=      Режим запуска формирования отчета с предварительным просмотром
        --export=|-E=       Режим запуска формирования отчета с последующей конвертацией в office
        --select=|-S=       Режим запуска формирования отчета с последующим выбором действия.
        --gen=              Режим запуска формирования отчета с дополнительными параметрами
        --db=               Указание связи с БД (в виде url)
        --sql=              Указание запроса SQL для получения таблицы отчета
        --postprint         Распечатать отчет после генерации
        --postpreview       Предварительный просмотр отчета после генерации
        --postexport        Конвертация в Office отчета после генерации
        --stylelib=         Указание библиотеки стилей для единого оформления отчетов
        --var=              Добавление переменной для заполнения в отчете
        --path=             Указание папки отчетов
"""


import sys
import os
import os.path
import getopt
import wx

from ic import config
from ic.std.log import log

from ic.report import do_report

__version__ = (0, 0, 7, 5)

DEFAULT_REPORTS_PATH = './reports'


def main(argv):
    """
    Основная запускающая функция.
    @param argv: Список параметров коммандной строки.
    """
    # Разбираем аргументы командной строки
    try:
        options, args = getopt.getopt(argv, 'h?vdVEDpPES',
                                      ['help', 'version', 'debug',
                                       'viewer', 'editor',
                                       'postprint', 'postpreview', 'postexport',
                                       'print=', 'preview=', 'export=', 'select=',
                                       'gen=', 'db=', 'sql=',
                                       'stylelib=', 'var=', 'path='])
    except getopt.error, msg:
        print(msg)
        print('For help use --help option')
        sys.exit(2)

    # Инициализоция системы журналирования
    log.init(config)

    # Параметры запуска генерации отчета из коммандной строки
    report_filename = None
    db = None
    sql = None
    do_cmd = None
    stylelib = None
    vars = dict()
    path = None
    mode = 'default'
    mode_arg = None

    for option, arg in options:
        if option in ('-h', '--help', '-?'):
            print(__doc__)
            sys.exit(0)   
        elif option in ('-v', '--version'):
            print('icReport version: %s' % '.'.join([str(ver) for ver in __version__]))
            sys.exit(0)
        elif option in ('-d', '--debug'):
            config.set_glob_var('DEBUG_MODE', True)
        elif option in ('-l', '--log'):
            config.set_glob_var('LOG_MODE', True)
        elif option in ('-V', '--viewer'):
            mode = 'view'
        elif option in ('-E', '--editor'):
            mode = 'edit'
        elif option in ('-p', '--print'):
            mode = 'print'
            mode_arg = arg
        elif option in ('-P', '--preview'):
            mode = 'preview'
            mode_arg = arg
        elif option in ('-E', '--export'):
            mode = 'export'
            mode_arg = arg
        elif option in ('-S', '--select'):
            mode = 'select'
            mode_arg = arg
        elif option in ('--gen',):
            report_filename = arg
        elif option in ('--db',):
            db = arg
        elif option in ('--sql',):
            sql = arg
        elif option in ('--postprint',):
            do_cmd = do_report.DO_COMMAND_PRINT
        elif option in ('--postpreview',):
            do_cmd = do_report.DO_COMMAND_PREVIEW
        elif option in ('--postexport',):
            do_cmd = do_report.DO_COMMAND_EXPORT
        elif option in ('--stylelib',):
            stylelib = arg
        elif option in ('--var',):
            var_name = arg.split('=')[0].strip()
            var_value = arg.split('=')[-1].strip()
            vars[var_name] = var_value
            log.debug(u'Дополнительная переменная <%s>. Значение [%s]' % (unicode(var_name, config.DEFAULT_ENCODING) ,
                                                                          unicode(var_value, config.DEFAULT_ENCODING)))
        elif option in ('--path',):
            path = arg

    # ВНИМАНИЕ! Небходимо добавить путь к папке отчетов,
    # чтобы проходили импорты модулей отчетов
    if path is None:
        path = DEFAULT_REPORTS_PATH
    if os.path.exists(path) and os.path.isdir(path) and path not in sys.path:
        sys.path.append(path)

    # Внимание! Приложение создается для
    # управления диалоговыми окнами отчетов
    app = wx.PySimpleApp()
    # ВНИМАНИЕ! Выставить русскую локаль
    # Это необходимо для корректного отображения календарей,
    # форматов дат, времени, данных и т.п.
    locale = wx.Locale()
    locale.Init(wx.LANGUAGE_RUSSIAN)

    if mode == 'default':
        if report_filename:
            # Запустить генерацию отчета из комадной строки
            do_report.doReport(report_filename=report_filename, report_dir=path,
                               db_url=db, sql=sql, command=do_cmd,
                               stylelib_filename=stylelib, variables=vars)
    elif mode == 'view':
        do_report.ReportViewer(report_dir=path)
    elif mode == 'edit':
        do_report.ReportEditor(report_dir=path)
    elif mode == 'print':
        do_report.ReportPrint(report_filename=mode_arg, report_dir=path,
                              db_url=db, sql=sql, command=do_cmd,
                              stylelib_filename=stylelib, variables=vars)
    elif mode == 'preview':
        do_report.ReportPreview(report_filename=mode_arg, report_dir=path,
                                db_url=db, sql=sql, command=do_cmd,
                                stylelib_filename=stylelib, variables=vars)
    elif mode == 'export':
        do_report.ReportExport(report_filename=mode_arg, report_dir=path,
                               db_url=db, sql=sql, command=do_cmd,
                               stylelib_filename=stylelib, variables=vars)
    elif mode == 'select':
        do_report.ReportSelect(report_filename=mode_arg, report_dir=path,
                               db_url=db, sql=sql, command=do_cmd,
                               stylelib_filename=stylelib, variables=vars)

    app.MainLoop()


if __name__ == '__main__':
    main(sys.argv[1:])