#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль системы генератора отчетов ReportManager.
"""

# --- Подключение библиотек ---
import os
import os.path

# PyReportManager - the ActiveX wrapper code
try:
    from ic.report import reportman
except:
    print('PyReportManager Import Error')

from ic.report import icrepgensystem
from ic.std.log import log

__version__ = (0, 0, 1, 4)

# --- Константы подсистемы ---
DEFAULT_REP_FILE_NAME = 'c:/temp/new_report.rep'


# --- Описания классов ---
class icReportManagerGeneratorSystem(icrepgensystem.icReportGeneratorSystem):
    """
    Класс системы генерации отчетов ReportManager.
    """

    def __init__(self, Rep_=None, ParentForm_=None):
        """
        Конструктор класса.
        @param Rep_: Шаблон отчета.
        @param ParentForm_: Родительская форма, необходима для вывода сообщений.
        """
        # вызов конструктора предка
        icrepgensystem.icReportGeneratorSystem.__init__(self, Rep_, ParentForm_)

        # Имя файла шаблона отчета
        self.RepTmplFileName = None
        
        # Папка отчетов.
        self._report_dir = None
        if self._ParentForm:
            self._report_dir = os.path.abspath(self._ParentForm.GetReportDir())
        
    def getReportDir(self):
        """
        Папка отчетов.
        """
        return self._report_dir

    def Preview(self, Rep_=None):
        """
        Предварительный просмотр.
        @param Rep_: Полное описание шаблона отчета.
        """
        if Rep_ is None:
            Rep_ = self._Rep
        # Создание связи с ActiveX
        report_dir = os.path.abspath(ic.engine.ic_user.icGet('REPORT_DIR'))
        report_file = os.path.abspath(Rep_['generator'], report_dir)
        try:
            report = reportman.ReportMan(report_file)
            # Параметры отчета
            params = self._getReportParameters(Rep_)
            # Установка параметров
            self._setReportParameters(report, params)
            # Вызов предварительного просмотра
            report.preview(u'Предварительный просмотр: %s' % report_file)
        except:
            log.fatal(u'Ошибка передварительного просмотра отчета %s' % report_file)
            
    def Print(self, Rep_=None):
        """
        Печать.
        @param Rep_: Полное описание шаблона отчета.
        """
        if Rep_ is None:
            Rep_ = self._Rep
        # Создание связи с ActiveX
        report_dir = os.path.abspath(ic.engine.ic_user.icGet('REPORT_DIR'))
        report_file = os.path.abspath(Rep_['generator'], report_dir)
        try:
            report = reportman.ReportMan(report_file)
            # Параметры отчета
            params = self._getReportParameters(Rep_)
            # Установка параметров
            self._setReportParameters(report, params)
            # Вызов предварительного просмотра
            report.printout(u'Печать: %s' % report_file, True, True)
        except:
            log.fatal(u'Ошибка печати отчета %s' % report_file)

    def PageSetup(self):
        """
        Установка параметров страницы.
        """
        pass

    def Convert(self, Rep_=None, ToXLSFile_=None, *args, **kwargs):
        """
        Вывод результатов отчета в Excel.
        @param Rep_: Полное описание шаблона отчета.
        @param ToXLSFile_: Имя файла, куда необходимо сохранить отчет.
        """
        if Rep_ is None:
            Rep_ = self._Rep
        # Создание связи с ActiveX
        report_dir = os.path.abspath(ic.engine.ic_user.icGet('REPORT_DIR'))
        report_file = os.path.abspath(Rep_['generator'], report_dir)
        try:
            report = reportman.ReportMan(report_file)
            # Параметры отчета
            params = self._getReportParameters(Rep_)
            # Установка параметров
            self._setReportParameters(report, params)
            # Вызов предварительного просмотра
            report.execute()
        except:
            log.fatal(u'Ошибка конвертирования отчета %s' % report_file)

    def Edit(self, RepFileName_=None):
        """
        Редактирование отчета.
        @param RepFileName_: Полное имя файла шаблона отчета.
        """
        # Создание связи с ActiveX
        rprt_file_name = os.path.abspath(RepFileName_)
        rep = ic_res.LoadResource(rprt_file_name)
        report_dir = os.path.abspath(ic.engine.ic_user.icGet('REPORT_DIR'))
        rep_file = os.path.abspath(rep['generator'], report_dir)
        
        reportman_designer_key = ic_util.GetRegValue('Software\\Classes\\Report Manager Designer\\shell\\open\\command', None)
        if reportman_designer_key:
            reportman_designer_run = reportman_designer_key.replace('\'%1\'', '\'%s\'') % rep_file
            cmd = 'start %s' % reportman_designer_run
            log.debug(u'Запуск команды ОС: <%s>' % cmd)
            # и запустить Report Manager Designer
            os.system(cmd)
        else:
            msg = u'Не определен дизайнер отчетов Report Manager Designer %s' % reportman_designer_key
            log.warning(msg)
            ic_dlg.icWarningBox(u'ВНИМАНИЕ!', msg)

        # Определить файл *.xml
        xml_file = os.path.normpath(os.path.abspath(os.path.splitext(RepFileName_)[0]+'.xml'))
        cmd = 'start excel.exe \'%s\'' % xml_file
        log.debug(u'Запуск команды ОС: <%s>' % cmd)
        # и запустить MSExcel
        os.system(cmd)

    def _getReportParameters(self, Rep_=None):
        """
        Запустить генератор отчета.
        @param Rep_: Шаблон отчета.
        @return: Возвращает словарь параметров. {'Имя параметра отчета':Значение параметра отчета}.
        """
        try:
            if Rep_ is not None:
                self._Rep = Rep_
            else:
                Rep_ = self._Rep

            # 1. Получить параметры запроса отчета
            query = Rep_['query']
            if query is not None:
                if self._isQueryFunc(query):
                    query = self._execQueryFunc(query)
                else:
                    query = ic_exec.ExecuteMethod(query, self)
                
            return query
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка определения параметров отчета %s.' % Rep_['name'])
            return None

    def _setReportParameters(self, ReportObj_, Parameters_):
        """
        Установить параметры для отчета.
        @param ReportObj_: Объект отчета ReportManger.
        @param Parameters_: Словарь параметров. {'Имя параметра отчета':Значение параметра отчета}.
        @return: Возвращает результат выполнения операции.
        """
        try:
            if Parameters_:
                for param_name, param_value in Parameters_.items():
                    ReportObj_.set_param(param_name, param_value)
            return True
        except:
            log.fatal(u'Ошибка установки параметров %s в отчете %s' % (Parameters_, ReportObj_._report_filename))
            return False
