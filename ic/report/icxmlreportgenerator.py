# !/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль системы генератора отчетов, основанные на генерации XML файлов.
"""

# Подключение библиотек
import os 
import os.path

from ic.std.log import log
from ic.std.dlg import dlg

from ic.report import icrepgensystem
from ic.report import icreptemplate
from ic.report import icrepgen
from ic.report import icrepfile

__version__ = (0, 0, 1, 3)


class icXMLReportGeneratorSystem(icrepgensystem.icReportGeneratorSystem):
    """
    Класс системы генерации отчетов, основанные на генерации XML файлов.
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
        
    def reloadRepData(self, RepTmplFileName_=None):
        """
        Перегрузить данные отчета.
        @param RepTmplFileName_: Имя файла шаблона отчета.
        """
        if RepTmplFileName_ is None:
            RepTmplFileName_ = self.RepTmplFileName
        icrepgensystem.icReportGeneratorSystem.reloadRepData(self, RepTmplFileName_)
        
    def getReportDir(self):
        """
        Папка отчетов.
        """
        if self._report_dir is None:
            if self._ParentForm:
                self._report_dir = os.path.abspath(self._ParentForm.GetReportDir())
            else:
                dlg.getMsgBox(u'Ошибка', u'Не определена папка отчетов!')
                                
        return self._report_dir

    def _genXMLReport(self, Rep_):
        """
        Генерация отчета и сохранение его в XML файл.
        @param Rep_: Полное описание шаблона отчета.
        @return: Возвращает имя xml файла или None в случае ошибки.
        """
        if Rep_ is None:
            Rep_ = self._Rep
        data_rep = self.GenerateReport(Rep_)
        if data_rep:
            rep_file = icrepfile.icExcelXMLReportFile()
            rep_file_name = self.getReportDir()+'/%s_report_result.xml' % str(data_rep['name'])
            rep_file.write(rep_file_name, data_rep)
            log.info(u'Сохранение отчета в файл <%s>' % rep_file_name)
            return rep_file_name
        return None
        
    def Preview(self, Rep_=None):
        """
        Предварительный просмотр.
        @param Rep_: Полное описание шаблона отчета.
        """
        xml_rep_file_name = self._genXMLReport(Rep_)
        if xml_rep_file_name:
            # Открыть excel в режиме просмотра
            self.PreviewExcel(xml_rep_file_name)
            
    def PreviewExcel(self, XMLFileName_):
        """
        Открыть excel  в режиме предварительного просмотра.
        @param XMLFileName_: Имя xml файла, содержащего сгенерированный отчет.
        """
        try:
            # Установить связь с Excel
            excel_app = win32com.client.Dispatch('Excel.Application')
            # Скрыть Excel
            excel_app.Visible = 0
            # Закрыть все книги
            excel_app.Workbooks.Close()
            # Открыть
            rep_tmpl_book = excel_app.Workbooks.Open(XMLFileName_)
            # Показать Excel
            excel_app.Visible = 1
            
            excel_app.ActiveWindow.ActiveSheet.PrintPreview()
            return True
        except pythoncom.com_error:
            # Вывести сообщение об ошибке в лог
            log.fatal()
            return False

    def Print(self, Rep_=None):
        """
        Печать.
        @param Rep_: Полное описание шаблона отчета.
        """
        xml_rep_file_name = self._genXMLReport(Rep_)
        if xml_rep_file_name:
            # Открыть печать в excel
            self.PrintExcel(xml_rep_file_name)

    def PrintExcel(self, XMLFileName_):
        """
        Печать отчета с помощью excel.
        @param XMLFileName_: Имя xml файла, содержащего сгенерированный отчет.
        """
        try:
            # Установить связь с Excel
            excel_app = win32com.client.Dispatch('Excel.Application')
            # Скрыть Excel
            excel_app.Visible = 0
            # Закрыть все книги
            excel_app.Workbooks.Close()
            # Открыть
            rep_tmpl_book = excel_app.Workbooks.Open(XMLFileName_)
            # Показать Excel
            excel_app.Visible = 1
            return True
        except pythoncom.com_error:
            # Вывести сообщение об ошибке в лог
            log.fatal()
            return False
            
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
        xml_rep_file_name = self._genXMLReport(Rep_)
        if xml_rep_file_name:
            # Excel
            self.OpenExcel(xml_rep_file_name)

    def OpenExcel(self, XMLFileName_):
        """
        Открыть excel.
        @param XMLFileName_: Имя xml файла, содержащего сгенерированный отчет.
        """
        try:
            # Установить связь с Excel
            excel_app = win32com.client.Dispatch('Excel.Application')
            # Скрыть Excel
            excel_app.Visible = 0
            # Закрыть все книги
            excel_app.Workbooks.Close()
            # Открыть
            rep_tmpl_book = excel_app.Workbooks.Open(XMLFileName_)
            # Показать Excel
            excel_app.Visible = 1
            return True
        except pythoncom.com_error:
            # Вывести сообщение об ошибке в лог
            log.fatal()
            return False

    def Edit(self, RepFileName_=None):
        """
        Редактирование отчета.
        @param RepFileName_: Полное имя файла шаблона отчета.
        """
        # Определить файл *.xml
        xml_file = os.path.abspath(os.path.splitext(RepFileName_)[0]+'.xml')
        cmd = 'start excel.exe \"%s\"' % xml_file
        # и запустить MSExcel
        os.system(cmd)

    def GenerateReport(self, Rep_=None):
        """
        Запустить генератор отчета.
        @param Rep_: Шаблон отчета.
        @return: Возвращает сгенерированный отчет или None в случае ошибки.
        """
        try:
            if Rep_ is not None:
                self._Rep = Rep_

            # 1. Получить таблицу запроса
            query_tbl = self.getQueryTbl(self._Rep)
            if not query_tbl or not query_tbl['__data__']:
                if dlg.getAskBox(u'Внимание',
                                 u'Нет данных, соответствующих запросу: %s. Продолжить генерацию отчета?' % self._Rep['query']):
                    return None

            # 2. Запустить генерацию
            rep = icrepgen.icReportGenerator()
            data_rep = rep.generate(self._Rep, query_tbl)

            return data_rep
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации отчета <%s>.' % self._Rep['name'])
        return None
