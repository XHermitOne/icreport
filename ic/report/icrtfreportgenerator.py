#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль системы генератора отчетов, основанные на генерации RTF файлов.
"""

# --- Подключение библиотек ---
import os 
import os.path
import copy
import re
from ic.report import rtfReport
from ic.report import icrepgensystem
from ic.report import icreptemplate
from ic.std.log import log

__version__ = (0, 0, 1, 3)

# --- Константы ---
RTF_VAR_PATTERN = r'(#.*?#)'
# Список всех патернов используемых при разборе значений ячеек
ALL_PATERNS = [RTF_VAR_PATTERN]


class icRTFReportGeneratorSystem(icrepgensystem.icReportGeneratorSystem):
    """
    Класс системы генерации отчетов, основанные на генерации RTF файлов.
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
        return self._report_dir

    def _genRTFReport(self, Rep_):
        """
        Генерация отчета и сохранение его в RTF файл.
        @param Rep_: Полное описание шаблона отчета.
        @return: Возвращает имя rtf файла или None в случае ошибки.
        """
        if Rep_ is None:
            Rep_ = self._Rep
        data_rep = self.GenerateReport(Rep_)
        if data_rep:
            rep_file_name = os.path.join(self.getReportDir(), '%s_report_result.rtf' % data_rep['name'])
            template_file_name = os.path.abspath(data_rep['generator'], self.getReportDir())
            log.info(u'Сохранение отчета %s в файл %s' % (template_file_name, rep_file_name))
            
            data = self._predGenerateAllVar(data_rep['__data__'])
            rtfReport.rtfReport(data, rep_file_name, template_file_name)
            return rep_file_name
        return None
        
    def Preview(self, Rep_=None):
        """
        Предварительный просмотр.
        @param Rep_: Полное описание шаблона отчета.
        """
        rtf_rep_file_name = self._genRTFReport(Rep_)
        if rtf_rep_file_name:
            # Открыть в режиме просмотра
            self.PreviewWord(rtf_rep_file_name)
            
    def PreviewWord(self, RTFFileName_):
        """
        Открыть word  в режиме предварительного просмотра.
        @param RTFFileName_: Имя rtf файла, содержащего сгенерированный отчет.
        """
        try:
            # Установить связь с Word
            word_app = win32com.client.Dispatch('Word.Application')
            # Скрыть
            word_app.Visible=0
            # Открыть
            rep_tmpl_book = word_app.Documents.Open(RTFFileName_)
            # Показать
            word_app.Visible = 1
            
            rep_tmpl_book.PrintPreview()
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
        rtf_rep_file_name = self._genRTFReport(Rep_)
        if rtf_rep_file_name:
            # Открыть печать
            self.PrintWord(rtf_rep_file_name)

    def PrintWord(self, RTFFileName_):
        """
        Печать отчета с помощью word.
        @param RTFFileName_: Имя rtf файла, содержащего сгенерированный отчет.
        """
        try:
            # Установить связь с Word
            word_app = win32com.client.Dispatch('Word.Application')
            # Скрыть
            word_app.Visible = 0
            # Открыть
            rep_tmpl_book = word_app.Documents.Open(RTFFileName_)
            # Показать
            word_app.Visible = 1
            
            rep_tmpl_book.PrintOut()
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
        pass

    def OpenWord(self, RTFFileName_):
        """
        Открыть word.
        @param RTFFileName_: Имя rtf файла, содержащего сгенерированный отчет.
        """
        try:
            # Установить связь с Word
            word_app = win32com.client.Dispatch('Word.Application')
            # Скрыть
            word_app.Visible = 0
            # Открыть
            rep_tmpl_book = word_app.Open(RTFFileName_)
            # Показать
            word_app.Visible = 1
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
        # Определить файл *.rtf
        rtf_file = os.path.abspath(os.path.splitext(RepFileName_)[0]+'.rtf')
        cmd = 'start word.exe \'%s\'' % rtf_file
        log.info(u'Выполнение комманды ОС <%s>' % cmd)
        # и запустить MSWord
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
            # Данные отчета.
            # Формат:
            # {
            #    #Переменные
            #    '__variables__':{'имя переменной1':значение переменной1,...}, #Переменные
            #    #Список таблиц
            #    '__tables__':[
            #        {
            #        '__fields__':(('имя поля1'),...), #Поля
            #        '__data__':[(значение поля1,...)], #Данные
            #        },...
            #        ],
            #    #Циклы генерации
            #    '__loop__':{
            #        'имя цикла1':[{переменные и таблицы, используемые в цикле},...],
            #        ...
            #        },
            # }.

            query_data = self.GetQueryTbl(self._Rep)
            if not query_data:
                ic_dlg.icMsgBox(u'Внимание',
                                u'Нет данных, соответствующих запросу: %s' % self._Rep['query'],
                                self._ParentForm)
                return None

            # 2. Запустить генерацию
            rep_data = copy.deepcopy(self._Rep)
            rep_data['__data__'] = query_data
            return rep_data
        except:
            # Вывести сообщение об ошибке в лог
            log.fatal(u'Ошибка генерации отчета %s.' % self._Rep['name'])
            return None

    def _predGenerateAllVar(self, Data_):
        """
        Предобработка всех переменных. Выполняется рекурсивно.
        @return: Возвращает структуры с заполненными переменными.
        """
        if '__variables__' in Data_:
            # Сделать предобработку
            Data_['__variables__'] = self._predGenerateVar(Data_['__variables__'])
        if '__loop__' in Data_:
            for loop_name, loop_body in Data_['__loop__'].items():
                if loop_body:
                    for i_loop in range(len(loop_body)):
                        loop_body[i_loop] = self._predGenerateAllVar(loop_body[i_loop])
                    Data_['__loop__'][loop_name] = loop_body
        return Data_

    def _predGenerateVar(self, DataVar_, Value_=None):
        """
        Предобработка словаря переменных.
        @param DataVar_: Словарь переменных.
        @param Value_: Текущее обработываемое значение.
        """
        if Value_ is None:
            for name, value in DataVar_.items():
                DataVar_[name] = self._predGenerateVar(DataVar_, str(value))
            return DataVar_
        else:
            # Сначала заменить перевод каретки
            value = Value_.replace('\r\n', '\n').strip()
            # Затем распарсить
            parsed = self._funcStrParse(value)
            values = []
            for cur_var in parsed['func']:
                if re.search(RTF_VAR_PATTERN, cur_var):
                    if cur_var[1:-1] in DataVar_:
                        values.append(self._predGenerateVar(DataVar_,
                                      DataVar_[cur_var[1:-1]]))
                    else:
                        values.append('')
                else:
                    log.warning(u'Unknow tag: <%s>' % cur_var)
            # Заполнить формат
            val_str = self._valueFormat(parsed['fmt'], values)
            return val_str

    def _funcStrParse(self, Str_, Patterns_=ALL_PATERNS):
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
            ret = {}
            ret['fmt'] = ''
            ret['func'] = []
    
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
                if not func_find:
                    ret['fmt'] += parsed_str[i_parse]
            return ret
        except:
            # win32api.MessageBox(0, '%s' % (sys.exc_info()[1].args[1]))
            log.fatal()
            return None

    def _valueFormat(self, Fmt_, DataLst_):
        """
        Заполнение формата значения ячейки.
        @param Fmt_: Формат.
        @param DataLst_: Данные, которые нужно поместить в формат.
        @return: Возвращает строку, соответствующую формату.
        """
        # Заполнение формата
        if DataLst_ == []:
            value = Fmt_
        # Обработка значения None
        elif bool(None in DataLst_):
            data_lst = [{None: ''}.setdefault(val, val) for val in DataLst_]
            value = Fmt_ % tuple(data_lst)
        else:
            value = Fmt_ % tuple(DataLst_)
        return value
