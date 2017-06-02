#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль функций общего интерфейса к системе генерации.
"""

# Подключение библиотек
from ic.std.log import log
from ic.std.utils import res

from ic.report import icxmlreportgenerator
from ic.report import icodsreportgenerator
from ic.report import icreportmangenerator
from ic.report import icrtfreportgenerator

__version__ = (0, 0, 1, 1)

# Константы подсистемы
REP_GEN_SYS = None

# Спецификации
_ReportGeneratorSystemTypes = {'.xml': icxmlreportgenerator.icXMLReportGeneratorSystem,         # Мой XMLSS генератор
                               '.ods': icodsreportgenerator.icODSReportGeneratorSystem,         # Мой ODS генератор
                               '.rep': icreportmangenerator.icReportManagerGeneratorSystem,     # Report Manager
                               '.rtf': icrtfreportgenerator.icRTFReportGeneratorSystem,         # RTF генератор
                               }


# Функции управления
def getReportGeneratorSystem(RepFileName_, ParentForm_=None, bRefresh=True):
    """
    Получить объект системы генерации отчетов.
    @param RepFileName_: Имя файла шаблона отчета.
    @param ParentForm_: Родительская форма, необходима для вывода сообщений.
    @param bRefresh: Указание обновления данных шаблона отчета в генераторе.
    @return: Функция возвращает объект-наследник класса icReportGeneratorSystem. 
        None - в случае ошибки.
    """
    try:
        # Прочитать шаблон отчета
        rep = res.loadResourceFile(RepFileName_, bRefresh=True)
        
        global REP_GEN_SYS

        # Тип генератора
        if isinstance(rep['generator'], str):
            rep_gen_sys_class = _ReportGeneratorSystemTypes.setdefault(rep['generator'][-4:].lower(), None)
            if rep_gen_sys_class is not None:
                if REP_GEN_SYS is None:
                    REP_GEN_SYS = rep_gen_sys_class(rep, ParentForm_)
                    REP_GEN_SYS.RepTmplFileName = RepFileName_
                elif not isinstance(REP_GEN_SYS, rep_gen_sys_class):
                    REP_GEN_SYS = rep_gen_sys_class(rep, ParentForm_)
                    REP_GEN_SYS.RepTmplFileName = RepFileName_
                else:
                    if bRefresh:
                        # Просто установить обновление
                        REP_GEN_SYS.setRepData(rep)
                        REP_GEN_SYS.RepTmplFileName = RepFileName_
            else:
                log.warning(u'Не известный генератор <%s>' % rep['generator'][-4:])
        else:
            log.warning(u'Не определен генератор <%s>' % rep['generator'])

        # Если родительская форма не определена у системы генерации, то установить ее
        if REP_GEN_SYS and REP_GEN_SYS.getParentForm() is None:
            REP_GEN_SYS.setParentForm(ParentForm_)
            
        return REP_GEN_SYS
    except:
        log.error(u'Ошибка определения объекта системы генерации отчетов. Отчет <%s>.' % RepFileName_)
        raise
    return None


def getCurReportGeneratorSystem(ReportBrowserDlg_=None):
    """
    Возвратить текущую систему генерации.
    """
    global REP_GEN_SYS
    if REP_GEN_SYS is None:
        REP_GEN_SYS = icodsreportgenerator.icODSReportGeneratorSystem(ParentForm_=ReportBrowserDlg_)
    return REP_GEN_SYS
