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
from ic.report import icxlsreportgenerator
from ic.report import icreportmangenerator
from ic.report import icrtfreportgenerator

__version__ = (0, 0, 2, 1)

# Константы подсистемы
REP_GEN_SYS = None

# Спецификации
_ReportGeneratorSystemTypes = {'.xml': icxmlreportgenerator.icXMLReportGeneratorSystem,         # Мой XMLSS генератор
                               '.ods': icodsreportgenerator.icODSReportGeneratorSystem,         # Мой ODS генератор
                               '.xls': icxlsreportgenerator.icXLSReportGeneratorSystem,         # Мой XLS генератор
                               '.rep': icreportmangenerator.icReportManagerGeneratorSystem,     # Report Manager
                               '.rtf': icrtfreportgenerator.icRTFReportGeneratorSystem,         # RTF генератор
                               }

# Список расширений источников шаблонов
SRC_REPORT_EXT = _ReportGeneratorSystemTypes.keys()


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

        # Создание системы ренерации отчетов
        if REP_GEN_SYS is None:
            REP_GEN_SYS = createReportGeneratorSystem(rep['generator'], rep, ParentForm_)
            REP_GEN_SYS.RepTmplFileName = RepFileName_
        elif not REP_GEN_SYS.sameGeneratorType(rep['generator']):
            REP_GEN_SYS = createReportGeneratorSystem(rep['generator'], rep, ParentForm_)
            REP_GEN_SYS.RepTmplFileName = RepFileName_
        else:
            if bRefresh:
                # Просто установить обновление
                REP_GEN_SYS.setRepData(rep)
                REP_GEN_SYS.RepTmplFileName = RepFileName_

        # Если родительская форма не определена у системы генерации,
        # то установить ее
        if REP_GEN_SYS and REP_GEN_SYS.getParentForm() is None:
            REP_GEN_SYS.setParentForm(ParentForm_)
            
        return REP_GEN_SYS
    except:
        log.error(u'Ошибка определения объекта системы генерации отчетов. Отчет <%s>.' % RepFileName_)
        raise
    return None


def createReportGeneratorSystem(sRepGenSysType, dRep=None, ParentForm_=None):
    """
    Создать объект системы генерации отчетов.
    @param sRepGenSysType: Указание типа системы генерации отчетов.
        Тип задается расширением файла источника шаблона.
        В нашем случае один из SRC_REPORT_EXT.
    @param dRep: Словарь отчета.
    @param ParentForm_: Родительская форма, необходима для вывода сообщений.
    @return: Функция возвращает объект-наследник класса icReportGeneratorSystem.
        None - в случае ошибки.
    """
    rep_gen_sys_type = sRepGenSysType[-4:].lower() if type(sRepGenSysType) in (str, unicode) else None
    rep_gen_sys = None
    if rep_gen_sys_type:
        rep_gen_sys_class = _ReportGeneratorSystemTypes.setdefault(rep_gen_sys_type, None)
        if rep_gen_sys_class is not None:
            rep_gen_sys = rep_gen_sys_class(dRep, ParentForm_)
        else:
            log.warning(u'Не известный тип генератора <%s>' % rep_gen_sys_type)
    else:
        log.warning(u'Не корректный тип генератора <%s>' % sRepGenSysType)
    return rep_gen_sys


def getCurReportGeneratorSystem(ReportBrowserDlg_=None):
    """
    Возвратить текущую систему генерации.
    """
    global REP_GEN_SYS
    if REP_GEN_SYS is None:
        REP_GEN_SYS = icodsreportgenerator.icODSReportGeneratorSystem(ParentForm_=ReportBrowserDlg_)
    return REP_GEN_SYS
