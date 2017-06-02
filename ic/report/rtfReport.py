#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Заполнение rtf шаблона.
"""

import copy

__version__ = (0, 0, 1, 2)

tblDct = {'__fields__': (('n_lot',), ('predmet_lot',), ('cena_lot',)),
          '__name__': 'P1',
          '__data__': [(1, 'лот 1.1', 200.0),
                       (2, 'лот 1.2', 210.0),
                       (3, 'лот 1.3', 220.0),
                       (4, 'лот 1.4', 300.0)]}

tblDct2 = {'__fields__': (('n_lot',), ('predmet_lot',), ('cena_lot',)),
          '__name__': 'P2',
          '__data__': [(1, 'лот 2.1', 500.0),
                       (2, 'лот 2.2', 510.0),
                       (3, 'лот 2.3', 520.0)]}

tblDct3 = {'__fields__': (('n_lot',), ('predmet_lot',), ('cena_lot',)),
          '__name__': 'P1',
          '__data__': [(1, 'лот 3.1', 200.0),
                       (2, 'лот 3.2', 210.0),
                       (3, 'лот 3.3', 220.0),
                       (4, 'лот 3.4', 300.0)]}

tblDct4 = {'__fields__': (('n_lot',), ('predmet_lot',), ('cena_lot',)),
          '__name__': 'P2',
          '__data__': [(1, 'лот 4.1', 500.0),
                       (2, 'лот 4.2', 510.0),
                       (3, 'лот 4.3', 520.0)]}

LDct1 = {'__variables__': {'VAR_L': 'Слот 1'},
         '__tables__': [tblDct, tblDct2]}
LDct2 = {'__variables__': {'VAR_L': 'Слот 2'},
         '__tables__': [tblDct3, tblDct4]}
LDct3 = {'__variables__': {'VAR_L': 'Слот 3'}}


def SD(var, val):
    return {'__variables__': {var: val}}
    
DataDct = {'__variables__': {'form_torg': """Тип торгов
hjkfdshjkhjksfadhlkhfsa
sfdjhjkhfajkhfjdskhfhks
    GGGGGGGGGGGGGGGGGg
""",
                             'izveschen_url': 'www.abakan.ru',
                             'VAR_L_1': 'Слот 1',
                             'VAR_L_2': 'Слот 2',
                             'VAR_LG': 'Общий текст',
                             'name_torg': 'Имя торгов'},
            '__loop__': {'L': [LDct1, LDct2],
                         'I': [SD('I', 1), SD('I', 2), SD('I', 3)],
                         'J': [SD('J', 1), SD('J', 2), SD('J', 3)]},
            '__tables__': []}


def findNextVar(rep, pos=0):
    """
    Ищет следующую переменную.
    
    @type rep: C{string}
    @param rep: Шаблон.
    @type pos: C{int}
    @param pos: Позиция, с которой искать.
    @rtype: C{tuple}
    @return: Возвращает картеж. 1-й элемент начальная позиция текста замены;
        2-й элемент конечная позиция; 3-й элемент - имя переменной.
    """
    p1 = rep.find('#', pos)

    if p1 == -1:
        return -1, -1, None
    
    p2 = rep.find('#', p1+1)

    if p2 == -1:
        return -1, -1, None
    
    # --- Определяем имя переменной
    var = rep[p1+1:p2]
    
    # --- Выкидываем все лишнее
    
    # Ищем конец группы
    n2 = var.find('}')
        
    if n2 > -1:
        s = var[: n2]
    else:
        s = ''
        
    n1 = 0
    
    while 1:
        n1 = var.find(' ', n1)
        
        if n1 == -1:
            s = var
            break
        
        # Ищем конец группы
        n2 = var.find('}', n1+1)
        
        if n2 == -1:
            s += var[n1+1: p2]
            break
        
        s += var[n1+1: n2]
        n1 = n2+1
        
    s = s.replace('\r', '').replace('\n', '')

    if ' ' in s:
        p1, p2, s = findNextVar(rep, p1+1)

    return p1, p2, s


def getLoopLst(data, lname):
    """
    Определяет список элементов в цикле.
    """
    lst = []
    if '__loop__' in data and lname in data['__loop__']:
        ls = data['__loop__'][lname]
        
        #   Заимствуем описание родительских переменных
        for el in ls:
        
            # Заимствуем описание циклов
            if '__loop__' in el:
                elm={'__loop__': copy.copy(el['__loop__'])}
            else:
                elm = {'__loop__': {}}
                if '__loop__' in data:
                    for lp, sp in data['__loop__'].items():
                        if lp != lname:
                            elm['__loop__'][lp] = copy.copy(sp)
            
            # Заимствуем описание переменных
            elm['__variables__'] = copy.copy(data['__variables__'])
            elm['__variables__'].update(el['__variables__'])
            
            # Заимствуем описание таблиц
            elm['__tables__'] = copy.copy(data['__tables__'])
            nmLst = [x['__name__'] for x in elm['__tables__']]

            if '__tables__' in el:
                for tbl in el['__tables__']:
                    if '__name__' in tbl and not tbl['__name__'] in nmLst:
                        elm['__tables__'].append(tbl)
            
            lst.append(elm)
            
        return lst
    
    return lst


def getTableTempl(rep, pos):
    """
    """
    # --- Разбираем шаблон таблицы
    p2 = pos
    
    while 1:
        p1, p2, var = findNextVar(rep, p2+1)
    
        if not var:
            break

        #   Ищем коней кгруппы
        if var == '1D':
            break
        else:
            pass
            
    return pos, p2


repTest = '''
    #Variable1#
    #Variblee2#
'''


def _gen_table(table, templ):
    """
    Генерирует таблицу по шаблону и табличным данным.
    """
    txt = ''
    for r in table['__data__']:
        replDct = {}
        for indx, col in enumerate(table['__fields__']):
            replDct[col[0]] = str(r[indx])
            
        txt += parse_rtf(None, templ, replDct)

    return txt


def DoTabelByName(data, templ, name):
    """
    Генерация по имени таблицы.
    """
    table = None
    
    for tbl in data['__tables__']:
        if tbl['__name__'] == name:
            return _gen_table(tbl, templ)

    return ''


def DoTabelByIndx(data, templ, indx):
    """
    Генерация по индексу таблицы.
    """
    if len(data['__tables__']) > indx:
        tbl = data['__tables__'][indx]
        return _gen_table(tbl, templ)

    return ''


def parse_rtf(data, rep, replDct=None, indxLoop=None, indxKey=None):
    """
    """
    p2 = 0
        
    if not replDct:
        replDct = data['__variables__']

    if data and '__tables__' not in data:
        data['__tables__'] = []
        
    bTableBeg = False
    tBegTag = 0
    tEndTag = 0
    tName = None
    
    bLoopBeg = False
    lBegTag = 0
    lEndTag = 0
    lName = None

    indx = 0
    
    if indxLoop and indxKey:
        varKey = '%s_%s' % (indxLoop, indxKey)
    else:
        varKey = ''
    
    # --- Разбираем шаблон
    while 1:
        p1, p2, var = findNextVar(rep, p2+1)
    
        if not var:
            break
            
        if var[:2] == 'D1' and not bLoopBeg:
            bTableBeg = True
            tBegTag = p1
            tEndTag = p2
            
            if varKey:
                tName = var[3:]+'_'+indxKey
            else:
                tName = var[3:]

        elif var[:4] == 'LOOP' and not bLoopBeg:
            bLoopBeg = True
            lBegTag = p1
            lEndTag = p2
            lName = var[5:]

        elif var[:7] == 'ENDLOOP':
            loop_name = var[8:]
            
            if loop_name == lName:
                # Выделяем повторяющуюся часть
                templ = rep[lEndTag+2:p1]
                txt = ''
                
                n = templ.find('}')
                # Убираем первую группу из шаблона
                if n != -1:
                    templ = templ[n+1:]
                
                n = templ.rfind('{')
                # Убираем последнюю группу из шаблона
                if n != -1:
                    templ = templ[:n]

                lst = getLoopLst(data, lName)
                txt = ''
                if lst:
                    for sp in lst:
                        txt += parse_rtf(sp, templ)
                else:
                    # Если список не определен просто запустить генерацию
                    # чтобы стереть все теги из текста
                    txt += parse_rtf({'__variables__': {}}, templ)
                    
                # Вставляем текст
                if txt:
                    rep = rep[:lBegTag] + txt + rep[p2+1:]
                    p2 = lBegTag + 1
                    
                bLoopBeg = False
        
        elif var == '1D' and not bLoopBeg:
            # Выделяем табличную часть
            templ = rep[tEndTag+2:p1]
            txt = ''
            
            n = templ.find('}')
            # Убираем первую группу из шаблона
            if n != -1:
                templ = templ[n+1:]
            
            n = templ.rfind('{')
            # Убираем последнюю группу из шаблона
            if n != -1:
                templ = templ[:n]
            
            # Генерируем таблицу
            if tName:
                txt = DoTabelByName(data, templ, tName)
            else:
                txt = DoTabelByIndx(data, templ, indx)
                indx += 1

            # Вставляем таблицу
            rep = rep[:tBegTag] + txt + rep[p2+1:]
            p2 = tBegTag + 1
            bTableBeg = False

        #   Вставляем переменные
        elif not bTableBeg and not bLoopBeg:
            
            if varKey:
                v = var+'_'+str(indxKey)
                if v not in replDct.keys():
                    v = var
            else:
                v = var
                
            if v in replDct.keys():
                replTxt = str(replDct[v]).replace('\n', '\line ')
                rep = rep[:p1] + replTxt + rep[p2+1:]
                p2 = p1 + len(str(replDct[v]))+1
            else:
                rep = rep[:p1] + rep[p2+1:]
                p2 = p1 + 1
    
    return rep


def rtfReport(data, repFileName, templFileName):
    """
    Создает rtf отчет по шаблону.
    """
    f = open(templFileName, 'r')
    rep = f.read()
    f.close()

    rep = parse_rtf(data, rep)
    
    f = open(repFileName, 'w')
    f.write(rep)
    f.close()


def test():
    rtfReport(DataDct, 'V:/pythonprj/ReportRTF/IzvTEST.rtf', 'V:/pythonprj/ReportRTF/Blank_Izv.rtf')


if __name__ == '__main__':
    test()
