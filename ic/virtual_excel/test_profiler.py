#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Профилирование тестов.
После запуска профилирования необходимо сконвертировать логи
hotshot2calltree your_project.prof > your_project.out
Затем можно пранализировать просмотрщиком логов KCachegrind.
Установка: sudo apt-get install kcachegrind kcachegrind-converters
См. инструкцию http://habrahabr.ru/post/110537/.
"""

import os
import os.path
import hotshot

PROF_LOG_FILENAME = './log/virtual_excel_prof.log'
OUT_LOG_FILENAME = './log/virtual_excel_out.log'


def test_1():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/test.ods')
    
    excel.setCellValue('Лист1', 10, 9, 1.99)
    excel.setCellValue('Лист1', 11, 9, '=SUM(I8:I10)')
    
    excel.SaveAs('./testfiles/test_result.ods')
    excel.SaveAs('./testfiles/test_result.xml')
    
    # cmd='soffice ./testfiles/test_result.ods'
    # cmd='soffice ./testfiles/test_result.xml'
    # os.system(cmd)


def do_profiling(testFunction):
    """
    Выполнение профилирования тестовой функции.
    @param testFunction: Объект тестовой функции.
    """
    # Удалить старые логи
    if os.path.exists(PROF_LOG_FILENAME):
        os.remove(PROF_LOG_FILENAME)
        print('INFO. Delete PROF log file <%s>' % PROF_LOG_FILENAME)
     
    if os.path.exists(OUT_LOG_FILENAME):
        os.remove(OUT_LOG_FILENAME)
        print('INFO. Delete OUT log file <%s>' % OUT_LOG_FILENAME)
    
    prof = hotshot.Profile(PROF_LOG_FILENAME)
    prof.start()
  
    # your code goes here
    if testFunction:
        testFunction()
     
    prof.stop()
    prof.close()
    
    # Сконвертировать логи
    if os.path.exists(PROF_LOG_FILENAME):
        cmd = 'hotshot2calltree %s > %s' % (PROF_LOG_FILENAME, OUT_LOG_FILENAME)
        os.system(cmd)
    
        if os.path.exists(OUT_LOG_FILENAME):
            # Запустить просмотрщик логов
            cmd = 'kcachegrind %s' % OUT_LOG_FILENAME
            os.system(cmd)
        else:
            print('ERROR. Not find OUT log file <%s>' % OUT_LOG_FILENAME)
    else:
        print('ERROR. Not find PROF log file <%s>' % PROF_LOG_FILENAME)
            
    
if __name__ == '__main__':
    do_profiling(test_1)
