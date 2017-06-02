#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль тестов для виртуального Excel.
"""
import os


def test_oc1():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/OC1/OC1.ods')
    excel.SaveAs('./testfiles/OC1/result_oc1.xml')
    excel.SaveAs('./testfiles/OC1/result_oc1.ods')


def test_oc1_1():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/OC1/oc1.xml')
    excel.SaveAs('./testfiles/OC1/result_oc1_1.ods')


def test_oc1_2():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/OC1/oc1.xml')
    excel.SaveAs('./testfiles/OC1/result_oc1_2.xml')


def test_2():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/test2.ods')
    excel.SaveAs('./testfiles/result_test2.ods')


def test_3():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/test3.ods')
    excel.SaveAs('./testfiles/result_test3.ods')


def test_4():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/bpr1137.ods')
    excel.SaveAs('./testfiles/bpr1137_result.ods')


def test_5():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/narjad.ods')

    excel.setCellValue('Сдельн.', 19, 13, 1.99)

    excel.SaveAs('./testfiles/narjad_result.ods')
    excel.SaveAs('./testfiles/narjad_result.xml')
    
    # cmd='soffice ./testfiles/narjad.ods'
    # os.system(cmd)
    cmd = 'soffice ./testfiles/narjad_result.ods'
    os.system(cmd)


def test_6():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/t1.ods')
    excel.SaveAs('./testfiles/t1_result.ods')
    
    cmd = 'soffice ./testfiles/t1.ods'
    os.system(cmd)
    cmd = 'soffice ./testfiles/t1_result.ods'
    os.system(cmd)


def test_7():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/tabmilk.ods')
    excel.SaveAs('./testfiles/tabmilk_result.ods')
    
    cmd = 'soffice ./testfiles/tabmilk.ods'
    os.system(cmd)
    cmd = 'soffice ./testfiles/tabmilk_result.ods'
    os.system(cmd)


def test_8():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/narjadok.ods')
    excel.SaveAs('./testfiles/narjadok_result.ods')
    
    cmd = 'soffice ./testfiles/narjadok.ods'
    os.system(cmd)
    cmd = 'soffice ./testfiles/narjadok_result.ods'
    os.system(cmd)


def test_9():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/narjadpv.ods')
    excel.SaveAs('./testfiles/narjadpv_result.ods')
    excel.SaveAs('./testfiles/narjadpv_result.xml')
    
    cmd = 'soffice ./testfiles/narjadpv.ods'
    os.system(cmd)
    cmd = 'soffice ./testfiles/narjadpv_result.ods'
    os.system(cmd)


def test_10():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/new/narjadpv.ods')
    excel.SaveAs('./testfiles/new/narjadpv_result.ods')
    excel.SaveAs('./testfiles/new/narjadpv_result.xml')
    
    cmd = 'soffice ./testfiles/new/narjadpv.ods'
    os.system(cmd)
    cmd = 'soffice ./testfiles/new/narjadpv_result.ods'
    os.system(cmd)


def test_11():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/tabbnew.ods')
    # excel.SaveAs('./testfiles/tabbnew_result.ods')
    excel.SaveAs('./testfiles/tabbnew_result.xml')
    
    # cmd='soffice ./testfiles/tabbnew_result.ods'
    # os.system(cmd)


def test_sum():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/test.ods')
    excel.SaveAs('./testfiles/test_result.ods')
    # excel.SaveAs('./testfiles/test_result.xml')
    
    cmd = 'soffice ./testfiles/test_result.ods'
    # cmd='soffice ./testfiles/test_result.xml'
    os.system(cmd)


def test_12():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/test.ods')
    
    excel.setCellValue('Лист1', 10, 9, 1.99)
    excel.setCellValue('Лист1', 11, 9, '=SUM(I8:I10)')
    
    excel.SaveAs('./testfiles/test_result.ods')
    excel.SaveAs('./testfiles/test_result.xml')
    
    cmd = 'soffice ./testfiles/test_result.ods'
    # cmd='soffice ./testfiles/test_result.xml'
    os.system(cmd)


def test_13():
    if os.path.exists('./log/virtual_excel.log'):
        os.remove('./log/virtual_excel.log')
    if os.path.exists('./testfiles/test_result.ods'):
        os.remove('./testfiles/test_result.ods')

    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/test.ods')
    excel.SaveAs('./testfiles/test_result.ods')
    excel.SaveAs('./testfiles/test_result.xml')

    cmd = 'libreoffice ./testfiles/test_result.ods'
    os.system(cmd)


def test_14():
    if os.path.exists('./log/virtual_excel.log'):
        os.remove('./log/virtual_excel.log')

    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/report/rep.xml')
    excel.SaveAs('./testfiles/report/rep.ods')

    cmd = 'libreoffice ./testfiles/report/rep.ods'
    os.system(cmd)


def test_15():
    if os.path.exists('./log/virtual_excel.log'):
        os.remove('./log/virtual_excel.log')

    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/ttn/ttn_original.ods')
    # excel.Load('./testfiles/ttn/ttn.xml')
    excel.SaveAs('./testfiles/ttn/ttn_result.ods')
    excel.SaveAs('./testfiles/ttn/ttn_result.xml')

    cmd = 'libreoffice ./testfiles/ttn/ttn_result.ods'
    os.system(cmd)


def test_16():
    if os.path.exists('./log/virtual_excel.log'):
        os.remove('./log/virtual_excel.log')

    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/imns/prib101.ods')
    excel.SaveAs('./testfiles/imns/prib101_result.ods')

    cmd = 'libreoffice ./testfiles/imns/prib101_result.ods'
    os.system(cmd)


def test_17():
    if os.path.exists('./log/virtual_excel.log'):
        os.remove('./log/virtual_excel.log')

    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/bpr_735u.ods')
    excel.SaveAs('./testfiles/bpr735u_result.ods')

    cmd = 'libreoffice ./testfiles/bpr735u_result.ods'
    os.system(cmd)


def test_18():
    if os.path.exists('./log/virtual_excel.log'):
        os.remove('./log/virtual_excel.log')

    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/break_page.ods')
    excel.SaveAs('./testfiles/break_page_test.ods')

    cmd = 'libreoffice ./testfiles/break_page_test.ods'
    os.system(cmd)


def test_19():
    if os.path.exists('./log/virtual_excel.log'):
        os.remove('./log/virtual_excel.log')

    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/oc1.ods')
    excel.SaveAs('./testfiles/oc1_test.ods')

    cmd = 'libreoffice ./testfiles/oc1_test.ods'
    os.system(cmd)


def test_ods():
    import icexcel
    
    excel = icexcel.icVExcel()
    excel.Load('./testfiles/ods/schm_19.ods')
    excel.SaveAs('./testfiles/ods/result.ods')

    cmd = 'libreoffice ./testfiles/ods/result.ods'
    os.system(cmd)


def test_report():
    import icexcel

    excel = icexcel.icVExcel()
    excel.Load('/home/xhermit/.icreport/rep_tmpl_report_result.xml')
    excel.SaveAs('/home/xhermit/.icreport/rep_tmpl_report_result.ods')

    cmd = 'libreoffice /home/xhermit/.icreport/rep_tmpl_report_result.ods'
    os.system(cmd)


def test_worksheet_options():
    import icexcel

    # if os.path.exists('./log/virtual_excel.log'):
    #     os.remove('./log/virtual_excel.log')

    excel = icexcel.icVExcel()
    excel.Load('./testfiles/month_params_report.ods')
    excel.SaveAs('./testfiles/ods/result.ods')

    cmd = 'libreoffice ./testfiles/ods/result.ods'
    os.system(cmd)


if __name__ == '__main__':
    # test_oc1()
    # test_sum()
    # test_12()
    # test_5()
    # test_17()
    # test_ods()
    # test_worksheet_options()
    test_report()
