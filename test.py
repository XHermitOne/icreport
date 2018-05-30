#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Фукнции тестирования.
"""

import os

CMD = ('python icreport.py --help',
       # 'python icreport.py --export="./tst/Схема 19 позиций.xls" ',
       # 'rm ./tst/schema_19_pos.rprt; python icreport.py --path=./tst --export=schema_19_pos.xls  --var="n1=1000"',
       'python icreport.py --path=./tst --export=schema_19_pos.xls --var="n1=1000"',
       )


def test_all():
    """
    Протестировать все.
    """
    for cmd in CMD:
        os.system(cmd)


if __name__ == '__main__':
    test_all()
