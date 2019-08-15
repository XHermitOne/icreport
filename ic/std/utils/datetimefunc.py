#!/usr/bin/env python33
# -*- coding: utf-8 -*-

"""
Модуль функций работы с временными даннами и датами.
"""

# --- Подключение библиотек ---
import wx
import time
import datetime
import calendar

from ic.std.log import log

_ = wx.GetTranslation

__version__ = (0, 1, 1, 2)

# --- Константы и переменные ---
DEFAULT_DATETIME_FMT = '%d.%m.%Y'
DEFAULT_DATE_FMT = '%Y.%m.%d'

# Формат хранения даты/времени в БД
DEFAULT_DATETIME_DB_FMT = '%Y.%m.%d %H:%M:%S'

DEFAULT_TIME_FMT = '%H:%M:%S'


# --- Функции работы с датой/временем ---
def getWeekList():
    """
    Список дней недели.
    """
    return [_('Monday'), _('Tuesday'), _('Wednesday'),
            _('Thursday'), _('Friday'),
            _('Saturday'), _('Sunday')]


def isTimeInRange(time_range, time_hour_minute):
    """
    Проверка попадает ли указанное время в указанный временной диапазон.
    @param time_range: Временной диапазон в формате кортежа 
        (нач-час,нач-мин,окон-час,окон-мин).
    @param time_hour_minute: Время в формате кортежа (час,мин).
    @return: Возвращает True, если время в диапазоне.
    """
    if time_range[0] < time_hour_minute[0] < time_range[2]:
        return True
    elif time_hour_minute[0] == time_range[0] or time_hour_minute[0] == time_range[2]:
        if time_range[1] < time_hour_minute[1] < time_range[3]:
            return True
    return False


def DateTime2StdFmt(dt=None):
    """
    Представление времени в стандартном строковом формате.
    @param dt: Время и дата, если None,  то текущие.
    """
    if dt is None:
        dt = time.time()
    return time.strftime('%d.%m.%Y %H:%M:%S', time.localtime(dt))


def getTodayFmt(dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Сегодняшнее число в формате.
    @param dt_fmt: Задание формата.
    @return: Возвращает строку или None в случае ошибки.
    """
    return datetime.date.today().strftime(dt_fmt)


def getToday():
    """
    Сегодняшнее число в формате date.
    @return: Объект date или None в случае ошибки.
    """
    return datetime.date.today()


def getNow():
    """
    Сегодняшнее число/время.
    @return: <datetime>
    """
    return datetime.datetime.now()


def getNowFmt(dt_fmt='%d.%m.%Y %H:%M:%S'):
    """
    Сегодняшнее число/время в формате.
    @param dt_fmt: Задание формата.
    @return: Возвращает строку или None в случае ошибки.
    """
    return time.strftime(dt_fmt, time.localtime(time.time()))


def getMaxDayFmt(dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Максимально возможная дата в формате.
    @param dt_fmt: Задание формата.
    @return: Возвращает строку или None в случае ошибки.
    """
    return datetime.date(datetime.MAXYEAR, 12, 31).strftime(dt_fmt)


def getMinDayFmt(dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Минимально возможная дата в формате.
    @param dt_fmt: Задание формата.
    @return: Возвращает строку или None в случае ошибки.
    """
    return datetime.date(datetime.MINYEAR, 1, 1).strftime(dt_fmt)


def getDateTimeTuple(dt_string='01.01.2005', dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Представление даты_времени в виде кортежа.
    @param dt_string: Число в строковом формате.
    @param dt_fmt: Формат представления строковы данных.
    @return: Представление даты_времени в виде кортежа.
    """
    return time.strptime(dt_string, dt_fmt)


def getMonthDT(dt_string='01.01.2005', dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Месяц в формате datetime.
    @param dt_string: Число в строковом формате.
    @param dt_fmt: Формат представления строковы данных.
    @return: Возвращает укзанный в строке месяц в формате datetime.
    """
    dt_tuple = getDateTimeTuple(dt_string, dt_fmt)
    return datetime.date(dt_tuple[0], dt_tuple[1], 1)


def getOneMonthDelta():
    """
    1 месяц в формате timedelta.
    """
    return datetime.timedelta(31)


def setDayDT(dt, day=1):
    """
    Установить первой дату объекта dete.
    """
    return datetime.date(dt.year, dt.month, day)


def convertDateTimeFmt(dt_string, old_dt_fmt=DEFAULT_DATETIME_FMT, new_dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Преобразовать строковое представления даты-времени в другой формат.
    @param dt_string: Число в строковом формате.
    @param old_dt_fmt: Старый формат представления строковы данных.
    @param new_dt_fmt: Старый формат представления строковы данных.
    @return: Возвращает строку даты-времени в новом формате.
    """
    date_time_tuple = getDateTimeTuple(dt_string, old_dt_fmt)
    return time.strftime(new_dt_fmt, date_time_tuple)


def strDateFmt2DateTime(date_string, dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Преобразование строкового представления даты в указанном формате
    в формат datetime.
    @return: Возвращает объект datetime или None в случае ошибки.
    """
    date_time_tuple = getDateTimeTuple(date_string, dt_fmt)
    year = date_time_tuple[0]
    month = date_time_tuple[1]
    day = date_time_tuple[2]
    return datetime.date(year, month, day)


def strDateTimeFmt2DateTime(dt_string, dt_fmt=DEFAULT_DATETIME_FMT):
    """
    Преобразование строкового представления даты/времени в указанном формате
    в формат datetime.
    @return: Возвращает объект datetime или None в случае ошибки.
    """
    date_time_tuple = getDateTimeTuple(dt_string, dt_fmt)
    year = date_time_tuple[0]
    month = date_time_tuple[1]
    day = date_time_tuple[2]
    hour = date_time_tuple[3]
    minute = date_time_tuple[4]
    second = date_time_tuple[5]
    return datetime.datetime(year, month, day, hour, minute, second)


def getNowYear():
    """
    Текущий год.
    """
    return int(getNowFmt('%Y'))


def getMonthDaysCount(month, year=None):
    """
    Определить сколько дней в месяце по номеру месяца.
    @param month: Номер месяца 1..12.
    @param year: Год. Если None, то текущий год.
    """
    if year is None:
        year = getNowYear()
    else:
        year = int(year)
    month_days = 0
    calendar_list = calendar.Calendar().monthdayscalendar(year, month)
    for week in calendar_list:
        month_days += len([day for day in week if day != 0])
    return month_days


def getWeekDay(day, month, year=None):
    """
    Номер дня недели 1..7.
    @param day: День.
    @param month: Номер месяца 1..12.
    @param year: Год. Если None, то текущий год.
    """
    if year is None:
        year = getNowYear()
    return calendar.weekday(int(year), int(month), int(day)) + 1


def getWeekPeriod(n_week, year=None):
    """
    Возвращает период дат нужной недели.
    @return: Возващает картеж периода дат нужной недели. 
    """
    if not year:
        year = getNowYear()
        
    d1 = datetime.date(year, 1, 1)
    if d1.weekday() > 0:
        delt = datetime.timedelta(7 - d1.weekday())
    else:
        delt = datetime.timedelta(0)
        
    beg = d1 + datetime.timedelta((n_week - 1) * 7) + delt
    end = d1 + datetime.timedelta((n_week - 1) * 7 + 6) + delt
    return beg, end


def genUnicalTimeName():
    """
    Генерация уникальоного имени по текущему времени.
    """
    return getNowFmt('%Y%m%d_%H%M%S')


def pydate2wxdate(date):
    """
    Преобразовать <datetime> тип в <wx.DateTime>.
    @param date: Дата <datetime>.
    @return: Дата <wx.DateTime>.
    """
    if date is None:
        return None

    assert isinstance(date, (datetime.datetime, datetime.date))
    tt = date.timetuple()
    dmy = (tt[2], tt[1]-1, tt[0])
    return wx.DateTimeFromDMY(*dmy)


def wxdate2pydate(date):
    """
    Преобразовать <wx.DateTime> тип в <datetime>.
    @param date: Дата <wx.DateTime>.
    @return: Дата <datetime>.
    """
    if date is None:
        return None

    assert isinstance(date, wx.DateTime)
    if date.IsValid():
        ymd = list(map(int, date.FormatISODate().split('-')))
        return datetime.date(*ymd)
    else:
        return None


def pydatetime2wxdatetime(dt):
    """
    Преобразовать <datetime> тип в <wx.DateTime>.
    @param dt: Дата-время <datetime>.
    @return: Дата-время <wx.DateTime>.
    """
    if dt is None:
        return None

    assert isinstance(dt, (datetime.datetime, datetime.date))
    tt = dt.timetuple()
    dmy = (tt[2], tt[1]-1, tt[0])
    hms = (tt[2], tt[1]-1, tt[0])
    result = wx.DateTimeFromDMY(*dmy)
    result.SetHour(hms[0])
    result.SetMinute(hms[1])
    result.SetSecond(hms[2])
    return result


def wxdatetime2pydatetime(dt):
    """
    Преобразовать <wx.DateTime> тип в <datetime>.
    @param dt: Дата-время <wx.DateTime>.
    @return: Дата-время <datetime>.
    """
    if dt is None:
        return None

    assert isinstance(dt, wx.DateTime)
    if dt.IsValid():
        ymd = [int(t) for t in dt.FormatISODate().split('-')]
        hms = [int(t) for t in dt.FormatISOTime().split(':')]
        dt_args = ymd+hms
        return datetime.datetime(*dt_args)
    else:
        return None


def date2datetime(dt):
    """
    Конвертация datetime.date в datetime.datetime.
    @param dt: Дата в формате datetime.date
    @return: Дата в формате datetime.datetime
    """
    return datetime.datetime.combine(dt, datetime.datetime.min.time())


def test():
    """
    Тестирование функций.
    """
    pass


if __name__ == '__main__':
    test()
