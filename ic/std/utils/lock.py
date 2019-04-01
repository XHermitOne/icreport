#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Библиотека блокировок.
Формат:
Информация о координатах блокировок (таблица/запись) храняться:
таблица: имя директории
запись: имя файла в этой директории (если файл существует - запись заблокирована)

Если заблокирована таблица - добавляется к имени расширение '.lck'
"""

# --- Подключение пакетов ---
import os
import os.path
import stat

from . import ic_util
from . import textfunc
from ..log import log

__version__ = (0, 1, 1, 1)

# --- Константы ---
# Расширение файла блокировки
LOCK_FILE_EXT = '.lck'

# Имя блокировки по умолчанию
DEFAULT_LOCK_NAME = 'default'

# Это путь к общей директории блокировок
LOCK_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'lock')

# код текущей ошибки
ERROR_CODE = 0

# key - код оишбки, сообщение (строка)
ERROR_CODE2MESSAGE = {1: u'Таблица заблокирована. Не возможно заблокировать запись.',
                      2: u'Запись заблокирована.',      # №2
                      3: u'',                           # №3
                      4: u'',                           # №4
                      5: u'',                           # №5
                      99: u'Не известная ошибка.'       # 99
                      }


# --- Функции ---
def LockRecord(table, record, message=None):
    """
    Блокировка записи по имени/номеру таблицы и номеру записи
    @param table: -имя таблицы (int/String)
    @param record:  -номер записи (int/String)
    @param message: -тестовое сообщение (небязательное)
    """
    global ERROR_CODE
    ERROR_CODE = 0
    log.info(u'### Блокировка записи. Путь блокировки <%s>' % os.getcwd())
    table = __norm(table)
    record = __norm(record)

    if isLockTable(record):
        # Вся таблица уже заблокирована
        # Запись заблокировать невозможно
        ERROR_CODE = 1
        return ERROR_CODE

    # это путь к директории флагов блокировок этой таблицы
    table_lock_dirname = os.path.join(getLockDir(), table)
    if not os.path.isdir(table_lock_dirname):
        # директории блокировок
        # под эту таблицу еще не создали
        try:
            os.makedirs(table_lock_dirname)       # Создать директори. под эту таблицу
        except:
            log.error(u'Ошибка создания папки <%s>' % table_lock_dirname)
            ERROR_CODE = 99
    # Проверка на блокирвку всего файла недоделана!!!!!!!!!!!!!!!!
    # Генерация файла-флага блокировки
    record_lock_filename = os.path.join(table_lock_dirname, record)
    try:
        # Попытка создать файл
        f = os.open(record_lock_filename, os.O_CREAT | os.O_EXCL,
                    mode=stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
    except OSError:
        # Уже есть файл. Т.Е. уже заблокирован
        ERROR_CODE = 2
    else:
        # выполнено без ошибки
        if isinstance(message, str):
            os.close(f)
            try:
                f = os.open(record_lock_filename, os.O_WRONLY,
                            mode=stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
                os.write(f, message.encode())
            except:
                log.fatal(u'Ошибка записи файла блокировки <%s>' % record_lock_filename)
        os.close(f)
                
    return ERROR_CODE


def unLockRecord(table, record):
    """
    Разблокировка записи по имени/номеру таблицы и номеру записи.
    @type table: C{int/string}
    @param table: Имя таблицы.
    @type record: C{int/string}
    @parma record: Номер записи.
    """
    global ERROR_CODE
    ERROR_CODE = 0
    table = __norm(table)
    record = __norm(record)
    # это путь к директории флагов блокировок этой таблицы
    table_lock_dirname = os.path.join(getLockDir(), table)
    if os.path.isdir(table_lock_dirname):
        # директории блокировок под эту таблицу еще не создали
        record_lock_filename = os.path.join(table, record)
        if os.path.isfile(record_lock_filename):
            try:
                # удалить этот файл флага
                os.remove(record_lock_filename)
            except:
                log.error(u'Ошибка удаления файла <%s>' % record_lock_filename)
                ERROR_CODE = 99  # оишбка удаления файла флага
        return ERROR_CODE


def LockTable():
    pass


def UnLockTable():
    pass


def isLockTable(table):
    """
    Проверка на блокировку таблицы.
    """
    table = __norm(table)
    # log.info('%s, %s' % (table, getLockDir()))
    # это путь к директории флагов блокировок этой таблицы
    path = os.path.join(getLockDir(), table+'.lck')
    return os.path.isdir(path)


def readMessage(table, record):
    """
    Чтение текста собщения, если оно есть.
    @type table: C{int/string}
    @param table: Имя таблицы.
    @type record: C{int/string}
    @parma record: Номер записи.
    """
    ret = None
    f = None
    if isLockRecord(table, record) != 0:
        table = __norm(table)
        record = __norm(record)
        # это путь к директории флагов блокировок этой таблицы
        table_lock_dirname = os.path.join(getLockDir(), table)
        record_lock_filename = os.path.join(table_lock_dirname, record)
        try:
            # Попытка создать файл
            f = os.open(record_lock_filename, os.O_RDONLY | os.O_EXCL,
                        mode=stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
        except OSError:
            # Уже есть файл. Т.Е. уже заблокирован
            ERROR_CODE = 2

    if f:
        # выполнено без ошибки
        ret = os.read(f, 65535)
        os.close(f)
    return ret


def isLockRecord(table, record):
    """
    Проверка на блокировку записи.
    @type table: C{int/string}
    @param table: Имя таблицы.
    @type record: C{int/string}
    @parma record: Номер записи.
    """
    ret = None
    global ERROR_CODE
    table = __norm(table)
    record = __norm(record)
    ret = 0
    # это путь к директории флагов блокировок этой таблицы
    table_lock_dirname = os.path.join(getLockDir(), table)
    if os.path.isdir(table_lock_dirname):
        record_lock_filename = os.path.join(table_lock_dirname, record)
        if os.path.isfile(record_lock_filename):
            ret = 1
        else:
            ret = 0
    return ret


def lastErr():
    """
    Вернуть последнне значение ошибки (число).
    """
    global ERROR_CODE
    if ERROR_CODE == 0:
        return 0
    else:
        return ERROR_CODE


def lastErrMsg():
    """
    Вернуть последнне значение ошибки (Строковое сообщение).
    """
    global ERROR_CODE
    if ERROR_CODE == 0:
        return u''
    else:
        return ERROR_CODE2MESSAGE[ERROR_CODE]


def __norm(var):
    """
    Служебная функция 'приведения' к виду имения файла имени таблицы
    и номера записи.
    """
    if isinstance(var, str):
        pass
    elif isinstance(var, int):
        var = str(var)          # нормализация имени номера записи
    elif isinstance(var, float):
        var = str(int(var))     # нормализация имени номера записи
    return var


def getLockDir():
    """
    Определить папку блокировок.
    """
    lock_dir = None
    if not lock_dir:
        log.warning(u'Не определена папка блокировок. Используется папка по умолчанию <%s>' % LOCK_DIR)
        return LOCK_DIR
    return lock_dir


def DelMyLockInDir(LockMyID_, LockDir_, DirFilesLock_):
    """
    Удалить блокировки только из указанной папки.
    @param LockMyID_: Идентификация хозяина блокировок.
    @param LockDir_: Папка блокировок.
    @param DirFilesLock_: Имена файлов и папок в директории LockDir_
    """
    try:
        # Отфильтровать только файлы
        lock_files = [x for x in [os.path.join(LockDir_, x) for x in DirFilesLock_] if os.path.isfile(x)]
        # Выбрать только свои файлы-блокировки
        for cur_file in lock_files:
            f = None
            try:
                f = open(cur_file, 'rt')
                signature = f.read()
                f.close()
                try:
                    signature = eval(signature)
                    # Если владелец в сигнатуре совпадает, то
                    # удалить этот файл-блокировку
                    if signature['computer'] == LockMyID_:
                        os.remove(cur_file)
                        log.info(u'Файл блокировки <%s> удален' % cur_file)
                except:
                    log.warning(u'Не корректная сигнатура блокировки <%s>' % signature)
            except:
                if f:
                    f.close()
                log.fatal(u'Ошибка чтения сигнатуры файла блокировки <%s>' % cur_file)
    except:
        log.fatal(u'Ошибка удаления файлов блокировки. Папка блокировки <%s>' % LockDir_)


def DelMyLock(LockMyID_=None, LockDir_=LOCK_DIR):
    """
    Функция рекурсивного удаления блокировок записей.
    @param LockMyID_: Идентификация хозяина блокировок.
    @param LockDir_: Папка блокировок.
    """
    if not LockMyID_:
        LockMyID_ = GetMyHostName()
    return os.path.walk(LockDir_, DelMyLockInDir, LockMyID_)


# --- Блокировки ресурсов ---
def LockFile(FileName_, LockRecord_=None):
    """
    Блокировка файла.
    @param FileName_: Полное имя блокируемого файла.
    @param LockRecord_: Запись блокировки.
    @return: Возвращает кортеж:
        (результат выполения операции, запись блокировки).
    """
    lock_file_flag = False  # Флаг блокировки файла
    lock_rec = LockRecord_
    str_lock_rec = u''

    # Сгенерировать имя файла блокировки
    lock_file = os.path.splitext(FileName_)[0]+LOCK_FILE_EXT
    # Если файл не заблокирован, то заблокировать его
    if not os.path.isfile(lock_file):
        # Создать все директории для файла блокировки
        lock_dir = os.path.dirname(lock_file)
        if not os.path.isdir(lock_dir):
            try:
                os.makedirs(lock_dir)
            except:
                log.error(u'Ошибка создания папки <%s>' % lock_dir)
        
        # Генерация файла-флага блокировки
        # ВНИМАНИЕ! Создавать файл надо на самом нижнем уровне!
        f = None
        try:
            #  Попытка создать файл
            f = os.open(lock_file, os.O_CREAT | os.O_EXCL,
                        mode=stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
        except OSError:
            #  Уже есть файл. Т.Е. уже заблокирован
            lock_file_flag = True
            # Прочитать кем хоть заблокирован
            if f:
                os.close(f)     # Закрыть сначала
            lock_rec = ReadLockRecord(lock_file)
        else:
            # выполнено без ошибки
            # Записать запись блокировки в файл
            if LockRecord_ is not None:
                os.close(f)     # Закрыть сначала
                # Открыть для записи
                f = os.open(lock_file, os.O_WRONLY,
                            mode=stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
                if isinstance(LockRecord_, str):
                    str_lock_rec = LockRecord_
                else:
                    str_lock_rec = str(LockRecord_)
                os.write(f, str_lock_rec.encode())
            os.close(f)
    else:
        # Если файл заблокирован
        lock_file_flag = True
        lock_rec = ReadLockRecord(lock_file)

    return not lock_file_flag, lock_rec


def ReadLockRecord(LockFile_):
    """
    Прочитать запись блокировки из файла блокировки.
    @param LockFile_: Имя файла блокировки.
    @return: Возвращает запись блокировки или None в случае ошибки.
    """
    f = None
    lock_file = None
    try:
        lock_rec = None
        # На всякий случай преобразовать
        lock_file = os.path.splitext(LockFile_)[0]+LOCK_FILE_EXT
        # Если файла не существует, тогда и нечего прочитать
        if not os.path.exists(lock_file):
            return None
        # Открыть для чтения
        f = os.open(lock_file, os.O_RDONLY,
                    mode=stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
        lock_rec = os.read(f, 65535)
        os.close(f)
        try:
            # Если храниться какая-либо структура,
            # то сразу преобразовать ее
            return eval(lock_rec)
        except:
            return lock_rec
    except:
        if f:
            os.close(f)
        log.fatal(u'Ошибка чтения записи файла блокировки <%s>' % lock_file)
        return None


def IsLockedFile(FileName_):
    """
    Проверка блокировки файла.
    @param FileName_: Имя файла.
    @return: Возвращает результат True/False.
    """
    # Сгенерировать имя файла блокировки
    lock_file = os.path.splitext(FileName_)[0]+LOCK_FILE_EXT
    return os.path.isfile(lock_file)


def ComputerName():
    """
    Имя хоста.
    @return: Получит имя компа в сети.
    """
    comp_name = None
    if 'COMPUTERNAME' in os.environ:
        comp_name = os.environ['COMPUTERNAME']
    else:
        import socket
        comp_name = socket.gethostname()
        
    # ВНИМАНИЕ! Имена компьютеров должны задаваться только латиницей
    # Под Win32 можно задать имя компа русскими буквами и тогда
    # приходится заменять все на латиницу.
    if isinstance(comp_name, str):
        if ic_util.isOSWindowsPlatform():
            comp_name = textfunc.rus2lat(comp_name)
    return comp_name


def GetMyHostName():
    """
    Получит имя компа в сети.
    """
    return ComputerName()


def UnLockFile(FileName_, **If_):
    """
    Разблокировать файл.
    @param FileName_: Имя файла.
    @param If_: Условие проверки разблокировки.
        Ключ записи блокировки=значение.
        Проверка производится по 'И'.
        Если такого ключа в записи нет,
        то его значение берется None.
    @return: Возвращает результат True/False.
    """
    # Сгенерировать имя файла блокировки
    lock_file = os.path.splitext(FileName_)[0]+LOCK_FILE_EXT
    log.info(u'Сброс блокировки. Файл <%s> (%s). Условия снятия блокировки %s' % (lock_file,
                                                                                  os.path.exists(lock_file), If_))
    if os.path.exists(lock_file):
        if If_:
            lck_rec = ReadLockRecord(lock_file)
            # Если значения по указанным ключам равны, то все ОК
            can_unlock = bool(len([key for key in If_.keys() if lck_rec.setdefault(key, None) == If_[key]]) == len(If_))
            log.info(u'Запись блокировки %s. Право снятия блокировки <%s>' % (lck_rec, can_unlock))
            if can_unlock:
                # Ресурс можно разблокировать
                os.remove(lock_file)
            else:
                # Нельзя разблокировать файл
                return False
        else:
            # Ресурс можно разблокировать
            os.remove(lock_file)
    log.info(u'Блокировка снята <%s>' % lock_file)
    return True


def _UnLockFileWalk(args, CurDir_, CurNames_):
    """
    Вспомогательная функция разблокировки файла на уровне каталога по имени
    компьютера. Используется в функции os.path.walk().
    @param args: Кортеж (Имя компьютера файлы которого нужно раблокировать,
        Имя пользователя).
    @param CurDir_: Текущий директорий.
    @param CurNames_: Имена поддиректорий и файлов в текущей директории.
    """
    computer_name = args[0]
    user_name = args[1]
    # Отфильтровать только файлы блокировок
    lock_files = [x for x in [os.path.join(CurDir_, x) for x in CurNames_] if os.path.isfile(x) and os.path.splitext(x)[1] == LOCK_FILE_EXT]
    # Выбрать только свои файлы-блокировки
    for cur_file in lock_files:
        lock_record = ReadLockRecord(cur_file)
        if not user_name:
            if lock_record['computer'] == computer_name:
                os.remove(cur_file)
                log.info(u'Блокировка снята <%s>' % cur_file)
        else:
            if lock_record['computer'] == computer_name and \
               lock_record['user'] == user_name:
                os.remove(cur_file)
                log.info(u'Блокировка снята <%s>' % cur_file)


def UnLockAllFile(LockDir_, ComputerName_=None, UserName_=None):
    """
    Разблокировка всех файлов.
    @param LockDir_: Директория блокировок.
    @param ComputerName_: Имя компьютера файлы которого нужно раблокировать.
    @return: Возвращает результат True/False.
    """
    if not ComputerName_:
        ComputerName_ = ComputerName()
    if not UserName_:
        import ic.engine.ic_user
        UserName_ = ic.engine.ic_user.icGet('UserName')
    if LockDir_:
        return os.walk(LockDir_, _UnLockFileWalk, (ComputerName_, UserName_))


# --- Система блокировки произвольных ресурсов ---
class icLockSystem:
    """
    Система блокировки произвольных ресурсов.
    """

    def __init__(self, LockDir_=None):
        """
        Конструктор.
        @param LockDir_: Папка блокировки.
        """
        if LockDir_ is None:
            LockDir_ = LOCK_DIR
        
        self._LockDir = LockDir_
        
    # --- Папочные блокировки ---
    def lockDirRes(self, LockName_):
        """
        Поставить блокировку в виде директории.
        @param LockName_: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к директории.
        """
        pass
        
    def unLockDirRes(self, LockName_):
        """
        Убрать блокировку в виде директории.
        @param LockName_: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к директории.
        """
        pass

    # --- Файловые блокировки ---
    def _getLockFileName(self, LockName_):
        """
        Определитьимя файла блокировки по имени блокировки.
        @param LockName_: Имя блокировки.
        """
        lock_name = DEFAULT_LOCK_NAME
        lock_file_name = ''
        try:
            if isinstance(LockName_, list):
                lock_name = LockName_[-1]
            elif isinstance(LockName_, str):
                lock_name = os.path.splitext(os.path.basename(LockName_))[0]
            lock_file_name = os.path.join(self._LockDir, lock_name+LOCK_FILE_EXT)
            return lock_file_name
        except:
            log.error(u'Папка блокировки <%s>. Имя блокировки <%s>' % (self._LockDir, LockName_))
            log.fatal(u'Ошибка определения имени файла блокировки')
        return lock_file_name
        
    def lockFileRes(self, LockName_, LockRec_=None):
        """
        Поставить блокировку в виде файла.
        @param LockName_: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к файлу.
        @param LockRec_: Запись блокировки.
        """
        lock_file_name = self._getLockFileName(LockName_)
        if LockRec_ is None:
            import ic.engine.ic_user
            LockRec_ = {'computer': ComputerName(),
                        'user': ic.engine.ic_user.icGet('UserName')}
        return LockFile(lock_file_name, LockRec_)
        
    def unLockFileRes(self, LockName_):
        """
        Убрать блокировку в виде файла.
        @param LockName_: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к файлу.
        """
        lock_file_name = self._getLockFileName(LockName_)
        return UnLockFile(lock_file_name)

    def isLockFileRes(self, LockName_):
        """
        Существует ли файловая блокировка с именем.
        @param LockName_: Имя блокировки.
        """
        lock_file_name = self._getLockFileName(LockName_)
        return IsLockedFile(lock_file_name)
    
    def getLockRec(self, LockName_):
        """
        Определить запись блокировки.
        @param LockName_: Имя блокировки.
        """
        lock_file_name = self._getLockFileName(LockName_)
        return ReadLockRecord(lock_file_name)
    
    # --- Общие функции блокировки ---
    def isLockRes(self, LockName_):
        """
        Существует ли блокировка с именем.
        @param LockName_: Имя блокировки.
        """
        pass
        
    def unLockAllMy(self):
        """
        Разблокировать все мои блокировки.
        """
        return UnLockAllFile(self._LockDir, ComputerName())
