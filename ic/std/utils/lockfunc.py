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

from . import utilfunc
from . import textfunc
from ..log import log

__version__ = (0, 1, 1, 2)

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
def lockRecord(table, record, message=None):
    """
    Блокировка записи по имени/номеру таблицы и номеру записи
    :param table: -имя таблицы (int/String)
    :param record:  -номер записи (int/String)
    :param message: -тестовое сообщение (небязательное)
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
    :type table: C{int/string}
    :param table: Имя таблицы.
    :type record: C{int/string}
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


def lockTable():
    pass


def unLockTable():
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
    :type table: C{int/string}
    :param table: Имя таблицы.
    :type record: C{int/string}
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
    :type table: C{int/string}
    :param table: Имя таблицы.
    :type record: C{int/string}
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


def delMyLockInDir(lock_id, lock_dirname, dir_lock_filenames):
    """
    Удалить блокировки только из указанной папки.
    :param lock_id: Идентификация хозяина блокировок.
    :param lock_dirname: Папка блокировок.
    :param dir_lock_filenames: Имена файлов и папок в директории lock_dirname.
    """
    try:
        # Отфильтровать только файлы
        lock_files = [x for x in [os.path.join(lock_dirname, x) for x in dir_lock_filenames] if os.path.isfile(x)]
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
                    if signature['computer'] == lock_id:
                        os.remove(cur_file)
                        log.info(u'Файл блокировки <%s> удален' % cur_file)
                except:
                    log.warning(u'Не корректная сигнатура блокировки <%s>' % signature)
            except:
                if f:
                    f.close()
                log.fatal(u'Ошибка чтения сигнатуры файла блокировки <%s>' % cur_file)
    except:
        log.fatal(u'Ошибка удаления файлов блокировки. Папка блокировки <%s>' % lock_dirname)


def delMyLock(lock_id=None, lock_dirname=LOCK_DIR):
    """
    Функция рекурсивного удаления блокировок записей.
    :param lock_id: Идентификация хозяина блокировок.
    :param lock_dirname: Папка блокировок.
    """
    if not lock_id:
        lock_id = getHostName()
    return os.path.walk(lock_dirname, delMyLockInDir, lock_id)


# --- Блокировки ресурсов ---
def lockFile(filename, lock_record=None):
    """
    Блокировка файла.
    :param filename: Полное имя блокируемого файла.
    :param lock_record: Запись блокировки.
    :return: Возвращает кортеж:
        (результат выполения операции, запись блокировки).
    """
    lock_file_flag = False  # Флаг блокировки файла
    lock_rec = lock_record
    str_lock_rec = u''

    # Сгенерировать имя файла блокировки
    lock_file = os.path.splitext(filename)[0] + LOCK_FILE_EXT
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
            lock_rec = readLockRecord(lock_file)
        else:
            # выполнено без ошибки
            # Записать запись блокировки в файл
            if lock_record is not None:
                os.close(f)     # Закрыть сначала
                # Открыть для записи
                f = os.open(lock_file, os.O_WRONLY,
                            mode=stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
                if isinstance(lock_record, str):
                    str_lock_rec = lock_record
                else:
                    str_lock_rec = str(lock_record)
                os.write(f, str_lock_rec.encode())
            os.close(f)
    else:
        # Если файл заблокирован
        lock_file_flag = True
        lock_rec = readLockRecord(lock_file)

    return not lock_file_flag, lock_rec


def readLockRecord(lock_filename):
    """
    Прочитать запись блокировки из файла блокировки.
    :param lock_filename: Имя файла блокировки.
    :return: Возвращает запись блокировки или None в случае ошибки.
    """
    f = None
    lock_file = None
    try:
        lock_rec = None
        # На всякий случай преобразовать
        lock_file = os.path.splitext(lock_filename)[0] + LOCK_FILE_EXT
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


def isLockedFile(filename):
    """
    Проверка блокировки файла.
    :param filename: Имя файла.
    :return: Возвращает результат True/False.
    """
    # Сгенерировать имя файла блокировки
    lock_file = os.path.splitext(filename)[0] + LOCK_FILE_EXT
    return os.path.isfile(lock_file)


def getComputerName():
    """
    Имя хоста.
    :return: Получит имя компа в сети.
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
        if utilfunc.isOSWindowsPlatform():
            comp_name = textfunc.rus2lat(comp_name)
    return comp_name


def getHostName():
    """
    Получит имя компа в сети.
    """
    return getComputerName()


def unLockFile(filename, **unlock_compare):
    """
    Разблокировать файл.
    :param filename: Имя файла.
    :param unlock_compare: Условие проверки разблокировки.
        Ключ записи блокировки=значение.
        Проверка производится по 'И'.
        Если такого ключа в записи нет,
        то его значение берется None.
    :return: Возвращает результат True/False.
    """
    # Сгенерировать имя файла блокировки
    lock_file = os.path.splitext(filename)[0] + LOCK_FILE_EXT
    log.info(u'Сброс блокировки. Файл <%s> (%s). Условия снятия блокировки %s' % (lock_file,
                                                                                  os.path.exists(lock_file), unlock_compare))
    if os.path.exists(lock_file):
        if unlock_compare:
            lck_rec = readLockRecord(lock_file)
            # Если значения по указанным ключам равны, то все ОК
            can_unlock = bool(len([key for key in unlock_compare.keys() if lck_rec.setdefault(key, None) == unlock_compare[key]]) == len(unlock_compare))
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


def _unLockFileWalk(args, cur_dir, cur_names):
    """
    Вспомогательная функция разблокировки файла на уровне каталога по имени
    компьютера. Используется в функции os.path.walk().
    :param args: Кортеж (Имя компьютера файлы которого нужно раблокировать,
        Имя пользователя).
    :param cur_dir: Текущий директорий.
    :param cur_names: Имена поддиректорий и файлов в текущей директории.
    """
    computer_name = args[0]
    user_name = args[1]
    # Отфильтровать только файлы блокировок
    lock_files = [x for x in [os.path.join(cur_dir, x) for x in cur_names] if os.path.isfile(x) and os.path.splitext(x)[1] == LOCK_FILE_EXT]
    # Выбрать только свои файлы-блокировки
    for cur_file in lock_files:
        lock_record = readLockRecord(cur_file)
        if not user_name:
            if lock_record['computer'] == computer_name:
                os.remove(cur_file)
                log.info(u'Блокировка снята <%s>' % cur_file)
        else:
            if lock_record['computer'] == computer_name and \
               lock_record['user'] == user_name:
                os.remove(cur_file)
                log.info(u'Блокировка снята <%s>' % cur_file)


def unLockAllFile(lock_dirname, computer_name=None, username=None):
    """
    Разблокировка всех файлов.
    :param lock_dirname: Директория блокировок.
    :param computer_name: Имя компьютера файлы которого нужно раблокировать.
    :return: Возвращает результат True/False.
    """
    if not computer_name:
        computer_name = getComputerName()
    if not username:
        import ic.engine.ic_user
        username = ic.engine.ic_user.icGet('UserName')
    if lock_dirname:
        return os.walk(lock_dirname, _unLockFileWalk, (computer_name, username))


# --- Система блокировки произвольных ресурсов ---
class icLockSystem:
    """
    Система блокировки произвольных ресурсов.
    """

    def __init__(self, lock_dirname=None):
        """
        Конструктор.
        :param lock_dirname: Папка блокировки.
        """
        if lock_dirname is None:
            lock_dirname = LOCK_DIR
        
        self._LockDir = lock_dirname
        
    # --- Папочные блокировки ---
    def lockDirRes(self, lock_name):
        """
        Поставить блокировку в виде директории.
        :param lock_name: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к директории.
        """
        pass
        
    def unLockDirRes(self, lock_name):
        """
        Убрать блокировку в виде директории.
        :param lock_name: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к директории.
        """
        pass

    # --- Файловые блокировки ---
    def _getLockFileName(self, lock_name):
        """
        Определитьимя файла блокировки по имени блокировки.
        :param lock_name: Имя блокировки.
        """
        lock_name = DEFAULT_LOCK_NAME
        lock_file_name = ''
        try:
            if isinstance(lock_name, list):
                lock_name = lock_name[-1]
            elif isinstance(lock_name, str):
                lock_name = os.path.splitext(os.path.basename(lock_name))[0]
            lock_file_name = os.path.join(self._LockDir, lock_name+LOCK_FILE_EXT)
            return lock_file_name
        except:
            log.error(u'Папка блокировки <%s>. Имя блокировки <%s>' % (self._LockDir, lock_name))
            log.fatal(u'Ошибка определения имени файла блокировки')
        return lock_file_name
        
    def lockFileRes(self, lock_name, lock_record=None):
        """
        Поставить блокировку в виде файла.
        :param lock_name: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к файлу.
        :param lock_record: Запись блокировки.
        """
        lock_file_name = self._getLockFileName(lock_name)
        if lock_record is None:
            import ic.engine.ic_user
            lock_record = {'computer': getComputerName(),
                        'user': ic.engine.ic_user.icGet('UserName')}
        return lockFile(lock_file_name, lock_record)
        
    def unLockFileRes(self, lock_name):
        """
        Убрать блокировку в виде файла.
        :param lock_name: Имя блокировки.
            М.б. реализовано в виде списка имен,
            что определяет путь к файлу.
        """
        lock_file_name = self._getLockFileName(lock_name)
        return unLockFile(lock_file_name)

    def isLockFileRes(self, lock_name):
        """
        Существует ли файловая блокировка с именем.
        :param lock_name: Имя блокировки.
        """
        lock_file_name = self._getLockFileName(lock_name)
        return isLockedFile(lock_file_name)
    
    def getLockRec(self, lock_name):
        """
        Определить запись блокировки.
        :param lock_name: Имя блокировки.
        """
        lock_file_name = self._getLockFileName(lock_name)
        return readLockRecord(lock_file_name)
    
    # --- Общие функции блокировки ---
    def isLockRes(self, lock_name):
        """
        Существует ли блокировка с именем.
        :param lock_name: Имя блокировки.
        """
        pass
        
    def unLockAllMy(self):
        """
        Разблокировать все мои блокировки.
        """
        return unLockAllFile(self._LockDir, getComputerName())
