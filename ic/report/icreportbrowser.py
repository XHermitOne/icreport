#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Модуль приложения генератора отчетов.
"""

# Подключение библиотек
import os
import os.path

import wx
import wx.lib.buttons

from ic.std.utils import ini
from ic.std.dlg import dlg
from ic.std.img import bmp
from ic.std.log import log
from ic.std.utils import filefunc
from ic.std.utils import res
from ic.std.utils import ic_mode

import ic.report
from ic.report import report_generator
from ic import config

__version__ = (0, 0, 1, 4)

# Константы
# Индексы полей списка кортежей

REP_FILE_IDX = 0        # полное имя файла/директория отчета
REP_NAME_IDX = 1        # имя отчета/директория
REP_DESCRIPT_IDX = 2    # описание отчета/директория
REP_ITEMS_IDX = 3       # вложенные объекты
REP_IMG_IDX = 4         # Образ отчета в дереве отчетов

# Режимы работы браузера
IC_REPORT_VIEWER_MODE = 0
IC_REPORT_EDITOR_MODE = 1

# Позиции и размеры кнопок управления
REP_BROWSER_BUTTONS_POS_X = 780
REP_BROWSER_BUTTONS_WIDTH = 200
REP_BROWSER_BUTTONS_HEIGHT = 30


def getReportList(ReportDir_, is_sort=True):
    """
    Получить список отчетов.
    @param ReportDir_: Директорий, в котором лежат файлы отчетов.
    @type is_sort: bool.
    @param is_sort: Сортировать список по именам?
    @return: Возвращает список списков следующего формата:
        [
        [полное имя файла/директория отчета,имя отчета/директория,
            описание отчета/директория,None/вложенные объекты,индекс образа],
        .
        .
        .
        ]
        Описание директория берется из файла descript.ion,
        который должен находится в этой же директории.
        Если такой файл не найден, то описание директория - пустое.
        Вложенные объекты список, элементы которого имеют такую же структуру.
    """
    try:
        # Коррекция аргументов
        ReportDir_ = os.path.abspath(os.path.normpath(ReportDir_))

        # Выходой список
        dir_list = list()
        rep_list = list()

        # Сначала обработать под-папки
        sub_dirs = filefunc.getSubDirsFilter(ReportDir_)

        # то записать информацию в выходной список о директории
        img_idx = 0
        for sub_dir in sub_dirs:
            description_file = None
            try:
                description_file = open(sub_dir+'/descript.ion', 'r')
                dir_description = description_file.read()
                description_file.close()
            except:
                if description_file:
                    description_file.close()
                dir_description = sub_dir

            # Для поддиректориев рекурсивно вызвать эту же функцию
            data = [sub_dir, os.path.basename(sub_dir), dir_description,
                    getReportList(sub_dir, is_sort), img_idx]
            dir_list.append(data)
        if is_sort:
            # ВНИМАНИЕ! Сортировка по 3-й колонке
            dir_list.sort(key=lambda i: i[2])

        # Получить список всех файлов
        file_rep_list = [filename for filename in filefunc.getFilesByExt(ReportDir_, '.rprt')
                         if filename[-8:].lower() != '_pkl.rprt']

        for rep_file_name in file_rep_list:
            # записать данные о этом файле в выходной список
            rep_struct = res.loadResourceFile(rep_file_name, bRefresh=True)
            # Определение образа
            img_idx = 2
            try:
                if rep_struct['generator'][-3:].lower() == 'xml':
                    img_idx = 1
            except:
                log.warning('Error read report type')
            # Данные
            try:
                data = [rep_file_name, rep_struct['name'],
                        rep_struct['description'], None, img_idx]
                rep_list.append(data)
            except:
                log.fatal(u'Ошибка чтения шаблона отчета <%s>' % rep_file_name)
        if is_sort:
            # ВНИМАНИЕ! Сортировка по 3-й колонке
            rep_list.sort(key=lambda i: i[2])

        return dir_list + rep_list
    except:
        # Вывести сообщение об ошибке в лог
        log.fatal(u'Ошибка заполнения информации о файлах отчетов <%s>.' % ReportDir_)


def get_root_dirname():
    """
    Путь к корневой папке.
    """
    cur_dirname = os.path.dirname(__file__)
    if not cur_dirname:
        cur_dirname = os.getcwd()
    return os.path.dirname(os.path.dirname(cur_dirname))


def get_img_dirname():
    """
    Путь к папке образов.
    """
    cur_dirname = os.path.dirname(__file__)
    if not cur_dirname:
        cur_dirname = os.getcwd()
    return cur_dirname+'/img/'


class icReportBrowserPrototype:
    """
    Форма браузера отчетов.
    """

    def __init__(self, Mode_=IC_REPORT_VIEWER_MODE, report_dir=''):
        """
        Конструктор.
        """
        # Папка отчетов
        self._ReportDir = report_dir

        self.icon = None
        self.SetIcon(self.get_icon_obj())

        # Строка папки отчетов
        self.dir_txt = wx.StaticText(self, id=wx.NewId(),
                                     label='',
                                     pos=wx.Point(10, 10), size=wx.DefaultSize,
                                     style=0)

        # Если папки отчетов не определена или не существует, то ...
        if not self._ReportDir or not os.path.exists(self._ReportDir):
            # ... считать путь к папке отчетов из файла настройки
            self._ReportDir = ini.loadParamINI(self.getReportSettingsINIFile(), 'REPORTS', 'report_dir')
            if not self._ReportDir or not os.path.exists(self._ReportDir):
                self._ReportDir = dlg.getDirDlg(self,
                                                u'Папка отчетов <%s> не найдена. Выберите папку отчетов.' %  self._ReportDir)
                # Сохранить сразу в конфигурационном файле
                if self._ReportDir:
                    self._ReportDir = os.path.normpath(self._ReportDir)
                    ini.saveParamINI(self.getReportSettingsINIFile(),
                                     'REPORTS', 'report_dir', self._ReportDir)
        # Отобразить новый путь в окне
        self.dir_txt.SetLabel(self._ReportDir)

        # Список отчетов
        self.rep_tree = wx.TreeCtrl(self, wx.NewId(),
                                    pos=wx.Point(10, 30), size=wx.Size(750, 390), style=wx.TR_HAS_BUTTONS,
                                    validator=wx.DefaultValidator, name='ReportTree')

        self.img_list = wx.ImageList(16, 16)
        self.img_list.Add(bmp.createBitmap(os.path.join(get_img_dirname(), 'reports.png')))
        self.img_list.Add(bmp.createBitmap(os.path.join(get_img_dirname(), 'report-excel.png')))
        self.img_list.Add(bmp.createBitmap(os.path.join(get_img_dirname(), 'report.png')))
        self.rep_tree.AssignImageList(self.img_list)

        self.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnSelectChanged, id=self.rep_tree.GetId())

        # Кнопки управления

        # Кнопка вывода отчета/предварительного просмотра/печати
        self.rep_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                             bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                           'document-search-result.png')),
                                                             u'Предв. просмотр',
                                                             size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                   REP_BROWSER_BUTTONS_HEIGHT),
                                                             pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 30))
        self.Bind(wx.EVT_BUTTON, self.OnPreviewRepButton, id=self.rep_button.GetId())

        # Кнопка /печати
        self.print_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                               bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                             'printer.png')),
                                                               u'Печать',
                                                               size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                     REP_BROWSER_BUTTONS_HEIGHT),
                                                               pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 70))
        self.Bind(wx.EVT_BUTTON, self.OnPrintRepButton, id=self.print_button.GetId())

        # Кнопка установки параметров страницы
        self.page_setup_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                                    bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                                  'printer--pencil.png')),
                                                                    u'Параметры страницы',
                                                                    size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                          REP_BROWSER_BUTTONS_HEIGHT),
                                                                    pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 110))
        self.Bind(wx.EVT_BUTTON, self.OnPageSetupButton, id=self.page_setup_button.GetId())

        # Кнопка конвертирования отчета
        self.convert_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                                 bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                               'document-export.png')),
                                                                 u'Конвертация',
                                                                 size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                       REP_BROWSER_BUTTONS_HEIGHT),
                                                                 pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 150))
        self.Bind(wx.EVT_BUTTON, self.OnConvertRepButton, id=self.convert_button.GetId())

        if Mode_ == IC_REPORT_EDITOR_MODE:
            # Кнопка настройки
            self.set_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                                 bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                               'folder-open-document-text.png')),
                                                                 u'Папка отчетов',
                                                                 size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                       REP_BROWSER_BUTTONS_HEIGHT),
                                                                 pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 190))
            self.Bind(wx.EVT_BUTTON, self.OnSetRepDirButton, id=self.set_button.GetId())

            # Кнопка создания нового отчета
            self.new_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                                 bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                               'report--plus.png')),
                                                                 u'Создание',
                                                                 size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                       REP_BROWSER_BUTTONS_HEIGHT),
                                                                 pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 230))
            self.Bind(wx.EVT_BUTTON, self.OnNewRepButton, id=self.new_button.GetId())

            # Кнопка редактирования отчета
            self.edit_button=wx.lib.buttons.GenBitmapTextButton(self,wx.NewId(),
                                                                bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                              'report--pencil.png')),
                                                                u'Редактирование',
                                                                size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                      REP_BROWSER_BUTTONS_HEIGHT),
                                                                pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 270))
            self.Bind(wx.EVT_BUTTON, self.OnEditRepButton, id=self.edit_button.GetId())

            # Кнопка Обновления отчета
            self.convert_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                                     bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                                   'arrow-circle-double.png')),
                                                                     u'Обновление',
                                                                     size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                           REP_BROWSER_BUTTONS_HEIGHT),
                                                                     pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 310))
            self.Bind(wx.EVT_BUTTON, self.OnUpdateRepButton, id=self.convert_button.GetId())

            # Кнопка модуля отчета
            self.module_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                                    bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                                  'script-attribute-p.png')),
                                                                    u'Модуль отчета',
                                                                    size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                          REP_BROWSER_BUTTONS_HEIGHT),
                                                                    pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 350))
            self.Bind(wx.EVT_BUTTON, self.OnModuleRepButton, id=self.module_button.GetId())
                
        # Кнопка выхода
        self.exit_button = wx.lib.buttons.GenBitmapTextButton(self, wx.NewId(),
                                                              bmp.createBitmap(os.path.join(get_img_dirname(),
                                                                                            'door-open-out.png')),
                                                              u'Выход',
                                                              size=(REP_BROWSER_BUTTONS_WIDTH,
                                                                    REP_BROWSER_BUTTONS_HEIGHT),
                                                              pos=wx.Point(REP_BROWSER_BUTTONS_POS_X, 390),
                                                              style=wx.ALIGN_LEFT)
        self.Bind(wx.EVT_BUTTON, self.OnExitButton, id=self.exit_button.GetId())

        # Заполнить дерево отчетов
        self._fillReportTree(self._ReportDir)

    def get_icon_obj(self):
        """
        Функция получения иконки формы браузера.
        """
        if self.icon is None:
            icon_filename = os.path.join(get_img_dirname(), 'reports-stack.png')
            self.icon = wx.Icon(icon_filename, wx.BITMAP_TYPE_PNG)
        return self.icon

    def getReportSettingsINIFile(self):
        """
        Определить имя конфигурационного файла, 
            в котором хранится путь к папке отчетов.
        """
        return os.path.join(get_root_dirname(), 'settings.ini')
        
    def OnPreviewRepButton(self, event):
        """
        Обработчик нажатия кнопки 'Предварительный просмотр/Печать'.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        log.debug(u'Preview <%s>' % item_data[REP_FILE_IDX])
        # Если это файл отчета, то получить его
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            # Получение отчета
            report_generator.getReportGeneratorSystem(item_data[REP_FILE_IDX],
                                                      ParentForm_=self,
                                                      bRefresh=True).Preview()
        # Если это папка, то вывести сообщение
        else:
            dlg.getMsgBox(u'Необходимо выбрать отчет!', parent=self)
            
    def OnPrintRepButton(self, event):
        """
        Обработчик нажатия кнопки 'Печать отчета'.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        log.debug(u'Print <%s>' % item_data[REP_FILE_IDX])
        # Если это файл отчета, то получить его
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            # Получение отчета
            report_generator.getReportGeneratorSystem(item_data[REP_FILE_IDX],
                                                      ParentForm_=self,
                                                      bRefresh=True).Print()
        # Если это папка, то вывести сообщение
        else:
            dlg.getMsgBox(u'Необходимо выбрать отчет!', parent=self)
        event.Skip()

    def OnPageSetupButton(self, event):
        """
        Обработчик нажатия кнопки 'Параметры страницы'.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        # Если это файл отчета, то получить его
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            # Получение отчета
            report_generator.getReportGeneratorSystem(item_data[REP_FILE_IDX], ParentForm_=self).PageSetup()
        # Если это папка, то вывести сообщение
        else:
            dlg.getMsgBox(u'Необходимо выбрать отчет!', parent=self)

    def OnSetRepDirButton(self, event):
        """
        Обработчик нажатия кнопки 'Папка отчетов'.
        """
        # Считать путь к папке отчетов из файла настройки
        self._ReportDir = ini.loadParamINI(self.getReportSettingsINIFile(), 'REPORTS', 'report_dir')
        # Выбрать папку отчетов
        dir_dlg = wx.DirDialog(self, u'Выберите путь к папке отчетов:',
                               style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        # Установка пути по умолчанию
        if self._ReportDir:
            dir_dlg.SetPath(self._ReportDir)
        if dir_dlg.ShowModal() == wx.ID_OK:
            self._ReportDir = dir_dlg.GetPath()
        
        dir_dlg.Destroy()        

        # Сохранить новую выбранную папку
        ok = ini.saveParamINI(self.getReportSettingsINIFile(), 'REPORTS', 'report_dir', self._ReportDir)
            
        if ok is True:
            # Отобразить новый путь в окне
            self.dir_txt.SetLabel(self._ReportDir)
            # и обновить дерево отчетов
            self._fillReportTree(self._ReportDir)

    def OnExitButton(self, event):
        """
        Обработчик нажатия кнопки 'Выход'.
        """
        self.Close()

    def OnNewRepButton(self, event):
        """
        Обработчик нажатия ккнопки 'Новый отчет'.
        """
        # Запустить редакторы
        report_generator.getCurReportGeneratorSystem().New(self._ReportDir)
        event.Skip()

    def OnEditRepButton(self, event):
        """
        Обработчик нажатия ккнопки 'Редактирование отчета'.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        # Если это файл отчета, то запустить на редактирование
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            # Запустить на редактирование
            rep_generator = report_generator.getReportGeneratorSystem(item_data[REP_FILE_IDX], ParentForm_=self)
            if rep_generator is not None:
                rep_generator.Edit(item_data[0])
            else:
                log.warning('Not defaine Report Generator. Type <%s>' % item_data[REP_FILE_IDX])

        event.Skip()

    def OnUpdateRepButton(self, event):
        """
        Обработчик нажатия кнопки 'Обновление отчета'.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        # Если это файл отчета, то запустить обновление шаблона
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            # Запустить обновление шаблона отчета
            log.debug('Update <%s> report' % item_data[0])
            report_generator.getReportGeneratorSystem(item_data[REP_FILE_IDX], ParentForm_=self).Update(item_data[0])
        else:
            report_generator.getCurReportGeneratorSystem(self).Update()
                
        # Заполнить дерево отчетов
        self._fillReportTree(self._ReportDir)

        event.Skip()

    def OnConvertRepButton(self, event):
        """
        Обработчик нажатия кнопки 'Конвертация'.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        log.debug(u'Convert <%s>' % item_data[REP_FILE_IDX])
        # Если это файл отчета, то получить его
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            # Получение отчета
            report_generator.getReportGeneratorSystem(item_data[REP_FILE_IDX],
                                                      ParentForm_=self,
                                                      bRefresh=True).Convert()
        else:
            dlg.getMsgBox(u'Необходимо выбрать отчет!', parent=self)

        event.Skip()

    def OnModuleRepButton(self, event):
        """
        Обработчик нажатия кнопки 'Модуль отчета'.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        # Если это файл отчета, то запустить открытие модуля отчета
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            report_generator.getReportGeneratorSystem(item_data[REP_FILE_IDX],
                                                      ParentForm_=self).OpenModule(item_data[REP_FILE_IDX])
        else:
            dlg.getMsgBox(u'Необходимо выбрать отчет!', parent=self)

        event.Skip()

    def OnRightMouseClick(self, event):
        """
        Нажаитие правой кнопки мыши на дереве.
        """
        # Создать всплывающее меню
        popup_menu = wx.Menu()
        id_rename = wx.NewId()
        popup_menu.Append(id_rename, u'Переименовать отчет')
        self.Bind(wx.EVT_MENU, self.OnRenameReport, id=id_rename)
        self.rep_tree.PopupMenu(popup_menu,event.GetPosition())

    def OnRenameReport(self, event):
        """
        Переименовать отчет.
        """
        # Определить выбранный пункт дерева
        item = self.rep_tree.GetSelection()
        item_data = self.rep_tree.GetPyData(item)
        # Если это файл отчета, то переименовать его
        if item_data is not None and item_data[REP_ITEMS_IDX] is None:
            old_rep_name = os.path.splitext(os.path.split(item_data[REP_FILE_IDX])[1])[0]
            new_rep_name = dlg.getTextInputDlg(self, u'Переименование отчета',
                                               u'Введите новое имя отчета', old_rep_name)
            # Если имя введено и имя не старое, то переименовать
            if new_rep_name and new_rep_name != old_rep_name:
                new_rep_file_name = os.path.join(os.path.split(item_data[REP_FILE_IDX])[0],
                                                 new_rep_name+'.rprt')
                # Если новый файл не существует, то переименовать старый
                if not os.path.isfile(new_rep_file_name):
                    self.RenameReport(item_data[REP_FILE_IDX], new_rep_name)
                else:
                    dlg.getMsgBox(self, u'Невозможно поменять имя отчета. Отчет с таким именем уже существует.')

        event.Skip()

    def RenameReport(self, RepFileName_, NewName_):
        """
        Переименовать отчет.
        """
        old_name = os.path.splitext(os.path.split(RepFileName_)[1])[0]
        old_rep_file_name = RepFileName_
        old_rep_pkl_file_name = os.path.splitext(old_rep_file_name)[0] + '_pkl.rprt'
        old_xls_file_name = os.path.splitext(old_rep_file_name)[0] + '.xls'
        new_rep_file_name = os.path.join(os.path.split(old_rep_file_name)[0],
                                         NewName_+'.rprt')
        if os.path.isfile(old_rep_file_name):
            try:
                os.rename(old_rep_file_name, new_rep_file_name)
            except:
                log.fatal(u'Ошибка переименования файла <%s>' % old_rep_file_name)
            # Убить пикловский файл
            if os.path.isfile(old_rep_pkl_file_name):
                os.remove(old_rep_pkl_file_name)
            # Поменять имя в файле отчета.
            report = ic.utils.util.readAndEvalFile(new_rep_file_name, bRefresh=True)
            report['name'] = NewName_
            try:
                rep_file = open(new_rep_file_name, 'w')
                rep_file.write(str(report))
                rep_file.close()
            except:
                rep_file.close()

        new_xls_file_name = os.path.join(os.path.split(old_rep_file_name)[0],
                                         NewName_+'.xls')
        if os.path.isfile(old_xls_file_name):
            os.rename(old_xls_file_name, new_xls_file_name)
            # Поменять имя листа отчета.
            try:
                # Установить связь с Excel
                excel_app = win32com.client.Dispatch('Excel.Application')
                # Сделать приложение невидимым
                excel_app.Visible = 0
                # Открыть
                rep_tmpl = new_xls_file_name.replace('./', os.getcwd()+'/')
                rep_tmpl_book = excel_app.Workbooks.Open(rep_tmpl)
                rep_tmpl_sheet = rep_tmpl_book.Worksheets(old_name)
                # Переименовать лист
                rep_tmpl_sheet.Name = NewName_
                # Сохранить и закрыть
                rep_tmpl_book.Save()
                excel_app.Quit()
            except pythoncom.com_error:
                # Вывести сообщение об ошибке в лог
                log.fatal(u'Ошибка переименования файла')

    def OnSelectChanged(self, event):
        """
        Изменение выделенного компонента.
        """
        pass
        
    def _fillReportTree(self, ReportDir_):
        """
        Наполнить дерево отчетов данными об отчетах.
        @param ReportDir_: Директория отчетов.
        """
        # Получить описание всех отчетов
        rep_data = getReportList(ReportDir_)
        if rep_data is None:
            log.warning(u'Данные не прочитались. Папка отчетов <%s>' % ReportDir_)
            return
        # Удалить все пункты
        self.rep_tree.DeleteAllItems()
        # Корень
        root = self.rep_tree.AddRoot(u'Отчеты', image=0)
        self.rep_tree.SetPyData(root, None)
        # Добавить пункты дерева по полученному описанию отчетов
        self._appendItemsReportTree(root, rep_data)
        # Развернуть дерево
        self.rep_tree.Expand(root)

    def _appendItemsReportTree(self, ParentId_, ItemsData_):
        """
        Добавить пункты дерева по полученному описанию отчетов.
        @param ParentId_: Идентификатор родительского узла.
        @param ItemsData_: Ветка описаний отчетов.
        """
        # Перебрать описания отчетов в описании.
        for item_data in ItemsData_:
            item = self.rep_tree.AppendItem(ParentId_, item_data[REP_DESCRIPT_IDX], -1, -1, data=None)
            # Если описание папки отчетов, то доавить рекурсивно ветку
            if item_data[REP_ITEMS_IDX] is not None:
                self._appendItemsReportTree(item, item_data[REP_ITEMS_IDX])
                # Добавить изображение папки
                self.rep_tree.SetItemImage(item, 0, wx.TreeItemIcon_Normal)
                self.rep_tree.SetItemImage(item, 0, wx.TreeItemIcon_Selected)
            else:
                # Добавить изображение отчета
                self.rep_tree.SetItemImage(item, item_data[REP_IMG_IDX], wx.TreeItemIcon_Normal)
                self.rep_tree.SetItemImage(item, item_data[REP_IMG_IDX], wx.TreeItemIcon_Selected)
            # Добавить связь с данными
            self.rep_tree.SetPyData(item, item_data)

    def SetReportDir(self, Dir_):
        """
        Установить директорий/папку отчетов.
        @param Dir_: Папка отчетов.
        """
        self._ReportDir = Dir_

    def GetReportDir(self):
        """
        Папка отчетов.
        """
        return self._ReportDir
        

# Размер диалогового окна
REP_BROWSER_DLG_WIDTH = 1000
REP_BROWSER_DLG_HEIGHT = 460

# Заголовок
TITLE = 'icReport'


class icReportBrowserDialog(icReportBrowserPrototype, wx.Dialog):
    """
    Диалоговое окно браузера отчетов.
    """

    def __init__(self, ParentForm_=None, Mode_=IC_REPORT_VIEWER_MODE, report_dir=''):
        """
        Конструктор.
        @param ParentForm_: Родительская форма.
        @param Mode_: Режим работы.
        @param report_dir: Папка отчетов.
        """
        # Версия в строковом виде
        ver = '.'.join([str(ident) for ident in config.__version__])
        # Создать экземпляр главного окна
        wx.Dialog.__init__(self, ParentForm_, wx.NewId(),
                           title=u'%s. Система управления отчетами. v. %s' % (TITLE, ver),
                           pos=wx.DefaultPosition, size=wx.Size(REP_BROWSER_DLG_WIDTH, REP_BROWSER_DLG_HEIGHT),
                           style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION)

        icReportBrowserPrototype.__init__(self, Mode_, report_dir)
