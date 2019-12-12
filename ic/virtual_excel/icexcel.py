#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import os.path
import copy

try:
    from . import icprototype
    from . import icworkbook
    from . import icods
    # from . import config
except ImportError:
    # Для запуска тестов
    import icprototype
    import icworkbook
    import icods
    # import config

try:
    # Если Virtual Excel работает в окружении icReport
    from ic.std.convert import xml2dict
    from ic.std.convert import dict2xml
    from ic.std.log import log
except ImportError:
    # Если Virtual Excel работает в окружении DEFIS
    from ic.convert import xml2dict
    from ic.convert import dict2xml
    from ic.log import log


__version__ = (0, 1, 2, 1)


class icVExcel(icprototype.icVPrototype):
    """
    Виртуальное представления объектной модели Excel.
    """
    def __init__(self, encoding='utf-8', *args, **kwargs):
        """
        Конструктор.
        """
        icprototype.icVPrototype.__init__(self, None, *args, **kwargs)

        # Данные активной книги
        self._data = {'name': 'Excel', 'children': []}

        # Словарь открытых книг
        self._workbooks = {}

        # Имя файла
        self.SpreadsheetFileName = None

        # Внутренняя кодировка
        self.encoding = encoding

        # Внутренний клиппорд для операций с листами
        self._worksheet_clipboard = {}
        # Признак того что после вставки старый лист нужно удалить
        self._is_cut_worksheet = False
        # Список групповых операций с листами
        self._worksheet_list_clipboard = []

    def _regWorkbook(self, xml_filename=None, workbook_data=None):
        """
        Зарегистрировать книгу как открытую.
        """
        self._workbooks[xml_filename] = workbook_data

    def _unregWorkbook(self, xml_filename=None):
        """
        Убрать из зарегистрированных книгу.
        """
        if xml_filename in self._workbooks:
            del self._workbooks[xml_filename]

    def _reregWorkbook(self, old_xml_filename=None, new_xml_filename=None):
        """
        Зарегистрировать под новым именем зарегистрированную книгу.
        """
        if old_xml_filename in self._workbooks:
            self._workbooks[new_xml_filename] = self._workbooks[old_xml_filename]
            del self._workbooks[old_xml_filename]

    def createNew(self):
        """
        Новый.
        """
        # Данные
        self._data = {'name': 'Excel', 'children': []}

        # Имя XML файла
        self.SpreadsheetFileName = None

        self._regWorkbook(self.SpreadsheetFileName, self._data)

    def getFileName(self):
        """
        Имя файла.
        """
        return self.SpreadsheetFileName
    
    def convertXLS2XML(self, xls_filename=None):
        """
        Сконвертировать XLS файл в XML. Конвертация будет происходить
        только при установленном Excel.

        :return: True - все нормально сконвертировалось, False - ошибка.
        """
        try:
            import win32com.client
            xlXMLSpreadsheet = 46   # 0x2e # from enum XlFileFormat
        except ImportError:
            print('import win32com error!')
            return False
        try:
            # Установить связь с Excel
            excel_app = win32com.client.Dispatch('Excel.Application')
            # Сделать приложение невидимым
            # Закрыть все книги
            excel_app.Workbooks.Close()
            # Открыть XLS
            wrkbook = excel_app.Workbooks.Open(xls_filename)
            # Сохранить как XML
            xml_file_name = os.path.splitext(xls_filename)[0] + '.xml'
            wrkbook.saveAs(xml_file_name,
                           FileFormat=xlXMLSpreadsheet,
                           ReadOnlyRecommended=False,
                           CreateBackup=False)
            excel_app.Workbooks.Close()
            return True
        except:
            log.error('convertXLS2XML function')
            return False

    def loadXML(self, xml_filename=None):
        """
        Загрузить из XML файла.
        """
        if xml_filename:
            self.SpreadsheetFileName = os.path.abspath(xml_filename)
            xls_file_name = os.path.splitext(self.SpreadsheetFileName)[0]+'.xls'
            if not os.path.exists(self.SpreadsheetFileName) and os.path.exists(xls_file_name):
                if not self.convertXLS2XML(xls_file_name):
                    return None
        self._data = xml2dict.XmlFile2Dict(self.SpreadsheetFileName, encoding=self.encoding)

        # Зарегистрировать открытую книгу
        self._regWorkbook(self.SpreadsheetFileName, self._data)

        return self._data

    def loadODS(self, ods_filename=None):
        """
        Загрузить из ODS файла.

        :param ods_filename: Полное имя ODS файла.
        """
        if ods_filename:
            self.SpreadsheetFileName = os.path.abspath(ods_filename)
        
            ods = icods.icODS()
            self._data = ods.load(ods_filename)
        
            # Зарегистрировать открытую книгу
            self._regWorkbook(self.SpreadsheetFileName, self._data)
        
        return self._data
        
    def load(self, filename):
        """
        Загрузить данные из файла. Тип файла определяется по расширению.

        :param filename: Полное имя файла.
        """
        if filename:
            filename = os.path.abspath(filename)
        else:
            filename = os.path.abspath(self.SpreadsheetFileName)
            
        if (not filename) or (not os.path.exists(filename)):
            log.warning(u'Не возможно загрузить файл <%s>' % filename)
            return None
        
        ext = os.path.splitext(filename)[1]
        if ext in ('.ODS', '.ods', '.Ods'):
            return self.loadODS(filename)
        elif ext in ('.XML', '.xml', '.Xml'):
            return self.loadXML(filename)
        else:
            log.warning(u'Не поддерживаемый тип файла <%s>' % ext)
        return None
            
    def save(self):
        """
        Сохранить.
        """
        return self.saveAs()
        
    def saveAs(self, filename=None):
        """
        Сохранить данные в файл.

        :param filename: Полное имя файла.
        """
        if not filename:
            filename = self.SpreadsheetFileName
        
        ext = os.path.splitext(filename)[1]
        if ext in ('.ODS', '.ods', '.Ods'):
            return self.saveAsODS(filename)
        elif ext in ('.XML', '.xml', '.Xml'):
            return self.saveAsXML(filename)
        else:
            log.warning(u'Не поддерживаемый тип файла <%s>' % ext)
        return None                
        
    def saveAsODS(self, ods_filename=None):
        """
        Сохранить в ODS файле.

        :param ods_filename: Полное имя ODS файла.
        """
        if ods_filename is None:
            ods_filename = os.path.splitext(self.SpreadsheetFileName)[0] + '.ods'
            
        if ods_filename:
            self._reregWorkbook(self.SpreadsheetFileName, ods_filename.strip())
            self.SpreadsheetFileName = ods_filename.strip()
            
        if os.path.exists(ods_filename):
            # Если файл существует, то удалить его
            os.remove(ods_filename)
            
        ods = icods.icODS()
        return ods.save(ods_filename, self._data)

    def saveODS(self):
        """
        Сохранить в ODS файл.
        """
        return self.saveAsODS()
        
    def saveXML(self):
        """
        Сохранить в XML файл.
        """
        return self.saveAsXML()

    def saveAsXML(self, xml_filename=None):
        """
        Сохранить в XML файл.
        """
        if xml_filename:
            self._reregWorkbook(self.SpreadsheetFileName, xml_filename.strip())
            self.SpreadsheetFileName = xml_filename.strip()

        work_book = self.getActiveWorkbook()

        # Установить ExpandedRowCount и ExpandedColumnCount если нобходимо
        work_sheet_names = work_book.getWorksheetNames()
        for name in work_sheet_names:
            work_sheet = work_book.findWorksheet(name)
            if work_sheet:
                tab = work_sheet.getTable()
                tab.setExpandedRowCount()
                tab.setExpandedColCount()

        # Удалить не используемые стили
        styles = work_book.getStyles()
        styles.clearUnUsedStyles()

        save_data = self.getData()['children'][0]
        try:
            return dict2xml.dict2XmlssFile(save_data, self.SpreadsheetFileName, encoding=self.encoding)
        except IOError:
            return self.save_copy_xml(save_data, self.SpreadsheetFileName)

    def save_copy_xml(self, save_data, xml_filename, n_copy=0):
        """
        Сохранение XML файла, если не получается, то сохранить копию.

        :param save_data: Сохраняемые данные.
        :param xml_filename: Имя XML файла.
        :param n_copy: Номер копии.
        """
        try:
            xml_copy_name = os.path.splitext(xml_filename)[0] + '_' + str(n_copy) + '.xml'
            return dict2xml.dict2XmlssFile(save_data, xml_copy_name, encoding=self.encoding)
        except IOError:
            log.warning(u'XML копия <%s> <%d>' % (xml_filename, n_copy + 1))
            return self.save_copy_xml(save_data, xml_filename, n_copy + 1)

    def getData(self):
        """
        Данные.
        """
        return self._data

    def get_attributes(self):
        """
        Атрибуты.
        """
        return self._data

    def createWorkbook(self):
        """
        Создать книгу.
        """
        if self._data is None:
            self._data = {'name': 'Excel', 'children': []}
            # Зарегистрировать открытую книгу
            self._regWorkbook(self.SpreadsheetFileName, self._data)
        work_book = icworkbook.icVWorkbook(self)
        attrs = work_book.create()
        return work_book

    def getWorkbook(self, name=None):
        """
        Книга.

        :param name: Имя книги - имя XML файла.
        """
        if name is None:
            if self._data['children']:
                attrs = [element for element in self._data['children'] if element['name'] == 'Workbook']
                if attrs:
                    work_book = icworkbook.icVWorkbook(self)
                    work_book.set_attributes(attrs[0])
                    return work_book
            else:
                return self.createWorkbook()
        else:
            self.load(name)
            work_book = self.getWorkbook()
            work_book.Name = name
            return work_book
        return None

    def getActiveWorkbook(self):
        """
        Активная книга.
        """
        return self.getWorkbook()

    def openWorkbook(self, xml_filename=None):
        """
        Открыть книгу.
        Книга - XML файл таблицы Spreadsheet.
        """
        return self.load(xml_filename)

    def closeWorkbook(self, xml_filename=None):
        """
        Закрыть книгу.

        :param xml_filename: Указание XML файла закрываемой книги.
        Если не указан, то закрывается текущая книга.
        """
        xml_file_name = xml_filename
        if xml_file_name is None:
            xml_file_name = self.SpreadsheetFileName

        self._unregWorkbook(xml_file_name)

        # Если активная книга - закрываемая, то поменять активную книгу
        if self.SpreadsheetFileName.lower() == xml_file_name.lower():
            if self._workbooks:
                # Заменить активную книгу на любую из зарегестрированных
                reg_workbook_xml_file_name = self._workbooks.keys()[0]
                self.activeWorkbook(reg_workbook_xml_file_name)
            else:
                # Если больше зарегистрированных книг нет, то тогда
                # создать новую книгу
                self.createNew()

    def activeWorkbook(self, xml_filename=None):
        """
        Сделать книгу активной.
        """
        xml_file_name = u''
        try:
            xml_file_name = os.path.abspath(xml_filename)
            if os.path.exists(xml_file_name):
                self.SpreadsheetFileName = xml_file_name
                self._data = self._workbooks[xml_file_name]
            else:
                log.warning(u'Файл <%s> не существует' % xml_file_name)

        except KeyError:
            log.error(u'Книга <%s> не зарегистрирована в  <%s>' % (xml_file_name, self._workbooks.keys()))
            raise

    def _findWorkbookData(self, xml_filename=None):
        """
        Найти данные указанной книги.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        """
        # Определить книгу
        if xml_filename is None:
            workbook_data = self._data
        else:
            xml_file_name = os.path.abspath(xml_filename)
            try:
                workbook_data = self._workbooks[xml_file_name]
            except KeyError:
                log.error(u'Книга <%s> не зарегистрирована в <%s>' % (xml_file_name, self._workbooks.keys()))
                raise
        workbook_data = [data for data in workbook_data['children'] if 'name' in data and data['name'] == 'Workbook']
        if workbook_data:
            workbook_data = workbook_data[0]
        else:
            log.warning(u'Книга не определена в <%s>' % xml_filename)
            return None
        return workbook_data

    def _findWorksheetData(self, xml_filename=None, sheet_name=None):
        """
        Найти данные указанного листа.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        :param sheet_name: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        # Определить книгу
        workbook_data = self._findWorkbookData(xml_filename)
        if workbook_data is None:
            return None
        # Определить лист
        worksheet_data = [data for data in workbook_data['children'] if data['name'] == 'Worksheet']
        if worksheet_data:
            if sheet_name is None:
                worksheet_data = worksheet_data[0]
            else:
                worksheet_data = [data for data in worksheet_data if data['Name'] == sheet_name]
                if worksheet_data:
                    worksheet_data = worksheet_data[0]
                else:
                    log.warning(u'Книга <%s> не найдена' % sheet_name)
                    return None
        else:
            log.warning(u'Книга не определена в <%s>' % xml_filename)
            return None
        return worksheet_data

    def _getWorkbookStyles(self, xml_filename=None):
        """
        Данные стилей указанной книги.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        """
        # Определить книгу
        workbook_data = self._findWorkbookData(xml_filename)
        if workbook_data is None:
            return None
        # Определить стили
        styles_data = [data for data in workbook_data['children'] if data['name'] == 'Styles']
        if not styles_data:
            log.warning(u'Стили в книге <%s> не определены' % xml_filename)
            return None
        else:
            styles_data = styles_data[0]
        return styles_data

    def _genNewStyleID(self, style_id, reg_styles_id):
        """
        Сгенерировать идентификатор стиля при работе с листами.

        :param style_id: Имя стиля.
        :param reg_styles_id: Список идентификаторов уже зарегистрированных
        стилей.
        """
        i = 0
        new_style_id = style_id
        while new_style_id in reg_styles_id:
           new_style_id = style_id+str(i)
           i += 1
        return new_style_id

    def copyWorksheet(self, xml_filename=None, sheet_name=None):
        """
        Положить в внутренний буфер обмена копию листа.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        :param sheet_name: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        xml_filename = self._unificXMLFileName(xml_filename)

        sheet_name = self._unicode2str(sheet_name)

        worksheet_data = self._findWorksheetData(xml_filename, sheet_name)
        styles_data = self._getWorkbookStyles(xml_filename)
        if worksheet_data:
            # Сделать копию
            self._worksheet_clipboard = dict()
            self._worksheet_clipboard[(xml_filename, sheet_name)] = copy.deepcopy(worksheet_data)
            self._worksheet_clipboard['styles'] = copy.deepcopy(styles_data)
            self._is_cut_worksheet = False
            return self._worksheet_clipboard[(xml_filename, sheet_name)]
        return None

    def cutWorksheet(self, xml_filename=None, sheet_name=None):
        """
        Вырезать лист.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        :param sheet_name: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        xml_filename = self._unificXMLFileName(xml_filename)

        sheet_name = self._unicode2str(sheet_name)

        worksheet_data = self._findWorksheetData(xml_filename, sheet_name)
        styles_data = self._getWorkbookStyles(xml_filename)
        if worksheet_data:
            # Сделать копию
            self._worksheet_clipboard = dict()
            self._worksheet_clipboard[(xml_filename, sheet_name)] = copy.deepcopy(worksheet_data)
            self._worksheet_clipboard['styles'] = copy.deepcopy(styles_data)
            self._is_cut_worksheet = True
            return self._worksheet_clipboard[[(xml_filename, sheet_name)]]
        return None

    def delWorksheet(self, xml_filename=None, sheet_name=None):
        """
        Удалить безвозвратно лист.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        :param sheet_name: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        :return: True-лист удален, False-лист по какойто причине не удален.
        """
        xml_filename = self._unificXMLFileName(xml_filename)

        sheet_name = self._unicode2str(sheet_name)

        # Определить книгу
        workbook_data = self._findWorkbookData(xml_filename)
        if workbook_data is None:
            log.warning(u'Книга <%s> не найдена' % xml_filename)
            return False
        # Найти и удалить лист
        result = False
        for i, data in enumerate(workbook_data['children']):
            if data['name'] == 'Worksheet':
                if sheet_name is None:
                    # Если имя листа не определено, то просто удалить первый попавшийся лист
                    del workbook_data['children'][i]
                    result = True
                    break
                else:
                    # Если имя листа определено, то проверить на соответствие имен листов
                    if data['Name'] == sheet_name:
                        log.info(u'Удаление книги <%s>' % sheet_name)
                        del workbook_data['children'][i]
                        result = True
                        break
        return result

    def delWithoutWorksheet(self, xml_filename=None, sheet_name_=None):
        """
        Удалить безвозвратно все листы из книги кроме указанного.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        :param sheet_name_: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        :return: True-листы удалены, False-листы по какойто причине не удалены.
        """
        sheet_name_ = self._unicode2str(sheet_name_)

        # Определить книгу
        workbook_data = self._findWorkbookData(xml_filename)
        if workbook_data is None:
            return False
        # Найти и удалить лист
        not_del_first = True
        for i,data in enumerate(workbook_data['children']):
            if data['name'] == 'Worksheet':
                if sheet_name_ is None:
                    # Если имя листа не определено, то просто первый попавшийся лист
                    if not not_del_first:
                        del workbook_data['children'][i]
                    not_del_first = False
                else:
                    # Если имя листа определено, то проверить на соответствие имен листов
                    if data['Name'] != sheet_name_:
                        del workbook_data['children'][i]

        return True

    def delSelectedWorksheetList(self, xml_filename=None):
        """
        Удалить выбранные листы.

        :param xml_filename: Имя XML файла книги из которй удаляется.
        Если не определено, то имеется ввиду активная книга.
        """
        xml_filename = self._unificXMLFileName(xml_filename)

        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_src in enumerate(self._worksheet_list_clipboard):
                workbook_name, worksheet_name = worksheet_src

                if xml_filename == workbook_name:
                    self.delWorksheet(workbook_name, worksheet_name)

        # После опрации очистить список выбранных листов
        self._worksheet_list_clipboard = []

    def delWithoutSelectedWorksheetList(self, xml_filename=None):
        """
        Удалить все не выбранные листы из книги.

        :param xml_filename: Имя XML файла книги для из которой удаляется.
        Если не определено, то имеется ввиду активная книга.
        """
        xml_filename = self._unificXMLFileName(xml_filename)

        not_deleted_worksheet_names = [worksheet_src[1] for worksheet_src in self._worksheet_list_clipboard if xml_filename == worksheet_src[0]]

        # Определить данные книги, из которой будет производиться удаление
        workbook_name = xml_filename
        workbook_data = self._findWorkbookData(workbook_name)
        if workbook_data is None:
            return None
        # Список имен листов книги
        worksheet_name_list = [data['Name'] for data in workbook_data['children'] if data['name'] == 'Worksheet']

        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_name in enumerate(worksheet_name_list):
                if worksheet_name not in not_deleted_worksheet_names:
                    self.delWorksheet(workbook_name,worksheet_name)

    def _pasteStyleIntoWorkbook(self, style, workbook_data):
        """
        Вставить стиль в готовую структуру книги.

        :param style: Данные вставляемого стиля.
        :param workbook_data: Данные книги.
        :return: Возвращает идентификатор стиля или None
        в случае ошибки.
        """
        try:
            workbook_styles = [data for data in workbook_data['children'] if data['name'] == 'Styles'][0]
            if style not in workbook_styles['children']:
                workbook_styles_id = [data['ID'] for data in workbook_styles['children'] if data['name'] == 'Style']

                new_style_id = self._genNewStyleID(style['ID'], workbook_styles_id)
                style['ID'] = new_style_id
                workbook_styles['children'].append(style)
                return new_style_id
            else:
                # Точно такой стиль уже есть и поэтому добавлять его не надо
                return style['ID']
        except:
            log.error(u'Ошибка в функции _pasteStyleIntoWorkbook')
            raise
        return None

    def _replaceStyleID(self, data, old_style_id, new_style_id):
        """
        Заменить идентификаторы стилей в вставляемых данных.

        :param data: Данные для вставки.
        :param old_style_id: Старый идентификатор стиля.
        :param new_style_id: Новый идентификатор стиля.
        :return: Возвращает данные с поправленным стилем.
        """
        if old_style_id == new_style_id:
            # Идентификаторы равны - замены не требуется
            return data
        if 'StyleID' in data and data['StyleID'] == old_style_id:
            data['StyleID'] = new_style_id
        if 'children' in data and data['children']:
            for i, child in enumerate(data['children']):
                data['children'][i] = self._replaceStyleID(child, old_style_id, new_style_id)
        return data

    def _genNewWorksheetName(self, worksheet_name, worksheet_names):
        """
        Подобрать имя для листа, чтобы не пересекалось с уже существующими.
        """
        new_sheet_name = worksheet_name

        i = 1
        while new_sheet_name in worksheet_names:
            new_sheet_name = worksheet_name+'_'+str(i)
            i += 1
        return new_sheet_name

    def pasteWorksheet(self, xml_filename=None, is_cut=None, new_worksheet_name=None):
        """
        Вставить лист из буфера обмена в указанную книгу.

        :param xml_filename: Имя XML файла книги для вставки. Если не определено,
        то имеется ввиду активная книга.
        :param is_cut: Признак того, что старый лист нужно удалить.
        Если None, то взять системный признак.
        :param new_worksheet_name: Новое имя листа.
        :return: True-лист вставлен, False-лист по какойто причине не вставлен.
        """
        xml_filename = self._unificXMLFileName(xml_filename)
        new_worksheet_name = self._unicode2str(new_worksheet_name)

        # Если в буфере обмена ничего нет, то вставку не производить
        if not self._worksheet_clipboard:
            return False

        # Определить книгу, в которую надо вставить лист
        workbook_data = self._findWorkbookData(xml_filename)
        if workbook_data is None:
            return False

        if is_cut is None:
            is_cut = self._is_cut_worksheet
        # Определить новое имя листа
        sheet_names = [data['Name'] for data in workbook_data['children'] if data['name'] == 'Worksheet']
        worksheet_source, worksheet_data = [(key, value) for key, value in self._worksheet_clipboard.items()
                                            if isinstance(key, tuple)][0]
        if new_worksheet_name:
            sheet_name = new_worksheet_name
        else:
            sheet_name = worksheet_data['Name']
        worksheet_data['Name'] = self._genNewWorksheetName(sheet_name, sheet_names)
        # Вставить стили
        if 'styles' in self._worksheet_clipboard:
            for i, style in enumerate(self._worksheet_clipboard['styles']['children']):
                old_style_id = style['ID']
                new_style_id = self._pasteStyleIntoWorkbook(style, workbook_data)
                worksheet_data = self._replaceStyleID(worksheet_data, old_style_id, new_style_id)
        # Вставить данные
        workbook_data['children'].append(worksheet_data)

        # Если нужно, то удалить старый лист
        if is_cut:
            self.delWorksheet(*worksheet_source)
        # Очистить буфер обмена
        self._worksheet_clipboard = {}
        return True

    def moveWorksheet(self, old_xml_filename=None, sheet_name=None, new_xml_filename=None):
        """
        Переместить лист из одной книги в другую.
        """
        self.cutWorksheet(old_xml_filename, sheet_name)
        return self.pasteWorksheet(new_xml_filename)

    def selectWorksheet(self, xml_filename=None, sheet_name=None):
        """
        Выбрать лист для групповых операций с листами.

        :param xml_filename: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        :param sheet_name: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        xml_filename = self._unificXMLFileName(xml_filename)
        sheet_name = self._unicode2str(sheet_name)

        self._worksheet_list_clipboard.append((xml_filename, sheet_name))

    def getLastSelectedWorksheet(self):
        """
        Последний выбранный лист.
        """
        if self._worksheet_list_clipboard:
            worksheet_src = self._worksheet_list_clipboard[-1]
            self.activeWorkbook(worksheet_src[0])
            active_workbook = self.getActiveWorkbook()
            if active_workbook:
                return active_workbook.findWorksheet(worksheet_src[1])
        return None

    def copyWorksheetListTo(self, xml_filename=None, new_worksheet_names=None):
        """
        Копирование списка выбранных листов в книгу.

        :param xml_filename: Имя XML файла книги для вставки. Если не определено,
        то имеется ввиду активная книга.
        :param new_worksheet_names: Список имен новых листов.
        """
        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_src in enumerate(self._worksheet_list_clipboard):
                workbook_name, worksheet_name = worksheet_src
                new_worksheet_name = None
                if new_worksheet_names:
                    if i < len(new_worksheet_names):
                        new_worksheet_name = new_worksheet_names[i]
                self.copyWorksheet(workbook_name, worksheet_name)
                self.pasteWorksheet(xml_filename, new_worksheet_name=new_worksheet_name)

        # После опрации очистить список выбранных листов
        self._worksheet_list_clipboard = []

    def moveWorksheetListTo(self, xml_filename=None, new_worksheet_names=None):
        """
        Перенести список выбранных листов в книгу.

        :param xml_filename: Имя XML файла книги для вставки. Если не определено,
        то имеется ввиду активная книга.
        :param new_worksheet_names: Список имен новых листов.
        """
        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_src in enumerate(self._worksheet_list_clipboard):
                workbook_name, worksheet_name = worksheet_src
                new_worksheet_name = None
                if new_worksheet_names:
                    if i < len(new_worksheet_names):
                        new_worksheet_name = new_worksheet_names[i]
                self.cutWorksheet(workbook_name, worksheet_name)
                self.pasteWorksheet(xml_filename, new_worksheet_name=new_worksheet_name)
        # После опрации очистить список выбранных листов
        self._worksheet_list_clipboard = []

    def _unicode2str(self, UnicodeStr_):
        """
        Преобразование unicode строки в обычную строку.
        """
        # if isinstance(UnicodeStr_, unicode):
        #    return UnicodeStr_.encode(self.encoding)
        return UnicodeStr_

    def _unificXMLFileName(self, xml_filename):
        """
        Привести к внутреннему виду имя XML файла книги.
        Если имя не определено то берется имя файла активной книги.
        В качестве внутреннего имени XML файла книги
        берется АБСОЛЮТНЫЙ ПУТЬ до файла.
        """
        if xml_filename:
            xml_filename = os.path.abspath(xml_filename)
        else:
            xml_filename = self.SpreadsheetFileName
        return xml_filename

    def mergeCell(self, sheet_name, row, column, merge_down, merge_across_, xml_filename=None):
        """
        Объединить ячейки.
        """
        if xml_filename is not None:
            self.activeWorkbook(xml_filename)
        work_book = self.getActiveWorkbook()
        work_sheet = work_book.findWorksheet(sheet_name)
        table = work_sheet.getTable()
        cell = table.getCell(row, column)
        return cell.setMerge(merge_across_, merge_down)

    def setCellValue(self, sheet_name, row, column, value, xml_filename=None):
        """
        Установить значение в ячейку.
        """
        if xml_filename is not None:
            self.activeWorkbook(xml_filename)
        work_book = self.getActiveWorkbook()
        work_sheet = work_book.findWorksheet(sheet_name)
        table = work_sheet.getTable()
        cell = table.getCell(row, column)
        return cell.setValue(value)

    def setCellStyle(self, sheet_name, row, col, alignment=None,
                     left_border=None, right_border=None, top_border=None, bottom_border=None,
                     font=None, interior=None, number_format=None, xml_filename=None):
        """
        Установить стиль ячейки.
        """
        if xml_filename is not None:
            self.activeWorkbook(xml_filename)
        work_book = self.getActiveWorkbook()
        work_sheet = work_book.findWorksheet(sheet_name)
        table = work_sheet.getTable()
        cell = table.getCell(row, col)
        borders = {'name': 'Borders', 'children': []}
        if left_border:
            left_border['Position'] = 'Left'
            borders['children'].append(left_border)
        if right_border:
            right_border['Position'] = 'Right'
            borders['children'].append(right_border)
        if top_border:
            top_border['Position'] = 'Top'
            borders['children'].append(top_border)
        if bottom_border:
            bottom_border['Position'] = 'Bottom'
            borders['children'].append(bottom_border)

        return cell.setStyle(alignment=alignment, borders=borders, font=font,
                             interior=interior, number_format=number_format)

    def exec_cmd_script(self, cmd_script=None, bAutoSave=True):
        """
        Выполнить скрипт - список команд.

        :param cmd_script: Список команд формата
        [
        ('ИмяКоманды',(кортеж не именованных аргуметов),{словарь именованных аргументов}),
        ...
        ]
        :param bAutoSave: Автоматически по завершению работы сохранить.
        """
        if cmd_script:
            for cmd in cmd_script:
                try:
                    args = ()
                    len_cmd = len(cmd)
                    if len_cmd >= 2:
                        args = cmd[1]
                    kwargs = {}
                    if len_cmd >= 3:
                        kwargs = cmd[2]
                    # Непосредственный вызов функции
                    getattr(self, cmd[0])(*args, **kwargs)
                except:
                    log.error(u'Выполнение комманды <%s>' % cmd)
                    raise
        if bAutoSave:
            self.save()


if __name__ == '__main__':
    import tests
    tests.test_work_sheet()
