#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import os.path
import copy

import icprototype
import icworkbook
import icods
import config

try:
    # Если Virtual Excel работает в окружении icReport
    from ic.std.convert import xml2dict
    from ic.std.convert import dict2xml
    from ic.std.log import log
except ImportError:
    # Если Virtual Excel работает в окружении icServices
    from services.convert import xml2dict
    from services.convert import dict2xml
    from services.ic_std.log import log

__version__ = (0, 0, 1, 4)


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

    def _regWorkbook(self, XMLFileName_=None, WorkbookData_=None):
        """
        Зарегистрировать книгу как открытую.
        """
        self._workbooks[XMLFileName_] = WorkbookData_

    def _unregWorkbook(self, XMLFileName_=None):
        """
        Убрать из зарегистрированных книгу.
        """
        if XMLFileName_ in self._workbooks:
            del self._workbooks[XMLFileName_]

    def _reregWorkbook(self, OldXMLFileName_=None, NewXMLFileName_=None):
        """
        Зарегистрировать под новым именем зарегистрированную книгу.
        """
        if OldXMLFileName_ in self._workbooks:
            self._workbooks[NewXMLFileName_] = self._workbooks[OldXMLFileName_]
            del self._workbooks[OldXMLFileName_]

    def New(self):
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
    
    def convertXLS2XML(self, XLSFileName_=None):
        """
        Сконвертировать XLS файл в XML. Конвертация будет происходить
        только при установленном Excel.
        @return: True - все нормально сконвертировалось, False - ошибка.
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
            wrkbook = excel_app.Workbooks.Open(XLSFileName_)
            # Сохранить как XML
            xml_file_name = os.path.splitext(XLSFileName_)[0]+'.xml'
            wrkbook.SaveAs(xml_file_name,
                           FileFormat=xlXMLSpreadsheet,
                           ReadOnlyRecommended=False,
                           CreateBackup=False)
            excel_app.Workbooks.Close()
            return True
        except:
            log.error('convertXLS2XML function')
            return False

    def LoadXML(self, XMLFileName_=None):
        """
        Загрузить из XML файла.
        """
        if XMLFileName_:
            self.SpreadsheetFileName = os.path.abspath(XMLFileName_)
            xls_file_name = os.path.splitext(self.SpreadsheetFileName)[0]+'.xls'
            if not os.path.exists(self.SpreadsheetFileName) and os.path.exists(xls_file_name):
                if not self.convertXLS2XML(xls_file_name):
                    return None
        self._data = xml2dict.XmlFile2Dict(self.SpreadsheetFileName, encoding=self.encoding)

        # Зарегистрировать открытую книгу
        self._regWorkbook(self.SpreadsheetFileName, self._data)

        return self._data

    def LoadODS(self, ODSFileName=None):
        """
        Загрузить из ODS файла.
        @param ODSFIleName: Полное имя ODS файла.
        """
        if ODSFileName:
            self.SpreadsheetFileName = os.path.abspath(ODSFileName)
        
            ods = icods.icODS()
            self._data = ods.Load(ODSFileName)
        
            # Зарегистрировать открытую книгу
            self._regWorkbook(self.SpreadsheetFileName, self._data)
        
        return self._data
        
    def Load(self, FileName):
        """
        Загрузить данные из файла. Тип файла определяется по расширению.
        @param FileName: Полное имя файла.
        """
        if FileName:
            FileName = os.path.abspath(FileName)
        else:
            FileName = os.path.abspath(self.SpreadsheetFileName)
            
        if (not FileName) or (not os.path.exists(FileName)):
            log.warning('Can\'t load file <%s>' % FileName)
            return None
        
        ext = os.path.splitext(FileName)[1]
        if ext in ('.ODS', '.ods', '.Ods'):
            return self.LoadODS(FileName)
        elif ext in ('.XML', '.xml', '.Xml'):
            return self.LoadXML(FileName)
        else:
            log.warning('Unsupported file type: <%s>' % ext)
        return None
            
    def Save(self):
        """
        Сохранить.
        """
        return self.SaveAs()
        
    def SaveAs(self, FileName=None):
        """
        Сохранить данные в файл.
        @param FileName: Полное имя файла.
        """
        if not FileName:
            FileName = self.SpreadsheetFileName
        
        ext = os.path.splitext(FileName)[1]
        if ext in ('.ODS', '.ods', '.Ods'):
            return self.SaveAsODS(FileName)
        elif ext in ('.XML', '.xml', '.Xml'):
            return self.SaveAsXML(FileName)
        else:
            log.warning('Unsupported file type: <%s>' % ext)
        return None                
        
    def SaveAsODS(self, ODSFileName=None):
        """
        Сохранить в ODS файле.
        @param ODSFIleName: Полное имя ODS файла.
        """
        if ODSFileName is None:
            ODSFileName = os.path.splitext(self.SpreadsheetFileName)[0]+'.ods'
            
        if ODSFileName:
            self._reregWorkbook(self.SpreadsheetFileName, ODSFileName.strip())
            self.SpreadsheetFileName = ODSFileName.strip()
            
        if os.path.exists(ODSFileName):
            # Если файл существует, то удалить его
            os.remove(ODSFileName)
            
        ods = icods.icODS()
        return ods.Save(ODSFileName, self._data)

    def SaveODS(self):
        """
        Сохранить в ODS файл.
        """
        return self.SaveAsODS()
        
    def SaveXML(self):
        """
        Сохранить в XML файл.
        """
        return self.SaveAsXML()

    def SaveAsXML(self, XMLFileName_=None):
        """
        Сохранить в XML файл.
        """
        if XMLFileName_:
            self._reregWorkbook(self.SpreadsheetFileName, XMLFileName_.strip())
            self.SpreadsheetFileName = XMLFileName_.strip()

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
            return dict2xml.Dict2XmlssFile(save_data, self.SpreadsheetFileName, encoding=self.encoding)
        except IOError:
            return self.save_copy_xml(save_data, self.SpreadsheetFileName)

    def save_copy_xml(self, SaveData_, XMLFileName_, CopyNum_=0):
        """
        Сохранение XML файла, если не получается, то сохранить копию.
        @param SaveData_: Сохраняемые данные.
        @param XMLFileName_: Имя XML файла.
        @param CopyNum_: Номер копии.
        """
        try:
            xml_copy_name = os.path.splitext(XMLFileName_)[0]+'_'+str(CopyNum_)+'.xml'
            return dict2xml.Dict2XmlssFile(SaveData_, xml_copy_name, encoding=self.encoding)
        except IOError:
            log.warning('XML COPY: <%s> <%d>' % (XMLFileName_, CopyNum_+1))
            return self.save_copy_xml(SaveData_, XMLFileName_, CopyNum_+1)

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

    def getWorkbook(self, Name_=None):
        """
        Книга. Имя книги - имя XML файла.
        """
        if Name_ is None:
            if self._data['children']:
                attrs = [element for element in self._data['children'] if element['name'] == 'Workbook']
                if attrs:
                    work_book = icworkbook.icVWorkbook(self)
                    work_book.set_attributes(attrs[0])
                    return work_book
            else:
                return self.createWorkbook()
        else:
            self.Load(Name_)
            work_book = self.getWorkbook()
            work_book.Name = Name_
            return work_book
        return None

    def getActiveWorkbook(self):
        """
        Активная книга.
        """
        return self.getWorkbook()

    def openWorkbook(self, XMLFileName_=None):
        """
        Открыть книгу.
        Книга - XML файл таблицы Spreadsheet.
        """
        return self.Load(XMLFileName_)

    def closeWorkbook(self, XMLFileName_=None):
        """
        Закрыть книгу.
        @param XMLFileName_: Указание XML файла закрываемой книги.
        Если не указан, то закрывается текущая книга.
        """
        xml_file_name = XMLFileName_
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
                self.New()

    def activeWorkbook(self, XMLFileName_=None):
        """
        Сделать книгу активной.
        """
        try:
            xml_file_name = os.path.abspath(XMLFileName_)
            if os.path.exists(xml_file_name):
                self.SpreadsheetFileName = xml_file_name
                self._data = self._workbooks[xml_file_name]
            else:
                log.warning('Workbook file <%s> not exists!' % xml_file_name)

        except KeyError:
            log.error('Workbook <%s> not registered in <%s>' % (xml_file_name, self._workbooks.keys()))
            raise

    def _findWorkbookData(self, XMLFileName_=None):
        """
        Найти данные указанной книги.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        """
        # Определить книгу
        if XMLFileName_ is None:
            workbook_data = self._data
        else:
            xml_file_name = os.path.abspath(XMLFileName_)
            try:
                workbook_data = self._workbooks[xml_file_name]
            except KeyError:
                log.error('Workbook <%s> not registered in <%s>' % (xml_file_name, self._workbooks.keys()))
                raise
        workbook_data = [data for data in workbook_data['children'] if 'name' in data and data['name'] == 'Workbook']
        if workbook_data:
            workbook_data = workbook_data[0]
        else:
            log.warning('Workbook in <%s> not defined' % XMLFileName_)
            return None
        return workbook_data

    def _findWorksheetData(self, XMLFileName_=None, SheetName_=None):
        """
        Найти данные указанного листа.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        @param SheetName_: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        # Определить книгу
        workbook_data = self._findWorkbookData(XMLFileName_)
        if workbook_data is None:
            return None
        # Определить лист
        worksheet_data = [data for data in workbook_data['children'] if data['name'] == 'Worksheet']
        if worksheet_data:
            if SheetName_ is None:
                worksheet_data = worksheet_data[0]
            else:
                worksheet_data = [data for data in worksheet_data if data['Name'] == SheetName_]
                if worksheet_data:
                    worksheet_data = worksheet_data[0]
                else:
                    log.warning('Worksheet <%s> not found' % SheetName_)
                    return None
        else:
            log.warning('Worksheets in <%s> not defined' % XMLFileName_)
            return None
        return worksheet_data

    def _getWorkbookStyles(self, XMLFileName_=None):
        """
        Данные стилей указанной книги.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        """
        # Определить книгу
        workbook_data = self._findWorkbookData(XMLFileName_)
        if workbook_data is None:
            return None
        # Определить стили
        styles_data = [data for data in workbook_data['children'] if data['name'] == 'Styles']
        if not styles_data:
            log.warning('Styles in worksbook <%s> not defined' % XMLFileName_)
            return None
        else:
            styles_data = styles_data[0]
        return styles_data

    def _genNewStyleID(self, StyleID_, RegStylesID_):
        """
        Сгенерировать идентификатор стиля при работе с листами.
        @param StyleID_: Имя стиля.
        @param RegStylesID_: Список идентификаторов уже зарегистрированных
        стилей.
        """
        i = 0
        new_style_id = StyleID_
        while new_style_id in RegStylesID_:
           new_style_id = StyleID_+str(i)
           i += 1
        return new_style_id

    def copyWorksheet(self, XMLFileName_=None, SheetName_=None):
        """
        Положить в внутренний буфер обмена копию листа.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        @param SheetName_: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        XMLFileName_ = self._unificXMLFileName(XMLFileName_)

        SheetName_ = self._unicode2str(SheetName_)

        worksheet_data = self._findWorksheetData(XMLFileName_, SheetName_)
        styles_data = self._getWorkbookStyles(XMLFileName_)
        if worksheet_data:
            # Сделать копию
            self._worksheet_clipboard = dict()
            self._worksheet_clipboard[(XMLFileName_, SheetName_)] = copy.deepcopy(worksheet_data)
            self._worksheet_clipboard['styles'] = copy.deepcopy(styles_data)
            self._is_cut_worksheet = False
            return self._worksheet_clipboard[(XMLFileName_, SheetName_)]
        return None

    def cutWorksheet(self, XMLFileName_=None, SheetName_=None):
        """
        Вырезать лист.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        @param SheetName_: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        XMLFileName_ = self._unificXMLFileName(XMLFileName_)

        SheetName_ = self._unicode2str(SheetName_)

        worksheet_data = self._findWorksheetData(XMLFileName_, SheetName_)
        styles_data = self._getWorkbookStyles(XMLFileName_)
        if worksheet_data:
            # Сделать копию
            self._worksheet_clipboard = dict()
            self._worksheet_clipboard[(XMLFileName_, SheetName_)] = copy.deepcopy(worksheet_data)
            self._worksheet_clipboard['styles'] = copy.deepcopy(styles_data)
            self._is_cut_worksheet = True
            return self._worksheet_clipboard[[(XMLFileName_, SheetName_)]]
        return None

    def delWorksheet(self, XMLFileName_=None, SheetName_=None):
        """
        Удалить безвозвратно лист.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        @param SheetName_: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        @return: True-лист удален, False-лист по какойто причине не удален.
        """
        XMLFileName_ = self._unificXMLFileName(XMLFileName_)

        SheetName_ = self._unicode2str(SheetName_)

        # Определить книгу
        workbook_data = self._findWorkbookData(XMLFileName_)
        if workbook_data is None:
            log.warning('Workbook <%s> not found' % XMLFileName_)
            return False
        # Найти и удалить лист
        result = False
        for i, data in enumerate(workbook_data['children']):
            if data['name'] == 'Worksheet':
                if SheetName_ is None:
                    # Если имя листа не определено, то просто удалить первый попавшийся лист
                    del workbook_data['children'][i]
                    result = True
                    break
                else:
                    # Если имя листа определено, то проверить на соответствие имен листов
                    if data['Name'] == SheetName_:
                        log.info('Delete <%s> worksheet' % SheetName_)
                        del workbook_data['children'][i]
                        result = True
                        break
        return result

    def delWithoutWorksheet(self, XMLFileName_=None, SheetName_=None):
        """
        Удалить безвозвратно все листы из книги кроме указанного.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        @param SheetName_: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        @return: True-листы удалены, False-листы по какойто причине не удалены.
        """
        SheetName_ = self._unicode2str(SheetName_)

        # Определить книгу
        workbook_data = self._findWorkbookData(XMLFileName_)
        if workbook_data is None:
            return False
        # Найти и удалить лист
        not_del_first = True
        for i,data in enumerate(workbook_data['children']):
            if data['name'] == 'Worksheet':
                if SheetName_ is None:
                    # Если имя листа не определено, то просто первый попавшийся лист
                    if not not_del_first:
                        del workbook_data['children'][i]
                    not_del_first = False
                else:
                    # Если имя листа определено, то проверить на соответствие имен листов
                    if data['Name'] != SheetName_:
                        del workbook_data['children'][i]

        return True

    def delSelectedWorksheetList(self, XMLFileName_=None):
        """
        Удалить выбранные листы.
        @param XMLFileName_: Имя XML файла книги из которй удаляется.
        Если не определено, то имеется ввиду активная книга.
        """
        XMLFileName_ = self._unificXMLFileName(XMLFileName_)

        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_src in enumerate(self._worksheet_list_clipboard):
                workbook_name, worksheet_name = worksheet_src

                if XMLFileName_ == workbook_name:
                    self.delWorksheet(workbook_name, worksheet_name)

        # После опрации очистить список выбранных листов
        self._worksheet_list_clipboard = []

    def delWithoutSelectedWorksheetList(self, XMLFileName_=None):
        """
        Удалить все не выбранные листы из книги.
        @param XMLFileName_: Имя XML файла книги для из которой удаляется.
        Если не определено, то имеется ввиду активная книга.
        """
        XMLFileName_ = self._unificXMLFileName(XMLFileName_)

        not_deleted_worksheet_names = [worksheet_src[1] for worksheet_src in self._worksheet_list_clipboard if XMLFileName_ == worksheet_src[0]]

        # Определить данные книги, из которой будет производиться удаление
        workbook_name = XMLFileName_
        workbook_data = self._findWorkbookData(workbook_name)
        if workbook_data is None:
            return None
        # Список имен листов книги
        worksheet_name_list = [data['Name'] for data in workbook_data['children'] if data['name'] == 'Worksheet']

        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_name in enumerate(worksheet_name_list):
                if worksheet_name not in not_deleted_worksheet_names:
                    self.delWorksheet(workbook_name,worksheet_name)

    def _pasteStyleIntoWorkbook(self, Style_, WorkbookData_):
        """
        Вставить стиль в готовую структуру книги.
        @param Style_: Данные вставляемого стиля.
        @param WorkbookData_: Данные книги.
        @return: Возвращает идентификатор стиля или None
        в случае ошибки.
        """
        try:
            workbook_styles = [data for data in WorkbookData_['children'] if data['name'] == 'Styles'][0]
            if Style_ not in workbook_styles['children']:
                workbook_styles_id = [data['ID'] for data in workbook_styles['children'] if data['name'] == 'Style']

                new_style_id = self._genNewStyleID(Style_['ID'], workbook_styles_id)
                Style_['ID'] = new_style_id
                workbook_styles['children'].append(Style_)
                return new_style_id
            else:
                # Точно такой стиль уже есть и поэтому добавлять его не надо
                return Style_['ID']
        except:
            log.error('_pasteStyleIntoWorkbook function')
            raise
        return None

    def _replaceStyleID(self, Data_, OldStyleID_, NewStyleID_):
        """
        Заменить идентификаторы стилей в вставляемых данных.
        @param Data_: Данные для вставки.
        @param OldStyleID_: Старый идентификатор стиля.
        @param NewStyleID_: Новый идентификатор стиля.
        @return: Возвращает данные с поправленным стилем.
        """
        if OldStyleID_ == NewStyleID_:
            # Идентификаторы равны - замены не требуется
            return Data_
        if 'StyleID' in Data_ and Data_['StyleID'] == OldStyleID_:
            Data_['StyleID'] = NewStyleID_
        if 'children' in Data_ and Data_['children']:
            for i, child in enumerate(Data_['children']):
                Data_['children'][i] = self._replaceStyleID(child, OldStyleID_, NewStyleID_)
        return Data_

    def _genNewWorksheetName(self, WorksheetName_, WorksheetNames_):
        """
        Подобрать имя для листа, чтобы не пересекалось с уже существующими.
        """
        new_sheet_name = WorksheetName_

        i = 1
        while new_sheet_name in WorksheetNames_:
            new_sheet_name = WorksheetName_+'_'+str(i)
            i += 1
        return new_sheet_name

    def pasteWorksheet(self, XMLFileName_=None, IsCut_=None, NewWorksheetName_=None):
        """
        Вставить лист из буфера обмена в указанную книгу.
        @param XMLFileName_: Имя XML файла книги для вставки. Если не определено,
        то имеется ввиду активная книга.
        @param IsCut_: Признак того, что старый лист нужно удалить.
        Если None, то взять системный признак.
        @param NewWorksheetName_: Новое имя листа.
        @return: True-лист вставлен, False-лист по какойто причине не вставлен.
        """
        XMLFileName_ = self._unificXMLFileName(XMLFileName_)
        NewWorksheetName_ = self._unicode2str(NewWorksheetName_)

        # Если в буфере обмена ничего нет, то вставку не производить
        if not self._worksheet_clipboard:
            return False

        # Определить книгу, в которую надо вставить лист
        workbook_data = self._findWorkbookData(XMLFileName_)
        if workbook_data is None:
            return False

        if IsCut_ is None:
            IsCut_ = self._is_cut_worksheet
        # Определить новое имя листа
        sheet_names = [data['Name'] for data in workbook_data['children'] if data['name'] == 'Worksheet']
        worksheet_source, worksheet_data = [(key, value) for key, value in self._worksheet_clipboard.items()
                                            if isinstance(key, tuple)][0]
        if NewWorksheetName_:
            sheet_name = NewWorksheetName_
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
        if IsCut_:
            self.delWorksheet(*worksheet_source)
        # Очистить буфер обмена
        self._worksheet_clipboard = {}
        return True

    def moveWorksheet(self, OldXMLFileName_=None, SheetName_=None, NewXMLFileName_=None):
        """
        Переместить лист из одной книги в другую.
        """
        self.cutWorksheet(OldXMLFileName_, SheetName_)
        return self.pasteWorksheet(NewXMLFileName_)

    def selectWorksheet(self, XMLFileName_=None, SheetName_=None):
        """
        Выбрать лист для групповых операций с листами.
        @param XMLFileName_: Имя XML файла книги. Если не определено,
        то имеется ввиду активная книга.
        @param SheetName_: Имя листа в указанной книге. Если не указано,
        то имеется ввиду первый лист.
        """
        XMLFileName_ = self._unificXMLFileName(XMLFileName_)
        SheetName_ = self._unicode2str(SheetName_)

        self._worksheet_list_clipboard.append((XMLFileName_, SheetName_))

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

    def copyWorksheetListTo(self, XMLFileName_=None, NewWorksheetNames_=None):
        """
        Копирование списка выбранных листов в книгу.
        @param XMLFileName_: Имя XML файла книги для вставки. Если не определено,
        то имеется ввиду активная книга.
        @param NewWorksheetNames_: Список имен новых листов.
        """
        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_src in enumerate(self._worksheet_list_clipboard):
                workbook_name, worksheet_name = worksheet_src
                new_worksheet_name = None
                if NewWorksheetNames_:
                    if i < len(NewWorksheetNames_):
                        new_worksheet_name = NewWorksheetNames_[i]
                self.copyWorksheet(workbook_name, worksheet_name)
                self.pasteWorksheet(XMLFileName_, NewWorksheetName_=new_worksheet_name)

        # После опрации очистить список выбранных листов
        self._worksheet_list_clipboard = []

    def moveWorksheetListTo(self, XMLFileName_=None, NewWorksheetNames_=None):
        """
        Перенести список выбранных листов в книгу.
        @param XMLFileName_: Имя XML файла книги для вставки. Если не определено,
        то имеется ввиду активная книга.
        @param NewWorksheetNames_: Список имен новых листов.
        """
        if type(self._worksheet_list_clipboard) in (list, tuple):
            for i, worksheet_src in enumerate(self._worksheet_list_clipboard):
                workbook_name, worksheet_name = worksheet_src
                new_worksheet_name = None
                if NewWorksheetNames_:
                    if i < len(NewWorksheetNames_):
                        new_worksheet_name = NewWorksheetNames_[i]
                self.cutWorksheet(workbook_name, worksheet_name)
                self.pasteWorksheet(XMLFileName_, NewWorksheetName_=new_worksheet_name)
        # После опрации очистить список выбранных листов
        self._worksheet_list_clipboard = []

    def _unicode2str(self, UnicodeStr_):
        """
        Преобразование unicode строки в обычную строку.
        """
        if isinstance(UnicodeStr_, unicode):
            return UnicodeStr_.encode(self.encoding)
        return UnicodeStr_

    def _unificXMLFileName(self, XMLFileName_):
        """
        Привести к внутреннему виду имя XML файла книги.
        Если имя не определено то берется имя файла активной книги.
        В качестве внутреннего имени XML файла книги
        берется АБСОЛЮТНЫЙ ПУТЬ до файла.
        """
        if XMLFileName_:
            XMLFileName_ = os.path.abspath(XMLFileName_)
        else:
            XMLFileName_ = self.SpreadsheetFileName
        return XMLFileName_

    def mergeCell(self, SheetName_, Row_, Col_, MergeDown_, MergeAcross_, XMLFileName_=None):
        """
        Объединить ячейки.
        """
        if XMLFileName_ is not None:
            self.activeWorkbook(XMLFileName_)
        work_book = self.getActiveWorkbook()
        work_sheet = work_book.findWorksheet(SheetName_)
        table = work_sheet.getTable()
        cell = table.getCell(Row_, Col_)
        return cell.setMerge(MergeAcross_, MergeDown_)

    def setCellValue(self, SheetName_, Row_, Col_, Value_, XMLFileName_=None):
        """
        Установить значение в ячейку.
        """
        if XMLFileName_ is not None:
            self.activeWorkbook(XMLFileName_)
        work_book = self.getActiveWorkbook()
        work_sheet = work_book.findWorksheet(SheetName_)
        table = work_sheet.getTable()
        cell = table.getCell(Row_, Col_)
        return cell.setValue(Value_)

    def setCellStyle(self, sheet_name, row, col, alignment=None,
                     left_border=None, right_border=None, top_border=None, bottom_border=None,
                     font=None, interior=None, number_format=None, XMLFileName_=None):
        """
        Установить стиль ячейки.
        """
        if XMLFileName_ is not None:
            self.activeWorkbook(XMLFileName_)
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

    def exec_cmd_script(self, CmdScript_=None, AutoSave_=True):
        """
        Выполнить скрипт - список команд.
        @param CmdScript_: Список команд формата
        [
        ('ИмяКоманды',(кортеж не именованных аргуметов),{словарь именованных аргументов}),
        ...
        ]
        @param AutoSave_: Автоматически по завершению работы сохранить.
        """
        if CmdScript_:
            for cmd in CmdScript_:
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
                    log.error('Execute command <%s>' % cmd)
                    raise
        if AutoSave_:
            self.Save()

if __name__ == '__main__':
    import tests
    tests.test_work_sheet()
