#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Модуль конвертора файлов Excel в xml формате в словарь.

ВНИМАНИЕ! Модули в этом пакете используются движком Virtual Excel.
Менять на другие/другой версии их нельзя. Можно только править в рамках
проекта icReport.
"""

# Подключение библиотек
import sys

from xml.sax import xmlreader
import xml.sax.handler

__version__ = (1, 1, 1, 2)


# Описания функций
def XmlFile2Dict(XMLFileName_, encoding='utf-8'):
    """
    Функция конвертации файлов Excel в xml формате в словарь Python.
    @param XMLFileName_: Имя xml файла.
    @param encoding: Кодировка XML файла.
    @return: Функция возвращает заполненный словарь, 
        или None в случае ошибки.
    """
    xml_file = None
    try:
        xml_file = open(XMLFileName_, 'rt', encoding=encoding)

        input_source = xmlreader.InputSource()
        input_source.setByteStream(xml_file)

        xml_reader = xml.sax.make_parser()
        xml_parser = icXML2DICTReader(encoding=encoding)
        xml_reader.setContentHandler(xml_parser)

        # включаем опцию namespaces
        xml_reader.setFeature(xml.sax.handler.feature_namespaces, 1)
        xml_reader.parse(input_source)
        xml_file.close()

        return xml_parser.getData()
    except:
        if xml_file:
            xml_file.close()
        info = str(sys.exc_info()[1])
        print('Error read file <%s> : %s.' % (XMLFileName_, info))
        return None


# Описания классов
class icXML2DICTReader(xml.sax.handler.ContentHandler):
    """
    Класс анализатора файлов Excel-xml формата.
    """
    def __init__(self, encoding='utf-8', *args, **kws):
        """
        Конструктор.
        @param encoding: Кодировка XML файла.
        """
        xml.sax.handler.ContentHandler.__init__(self, *args, **kws)

        # Выходной словарь
        self._data = {'name': 'Excel', 'children': []}
        # Текущий заполняемый узел
        self._cur_path = [self._data]

        # Текущее анализируемое значение
        self._cur_value = None

        # Кодировка
        self.encoding = encoding

    def getData(self):
        """
        Выходной словарь.
        """
        return self._data

    def _eval_value(self, value):
        """
        Попытка приведения типов данных.
        """
        try:
            # Попытка приведения типа
            return eval(value)
        except:
            # Скорее всего строка
            return value
        
    def characters(self, content):
        """
        Данные.
        """
        try:
            if content.strip():
                if self._cur_value is None:
                    self._cur_value = ''
                self._cur_value += content
        except:
            print('ERROR characters')
            raise

    def startElementNS(self, name, qname, attrs):
        """
        Разбор начала тега.
        """
        try:
            # Имя элемента задается кортежем
            if type(name) is tuple:

                # Имя элемента
                element_name = name[1]

                # Создать структуру,  соответствующую элементу
                self._cur_path[-1]['children'].append({'name': element_name, 'children': []})
                self._cur_path.append(self._cur_path[-1]['children'][-1])
                cur_node = self._cur_path[-1]

                # Имена параметров
                element_qnames = attrs.getQNames()
                if element_qnames:
                    # Разбор параметров элемента
                    for cur_qname in element_qnames:
                        # Имя параметра
                        element_qname = attrs.getNameByQName(cur_qname)[1]
                        # Значение параметра
                        element_value = attrs.getValueByQName(cur_qname)
                        cur_node[element_qname] = element_value
        except:
            print('ERROR startElementNS::', name, qname, attrs)
            raise

    def endElementNS(self, name, qname): 
        """
        Разбор закрывающего тега.
        """
        try:
            # Сохранить проанализированное значение
            if self._cur_value is not None:
                self._cur_path[-1]['value'] = self._cur_value
                self._cur_value = None
            
            del self._cur_path[-1]
        except:
            print('ERROR endElementNS::', name, qname)
            raise


def default_test():
    """
    Тестовая функция.
    """
    rep_file = None
    xml_file = None
    xml_file = open(sys.argv[1], 'rt', encoding='utf-8')

    input_source = xmlreader.InputSource()
    input_source.setByteStream(xml_file)
    xml_reader = xml.sax.make_parser()
    xml_parser = icXML2DICTReader()
    xml_reader.setContentHandler(xml_parser)
    # включаем опцию namespaces
    xml_reader.setFeature(xml.sax.handler.feature_namespaces, 1)
    xml_reader.parse(input_source)
    print(xml_parser.getData())
    xml_file.close()


def create_pkl_files_test():
    """
    Создание выходной структуры в файле pickle.
    """
    import pickle
    import time

    start_time = time.time()
    print('START Pickle file create test')
    
    data = XmlFile2Dict('./testfiles/SF02.xml')
    print('READ ... ok Time(s):', time.time()-start_time)
    
    start_time = time.time()
    f_out = open('./testfiles/SF02.txt', 'wt')
    f_out.write(str(data))
    f_out.close()
    print('WRITE text ... ok Time(s):', time.time()-start_time)
    
    start_time = time.time()
    f_out = open('./testfiles/SF02.pkl', 'wb')
    pkl = pickle.Pickler(f_out)
    pkl.dump(data)
    f_out.close()
    print('WRITE pickle ... ok Time(s):', time.time()-start_time)
    
    start_time = time.time()
    f_out = open('./testfiles/SF02.cpk', 'wb')
    pkl = pickle.dump(data, f_out)
    f_out.close()
    print('WRITE cPickle ... ok Time(s):', time.time()-start_time)


if __name__ == '__main__':
    # default_test()
    create_pkl_files_test()
