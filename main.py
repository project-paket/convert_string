import pandas as pd
import re


class Convert:
    # Инициализация пути до файлов,convertible_file это файл с стоками на обработку
    # template_file файл обработанный в ручную человеком
    def __init__(self):
        self.convertible_file = 'data/list.xls'
        self.template_file = 'data/result_conv2.xls'

    # Чтение файла convertible_file и получение с помощью iloc столбца содержащих строки
    def init_converible_file(self):
        file = pd.read_excel(self.convertible_file, sheet_name='Лист1 (2)')
        return file.iloc[:, 0]

    # Чтение файла template_file, передаем columns и на выходе получаем конкретный столбец с данными
    def init_templates_file(self, columns):
        file = pd.read_excel(self.template_file, sheet_name='result_conv2')
        return file.iloc[:, (columns-1):columns]

    # Выборка данных с весом, на выходе получаем 2 массива с данными
    # name_product_mass это массив с данными (без веса), save_weight массив с данными где хранится только вес
    def selection_weight(self):
        file_convert = self.init_converible_file()  # Получение строк для обработки
        patterns = open('data/re_search_weight.txt', 'r').read()  # Инициализация шаблонов с регулярными выражениями

        # Создал 2 массива, name_product_mass это массив где будут хранится данные (без веса),
        # save_weight массив с данными где будет хранится только вес
        name_product_mass = []
        save_weight = []

        '''
            Первый for прохожусь по полям обрабатываемого столбца, flag нужен для того, чтобы если
            не нашли совпадений то поле останется неизменным и добавится в массив name_product_mass, а 
            в save_weight добавляем пустую строку. Второй for прохожусь по шаблонам, где если существует
            поле, шаблон и искомое значение (например 1кг), то в name_product_mass добавляется строка без единиц веса
            (т.е была строка Мандарины 1кг, станет Мандарины), а в save_weight добавляется объект 
            match c данными о весе (1кг). И после этого переходим к следующей строке
        '''
        for field in file_convert:

            flag = False
            for pattern in patterns.split('\n'):

                if field and pattern and re.search(pattern, str(field)):
                    name_product_mass.append(re.sub(pattern, '', str(field)).strip())
                    save_weight.append(re.search(pattern, str(field)))
                    flag = True
                    break

            if not flag:
                name_product_mass.append(str(field))
                save_weight.append('')

        # По циклу получаю данные из объекта match и добавляю его в save_weight
        for field in range(len(save_weight)):

            if save_weight[field]:
                save_weight[field] = save_weight[field].group(0)

        return name_product_mass, save_weight

    # Выборка данных с брендом, на выходе получаем 2 массива с данными
    # name_product_mass это массив с данными (без бренда), save_data_brand массив с данными где хранится только бренд
    def selection_brand(self):
        mass_info_product, info_weight = self.selection_weight()  # Инициализация предыдущей выборки, без веса
        mass_info_product = [str(i).lower() for i in mass_info_product]  # Привел все значения к нижнему регистру
        patterns = open('data/re_search_brand.txt', 'r').read()  # Инициализация шаблонов с регулярными выражениями

        # Создал 2 массива, name_product_mass это массив где будут хранится данные (без веса и без бренда),
        # save_data_brand массив с данными где будет хранится только бренд
        name_product_mass = []
        save_data_brand = []

        '''
            Первый for прохожусь по полям обрабатываемого столбца, flag нужен для того, чтобы если
            не нашли совпадений то поле останется неизменным и добавится в массив name_product_mass, а 
            в save_data_brand добавляем пустую строку. Второй for прохожусь по шаблонам, где если существует
            поле, шаблон и искомое значение (например Danone), то в name_product_mass добавляется строка без бренда
            (т.е была строка йогурт Danone, станет йогурт), а в save_data_brand добавляется объект 
            match c данными о весе (Danone). И после этого переходим к следующей строке
        '''
        for field in mass_info_product:

            flag = False
            for pattern in patterns.split('\n'):

                if field and pattern and re.search(pattern, str(field)):
                    name_product_mass.append(re.sub(pattern, '', str(field)).strip())
                    save_data_brand.append(re.search(pattern, str(field)))
                    flag = True
                    break

            if not flag:
                name_product_mass.append(str(field))
                save_data_brand.append('')

        # По циклу получаю данные из объекта match и добавляю его в save_data_brand
        for field in range(len(save_data_brand)):

            if save_data_brand[field]:
                save_data_brand[field] = save_data_brand[field].group(0)

        return name_product_mass, save_data_brand

    # Удаление лишних символов по типу ("" , . , :) и т.п. на выходе получаем готовый массив с наименованием товара
    # в несокращенном виде
    def delete_selection_symbol(self):
        mass_product, mass_brand = self.selection_brand()  # Инициализация предыдущей выборки, без веса и бренда
        mass_patterns = open('data/re_delete_symbol.txt').read()  # Инициализация шаблонов с регулярными выражениями

        # Создал массив result_product где будут хранится данные (без веса и без бренда и лишних символов)
        result_product = []

        '''
            Первый for прохожусь по полям обрабатываемого столбца. Переменная string нужна для обработки всех шаблонов,
            для одной строки. Второй for прохожусь по шаблонам, где сначала убираю (125 строка) лишние пробелым
            между словами, а затем убираю лишние символы и добавляю в массив result_product.
        '''
        for field in mass_product:

            string = field
            for pattern in mass_patterns.split('\n'):
                string = " ".join(re.split("\s+", string))
                string = re.sub(pattern, '', str(string)).strip()
            result_product.append(string)

        return result_product

    # Метод для преобразования сокращенной формы веса к полной (пример было 1кг, стало 1 килограмм)
    # На выходе получаем массив mass_full_weight где хранятся данные о весе без сокращенной формы
    def to_full_weight(self):
        mass_name_prod, mass_weight = self.selection_weight()  # Инициализация предыдущей выборки, без веса
        patterns = open('data/re_full_weight.txt').read()  # Инициализация шаблонов с регулярными выражениями

        # Создал массив mass_full_weight где будут хранится данные о весе без сокращенной формы
        mass_full_weight = []

        '''
            Первый for прохожусь по полям обрабатываемого столбца. Переменная string нужна для обработки всех шаблонов,
            для одной строки. Второй for прохожусь по шаблонам, где проверяю если есть такой шаблон, то разделяю
            шабло на 2 значения регулярное выражение и на что заменить пример (регулярное выражение||| килограмм)
            Проверяю есть ли искомое значение с помощью регулярных выражение, если есть то изменяю значения из 
            сокращенной формы в полную форму (пример было 1кг, стало 1 килограмм) и добавляю в массив, когда обработал
            все шаблоны для данной строки. Если искомого значения нет, то строку никак не видоизменяю.
        '''
        for field in mass_weight:

            string = field
            for pattern in patterns.split('\n'):

                if pattern:
                    pattern = pattern.split('|||')

                    if re.search(pattern[0], string):
                        string = re.sub(pattern[0], r'\1' + pattern[1], string).strip()
                        break

            mass_full_weight.append(string)

        return mass_full_weight

    # Метод для преобразования сокращенной формы бренда к полной (пример было dano. , стало Danone)
    # На выходе получаем массив mass_full_brand где хранятся данные о брендах без сокращенной формы
    def to_full_brand(self):
        mass_name_prod, mass_brand = self.selection_brand()  # Инициализация предыдущей выборки, без брендов
        patterns = open('data/re_full_brand.txt').read()  # Инициализация шаблонов с регулярными выражениями

        # Создал массив mass_full_brand где будут хранится данные о бренде без сокращенной формы
        mass_full_brand = []

        '''
            Первый for прохожусь по полям обрабатываемого столбца. Переменная string нужна для обработки всех шаблонов,
            для одной строки. Второй for прохожусь по шаблонам, где проверяю если есть такой шаблон, то разделяю
            шабло на 2 значения регулярное выражение и на что заменить пример (регулярное выражение||| Danone)
            Проверяю есть ли искомое значение с помощью регулярных выражение, если есть то изменяю значения из 
            сокращенной формы в полную форму (пример было dano. , стало Danone) и добавляю в массив, когда обработал
            все шаблоны для данной строки. Если искомого значения нет, то строку никак не видоизменяю.
        '''
        for field in mass_brand:

            string = field
            for pattern in patterns.split('\n'):

                if pattern:
                    pattern = pattern.split('|||')

                    if re.search(pattern[0], string):
                        string = re.sub(pattern[0], pattern[1], string).strip()
                        break

            mass_full_brand.append(string)

        return mass_full_brand

    # Иницилизация исходного столбца с данными для обработки в массив mass_name_ns
    def create_list_xls(self):
        file_convert = self.init_converible_file()  # Получение строк для обработки

        # Массив для хранения исходных данных
        mass_name_ns = []

        # Прохожусь по полям столбца и добавляю их в массив
        for field in file_convert:
            mass_name_ns.append(field)

        return mass_name_ns

    # Метод для иницилазации обработанных данных в новую таблицу.
    def create_dataframe(self):
        name_product_mass, save_data_weight = self.selection_weight()  # Информация о весе в сокращенной форме
        data_mass_product, save_data_brand = self.selection_brand()  # Информация о брендах в сокращенной форме

        # Иницилазация столбцов и обработанных данных пример (Наименование сокращенное (НС) это заголовок столбца,
        # а self.create_list_xls() это исходнные данные для обработки)
        mydict = {'Наименование сокращенное (НС)': self.create_list_xls(),
                  'Наименование товара (НС)': self.delete_selection_symbol(),
                  'бренд (НС)': save_data_brand,
                  'Вес (НС)': save_data_weight,
                  'Бренд без сокращений': self.to_full_brand(),
                  'Вес без сокращений': self.to_full_weight()}

        df = pd.DataFrame({key: pd.Series(value) for key, value in mydict.items()})

        return df

    # Создание новой таблицы и внесение данных в неё
    def create_new_xls(self):
        open('result_table.xls', 'w')  # Содания нового файла формата xls

        #  Инициализация столбцов и запись в таблицу result_tables.xls информации
        df = self.create_dataframe()
        df.to_excel('data/result_table.xls', sheet_name='Лист 1', index=False)


if __name__ == "__main__":
    convert = Convert()
    convert.create_new_xls()