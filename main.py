import pandas as pd
import re
import xlwt

"""'Бренд из НС': ['None' for i in range(9999)],
                           'Объем/вес из НС': ['None' for i in range(9999)],
                           'Наименование полное (НП)': ['None' for i in range(9999)],
                           'Наименование товара из НП': ['None' for i in range(9999)],
                           'Бренд из НП': ['None' for i in range(9999)],
                           'Объем/вес из НП': ['None' for i in range(9999)]"""
class Convert:
    def __init__(self):
        self.convertible_file = 'data/list.xls'
        self.template_file = 'data/result_conv2.xls'

    def init_converible_file(self):
        file = pd.read_excel(self.convertible_file, sheet_name='Лист1 (2)')
        return file.iloc[:, 0]

    def init_templates_file(self, columns):
        file = pd.read_excel(self.template_file, sheet_name='result_conv2')
        return file.iloc[:, (columns-1):columns]

    def unnecessary(self):
        global df
        file_convert = self.init_converible_file()
        file_tamplates = self.init_templates_file(3)
        mass = []
        for i in file_convert:
            mass.append(i)

        pattern = r'([*#]\d{2,3})|(\d{1,3}[ ][Кк][Гг])|(\d{1,4}[ш][т])|(\d{1,3}[Мм][Лл])|((\d{1,3}[Х])\d{1,3}[ГгКк]$)|(\d{1,3}\w+[^%]\:\d)|(\d\,\d[Гг])|(\d{1,3}[Кк][Гг])|(\s{2}\d$)|([(]\d{1,2}[+]\d{1,2}[)])|([#*])|(\d{1,4}$)|(\d{1,3}[,.]\d{1,3}[LlЛл]|(\d{1,3}[LlЛл]))|(\d{1,2}[+])|(\d{1,3}[*]\d{1,3}[*]\d{1,3}[Мм]{1,2}[,])|(\d{1,4}([ ])[МмШшCc][ЛлТт])|([Кк][Гг]$)|([:]\d{1,3}\/\d{1,3})|([:]\d{1,3}[.,]\d{1,4})|(\d{1,4}[-:]\d{1,4}([Кк][Гг]))|(\d{1,3}[Гг])|(\d{2,3}\s([Гг]))|(\d{1}[,]\d{1}[^%]([Лл]))|(\d{1,3}[Гг][Рр])|(\d{1,3}[,.]\d{1,3}$)|(\d{1,4}[Гг]$)|(\d{1,3}[.]$)|(\d{1,3}[Х]\d{1,3}[Гг][Рр])|([(]$)|(\d{2,3})|(\d{1,4} [МмШшCc][ЛлТт][,.])|(\d{1,3}[Кк])'
        repl = r''
        mass3 = []
        for i in file_convert:
            mass3.append(re.sub(pattern, repl, str(i)).strip())

        mass2 = []
        for i in mass3:
            for j in range(len(file_tamplates)):
                if str(file_tamplates['Наименование товара из НС'][j]).lower() == str(i).lower():
                    tmp = file_tamplates['Наименование товара из НС'][j]
                    print(tmp)
                    mass2.append(tmp)
                    break
                else:
                    mass2.append('None')
        mydict = {'Наименование сокращенное (НС)': mass,
                  'Наименование товара из НС': mass2}
        df = pd.DataFrame({key: pd.Series(value) for key, value in mydict.items() })
        print(len(mass2))
        return df

    def create_new_xls(self):
        open('test.xls', 'w')
        df = self.unnecessary()
        df.to_excel('data/test.xls', sheet_name='Лист 1', index=False)
        return print('OK')


if __name__ == "__main__":
    a = Convert()
    a.create_new_xls()
    #a.unnecessary()
    #print(a.init_templates_file(3))
