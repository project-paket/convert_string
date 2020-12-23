import pandas as pd
import re
import xlwt
import time

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

    def unnecessary_name(self):
        file_convert = self.init_converible_file()
        file_tamplates = self.init_templates_file(3)


    def unnecessary(self):
        file_convert = self.init_converible_file()
        file_tamplates = self.init_templates_file(3)
        mass = []
        for i in file_convert:
            mass.append(i)

        pattern = r'([*#]\d{2,3})|(\d{1,3}[ ][Кк][Гг])|' \
                  r'(\d{1,4}[ш][т])|(\d{1,3}[Мм][Лл])|((\d{1,3}[Х])\d{1,3}[ГгКк]$)|(\d{1,3}\w+[^%]\:\d)|' \
                  r'(\d\,\d[Гг])|(\d{1,3}[Кк][Гг])|(\s{2}\d$)|([(]\d{1,2}[+]\d{1,2}[)])|([#*])|(\d{1,4}$)|' \
                  r'(\d{1,3}[,.]\d{1,3}[LlЛл]|(\d{1,3}[LlЛл]))|(\d{1,2}[+])|(\d{1,3}[*]\d{1,3}[*]\d{1,3}[Мм]{1,2}[,])' \
                  r'|(\d{1,4}([ ])[МмШшCc][ЛлТт])|([Кк][Гг]$)|([:]\d{1,3}\/\d{1,3})|([:]\d{1,3}[.,]\d{1,4})|(\d{1,4}[-:]\d{1,4}([Кк][Гг]))|(\d{1,3}[Гг])|' \
                  r'(\d{2,3}\s([Гг]))|(\d{1}[,]\d{1}[^%]([Лл]))|(\d{1,3}[Гг][Рр])|(\d{1,3}[,.]\d{1,3}$)|(\d{1,4}[Гг]$)|(\d{1,3}[.]$)' \
                  r'|(\d{1,3}[Х]\d{1,3}[Гг][Рр])|([(]$)|(\d{2,3})|(\d{1,4} [МмШшCc][ЛлТт][,.])|(\d{1,3}[Кк])'
        repl = r''
        mass3 = []
        for i in file_convert:
            mass3.append(re.sub(pattern, repl, str(i)).strip())

        mass3tmp = [str(i).lower() for i in mass3]
        template_tmp = [str(i).lower() for i in file_tamplates['Наименование товара из НС']]
        mass2 = []
        c = 0
        for i in range(len(mass3tmp)):
            if mass3tmp[i] in template_tmp:
                mass2.append(mass3[i])
                c += 1
            else:
                mass2.append('')
        for i in range(len(mass2)):
            mass2[i] = str(mass2[i]).lower()
        open('names.txt', 'w').write('\n'.join(set(mass2)).lower())
        open('names_all.txt', 'w').write('\n'.join(set(template_tmp)))

        mydict = {'Наименование сокращенное (НС)': mass,
                  'Наименование товара из НС': mass2}
        df = pd.DataFrame({key: pd.Series(value) for key, value in mydict.items()})
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
