import pandas as pd
import re


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

    def unnecessary_weight(self):
        file_convert = self.init_converible_file()
        pattern = open('data/re_search_weight.txt', 'r').read()
        repl = r''
        tmp_mass = []
        save_info = []
        for field in file_convert:
            flag = False
            for tmp in pattern.split('\n'):
                if field and tmp and re.search(tmp, str(field)):
                    tmp_mass.append(re.sub(tmp, repl, str(field)).strip())
                    save_info.append(re.search(tmp, str(field)))
                    flag = True
                    break
            if not flag:
                tmp_mass.append(str(field))
                save_info.append('')

        for match in range(len(save_info)):
            if save_info[match]:
                save_info[match] = save_info[match].group(0)
        return tmp_mass, save_info

    def unnecessary_brand(self):
        file_del_weight, mass = self.unnecessary_weight()
        file_del_weight = [str(i).lower() for i in file_del_weight]
        pattern = open('data/re_search_brand.txt', 'r').read()
        repl = r''
        tmp_mass = []
        save_info = []
        for field in file_del_weight:
            flag = False
            for tmp in pattern.split('\n'):
                if field and tmp and re.search(tmp, str(field)):
                    tmp_mass.append(re.sub(tmp, ' ' + repl, str(field)).strip())
                    save_info.append(re.search(tmp, str(field)))
                    flag = True
                    break
            if not flag:
                tmp_mass.append(str(field))
                save_info.append('')
        for match in range(len(save_info)):
            if save_info[match]:
                save_info[match] = save_info[match].group(0)
        return tmp_mass, save_info

    def delete_unnecessary_symbol(self):
        mass_product, mass_brand = self.unnecessary_brand()
        mass_pattern = open('data/re_delete_unsymbol.txt').read()
        new_mass_product = []
        for field in mass_product:
            string = field
            for pattern in mass_pattern.split('\n'):
                string = re.sub(pattern, '', str(string)).strip()
            new_mass_product.append(string)
        return new_mass_product

    def to_full_weight(self):
        mass_name_prod, mass_weight = self.unnecessary_weight()
        patterns = open('data/re_full_weight.txt').read()
        mass_full_weight = []
        for field in mass_weight:
            string = field
            for pattern in patterns.split('\n'):
                if pattern:
                    pattern = pattern.split('|||')
                    if re.search(pattern[0], string):
                        string = re.sub(pattern[0], r'\1' + pattern[1], string)
                        break
            mass_full_weight.append(string)
        return mass_full_weight

    def to_full_brand(self):
        mass_name_prod, mass_brand = self.unnecessary_brand()
        patterns = open('data/re_full_brand.txt').read()
        mass_full_brand = []
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

    def create_list_xls(self):
        file_convert = self.init_converible_file()
        mass = []
        for i in file_convert:
            mass.append(i)
        return mass

    def create_dataframe(self):
        name_ns, name_weight = self.unnecessary_weight()
        brand_ns, name_brand = self.unnecessary_brand()
        name_product = self.delete_unnecessary_symbol()
        mass_full_weight = self.to_full_weight()
        mass_full_brand = self.to_full_brand()
        mydict = {'Наименование сокращенное (НС)': self.create_list_xls(),
                  'Наименование товара': name_product,
                  'бренд из НС': name_brand,
                  'Вес': name_weight,
                  'Бренд без сокращений': mass_full_brand,
                  'Вес без сокращений': mass_full_weight}
        df = pd.DataFrame({key: pd.Series(value) for key, value in mydict.items()})
        return df

    def create_new_xls(self):
        open('result_table.xls', 'w')
        df = self.create_dataframe()
        df.to_excel('data/result_table.xls', sheet_name='Лист 1', index=False)
        return print('OK')


if __name__ == "__main__":
    a = Convert()
    a.create_new_xls()