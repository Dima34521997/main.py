import pandas as pd
from PIL import ImageFont

from DataModel import DataModel as dm

font = ImageFont.truetype('T-FLEXA.TTF', 12)


class Element():
    def __init__(self, ser: pd.Series = None):
        if ser is not None:
            self.ser = ser
            self.desig = ser['Designator']
            self.char = self.desig[0]
            self.rem = ser['Rem']
            self.body = ser['Корпус']
            self.tke = ser['TKE']
            self.val = ser['Value']
            self.tol = ser['Tolerance']
            self.pv = ser['Power/Voltage']
            self.man = ser['Manufacturer'].split(',')[0]
            self.manpnb = ser['ManufacturerPartNumber']
            self.quantity = ser['Quantity']
            self.comment = ser['Примечание']
            self.module = ""

            self.name = self.make_name()
            self.splt_desig = []

            if '0' not in self.body and '1' not in self.body and self.body != '':
                self.body = "«" + self.body + "»"

            if '"' in self.tke:
                output = list(self.tke)
                output[-1] = "»"
                output[-3] = "«"
                self.tke = ''.join(output)

    def __str__(self):
        return f"{self.name} [{self.quantity}] [{self.module}]"

    def __repr__(self):
        return f"{self.name} [{self.quantity}] [{self.module}]"


    def make_name_for_C(self):
        """
        # Для R и C Номинал и Погрешность переносятся вместе, поэтому сохраняем их в одну строку
        :return:
        """
        val_tol = self.val + " " + self.tol
        name = ''
        if 'ТУ' in self.man:
            if self.tke:
                name = f'{self.rem} {self.pv} - {val_tol} - {self.tke} {self.man}'
            else:
                name = f'{self.rem} {self.pv} - {val_tol} {self.man}'
        else:
            if self.tke:
                name = f'{self.rem} {self.body} {self.tke}  - {self.pv} - {val_tol}, {self.man}'
            else:
                name = f'{self.rem} {self.body} - {self.pv} - {val_tol}, {self.man}'
        return name


    def make_name_for_R(self):
        """
        Для R и C Номинал и Погрешность переносятся вместе, поэтому сохраняем их в одну строку
        :return:
        """
        val_tol = self.val + " " + self.tol
        name = ''
        if 'ТУ' in self.man:
            if self.tke:
                name = f'{self.rem} - {self.pv} - {val_tol} - {self.tke} {self.man}'
            else:
                name = f'{self.rem} - {self.pv} - {val_tol} {self.man}'
        else:
            name = f'{self.rem} {self.body} - {val_tol}, {self.man}'
        return name

    def make_name_for_other_elements(self):
        if 'ТУ' in self.man:
            name = f'{self.manpnb} {self.man}'
        else:
            name = f'{self.manpnb}, {self.man}'
        if self.manpnb == '':
            if 'ТУ' in self.man:
                name = f'{self.val} {self.man}'
            else:
                name = f'{self.val}, {self.man}'
        return name


    def make_name(self):
        """
        Формирование наименований
        :return:
        """
        if self.char == 'C':
            return self.make_name_for_C()
        elif self.char == 'R':
            return self.make_name_for_R()
        else:
            return self.make_name_for_other_elements()


    def split_desig(self, shift_treshold: int):
        """
        Переносит позиции элементов, если те не влезают в строку шаблона

        :param shift_treshold: Порог количества символов для переноса
        :return: Лист позиций. Каждый новый элемент на новой строке
        """
        shift_treshold = 53
        splitted_desig = self.desig.split(', ')
        #splitted_desig = splitted_desig1.split(", ")

        for i in range(len(splitted_desig)):
            next_i = i + 1
            # Проверка существования следующей позиции
            if next_i >= len(splitted_desig):
                break
            # Пока длина позиций не привышает порог, добавляем следующую позицию
            # -2 для вставки запятой в конце строки
            while font.getlength(splitted_desig[i] + ", " + splitted_desig[next_i]) <= shift_treshold - 2:
                splitted_desig[i] += ", " + splitted_desig[next_i]
                # Удаляем добавленную позицию из списка всех позиций
                splitted_desig.remove(splitted_desig[next_i])
                # Проверка существования следующей позиции

                if next_i >= len(splitted_desig):
                    break

            if next_i < len(splitted_desig):
                splitted_desig[i] += ","

        self.splt_desig = splitted_desig

    def split_name(self, shift_threshold, cat_name='', one_man=''):
        val_tol = self.val + " " + self.tol
        splitted_name = []

        if self.man.find("ТУ", 0) != -1 and self.manpnb != "":
            new_split = self.manpnb.split()

            i = 0
            while i + 1 < len(new_split):
                new_split[i] = new_split[i].replace('“', "«")
                new_split[i] = new_split[i].replace('”', "»")
                while new_split[i + 1] == '-' or new_split[i + 1] == 'В' or \
                        ("Ф" in new_split[i + 1] or "Гн" in new_split[i + 1] or "Ом" in new_split[i + 1]) or \
                        "%" in new_split[i + 1]:
                    new_split[i] = f"{new_split[i]} {new_split[i + 1]}"
                    del (new_split[i + 1])
                    if i + 1 == len(new_split):
                        break
                else:
                    i += 1
            splitted_name = new_split + [f"{self.man}"]
        elif self.char == 'C':
            if self.man.find("ТУ", 0) != -1 or (self.name is list and self.name.find("ТУ", 0) != -1):
                if self.tke != '':
                    splitted_name = [f"{self.rem}", f"{self.pv} -", f"{val_tol} -", f"{self.tke}",
                                     f"{self.man}"]
                else:
                    splitted_name = [f"{self.rem}", f"{self.pv} -", f"{val_tol}", f"{self.man}"]
            else:
                if self.tke != '':
                    splitted_name = [f"{self.rem}", f"{self.body}", f"{self.tke} -", f"{self.pv} -", f"{val_tol},",
                                     f"{self.man}"]
                else:
                    splitted_name = [f"{self.rem}", f"{self.body} -", f"{self.pv} -", f"{val_tol},", f"{self.man}"]
        elif self.char == 'R':
            if self.man.find("ТУ", 0) != -1 or (self.name is list and self.name.find("ТУ", 0) != -1):
                if self.tke != '':
                    splitted_name = [f"{self.rem} -", f"{self.pv} -", f"{val_tol} -", f"{self.tke}", f"{self.man}"]
                else:
                    splitted_name = [f"{self.rem} -", f"{self.pv} -", f"{val_tol} -", f"{self.man}"]
            else:
                splitted_name = [f"{self.rem}", f"{self.body} -", f"{val_tol},", f"{self.man}"]
        # Для прочих компонентов
        elif self.char != 'C' and self.char != 'R':
            if self.man.find("ТУ", 0) != -1:
                splitted_name = [f"{self.manpnb}", f"{self.man}"]
            else:
                splitted_name = [f"{self.manpnb},", f"{self.man}"]
            if self.manpnb == '':
                if self.man.find("ТУ", 0) != -1:
                    splitted_name = [f"{self.val}", f"{self.man}"]
                else:
                    splitted_name = [f"{self.val},", f"{self.man}"]

        if cat_name != '':
            splitted_name.insert(0, cat_name)

        if one_man != '':
            splitted_name.insert(-1, cat_name)

        i = 0 # Индекс для добавления части наименования в уже существующую строку
        name = ['']
        for name_part in splitted_name:
            # Если длина итоговой строки длиннее порога, то создаем новый перенос
            if len(f"{name[-1]} {name_part}") > shift_threshold:
                if name[-1][-1] == '-':
                    name_part = f"- {name_part}"
                name.append(name_part)
                i += 1
                continue
            name[i] += f" {name_part}"

        for index, name_part in enumerate(name):
            name[index] = name[index].lstrip().rstrip()
            if self.man == '':
                name[index] = name[index][:-1]

        self.name = name

    def split_man(self, shift_threshold):
        if self.man != '':
            i = 0
            new_man = ['']
            for man_part in self.man.split():
                # Если длина итоговой строки длиннее порога, то создаем новый перенос
                if len(f"{new_man[-1]} {man_part}") > shift_threshold:
                    if new_man[-1][-1] == '-':
                        man_part = f"- {man_part}"
                    new_man.append(man_part)
                    i += 1
                    continue
                new_man[i] += f" {man_part}"
            self.man = new_man

def create_name(df: pd.DataFrame, index: int, shift_treshold: int):
    """
    Формирует наименование компонента и переносит его, если тот не будет влезат в рамки шаблона

    :param df: Словарь с сортированными по обозначениям компонентами
    :param index: Индекс редактирумеого элемента
    :param shift_treshold: Порог длины строки, привышая который будет производиться перенос наимнования
    :return: Готовое наименование элемента, [name2, name3] - переносы
    """
    desig = df['Designator'][index]
    rem = df['Rem'][index]
    body = df['Корпус'][index]
    tke = df['TKE'][index]
    val = df['Value'][index]
    pv = df['Power/Voltage'][index]
    tol = df['Tolerance'][index]
    man = df['Manufacturer'][index]
    manpnb = df['ManufacturerPartNumber'][index]
    # Для R и C Номинал и Погрешность переносятся вместе, поэтому сохраняю их в одну строку
    val_tol = val + " " + tol

    name = ''
    name2 = ''
    name3 = ''

    if '0' not in body and '1' not in body and body != '':
        body = "«" + body + "»"

    if '"' in tke:
        output = list(tke)
        output[-1] = "»"
        output[-3] = "«"
        tke = ''.join(output)

    # Формирование наименований
    if desig.find("C", 0) != -1:
        if man.find("ТУ", 0) != -1:
            if tke != '':
                # Костыль для некоторых корпусов
                if '0' not in body and 'М' not in body:
                    body = ''
                if len(f"{rem} {body} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol} - {tke} {man}"
                    if len(name2) > shift_treshold:
                        name2 = f"- {val_tol} - {tke}"
                        name3 = f"{man}"
                elif len(f"{rem} {body} - {pv} - {val_tol} - {tke}") > shift_treshold:
                    name = f"{rem} {body} - {pv} - {val_tol} -"
                    name2 = f"- {tke} {man}"
                elif len(f"{rem} {body} {pv} - {val_tol} - {tke} {man}") > shift_treshold:
                    name = f"{rem} {body} {pv} - {val_tol} - {tke}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol} - {tke} {man}"
            else:
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} {man}"
                elif len(f"{rem} - {pv} - {val_tol} {man}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol} {man}"
        else:
            if tke != '':
                if len(f"{rem} {body} {tke} - {pv}") > shift_treshold:
                    name = f"{rem} {body} {tke} -"
                    name2 = f"- {pv} - {val_tol}, {man}"
                elif len(f"{rem} {body} {tke} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} {tke}  - {pv} -"
                    name2 = f"- {val_tol}, {man}"
                elif len(f"{rem} {body} {tke}  - {pv} - {val_tol}, {man}") > shift_treshold:
                    name = f"{rem} {body} {tke}  - {pv} - {val_tol},"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} {tke}  - {pv} - {val_tol}, {man}"
            else:
                if len(f"{rem} {body} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol}, {man}"
                elif len(f"{rem} {body} - {pv} - {val_tol}, {man}") > shift_treshold:
                    name = f"{rem} {body} - {pv} - {val_tol},"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol}, {man}"
    elif 'R' in desig:
        if 'ТУ' in desig:
            if tke != '':
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"{val_tol} - {tke} {man}"
                    if len(name2) > shift_treshold:
                        name2 = f"- {val_tol} - {tke}"
                        name3 = f"{man}"
                elif len(f"{rem} - {pv} - {val_tol} - {tke}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol} -"
                    name2 = f"- {tke} {man}"
                elif len(f"{rem} - {pv} - {val_tol} - {tke} {man}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke} {man}"
            else:
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} - {tke} {man}"
                elif len(f"{rem} - {pv} - {val_tol} {tke}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol}"
                    name2 = f"- {tke} {man}"
                elif len(f"{rem} - {pv} - {val_tol} {tke} {man}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke} {man}"
        else:
            if len(f"{rem} {body}") > shift_treshold:
                name = f"{rem} {body}"
                name2 = f"- {val_tol}, {man}"
            elif len(f"{rem} {body} - {val_tol}") > shift_treshold:
                name = f"{rem} {body} -"
                name2 = f"- {val_tol}, {man}"
            elif len(f"{rem} {body} - {val_tol}, {man}") > shift_treshold:
                name = f"{rem} {body} - {val_tol},"
                name2 = f"{man}"
            else:
                name = f"{rem} {body} - {val_tol}, {man}"
    # Для прочих компонентов
    if 'C' not in desig and 'R' not in desig:
        if 'ТУ' in man:
            if len(f"{manpnb} {man}") > shift_treshold:
                name = f"{manpnb}"
                name2 = f"{man}"
            else:
                name = f"{manpnb} {man}"
        else:
            if len(f"{manpnb}, {man}") > shift_treshold:
                name = f"{manpnb},"
                name2 = f"{man}"
            else:
                name = f"{manpnb}, {man}"
        if manpnb == '':
            if man.find("ТУ", 0) != -1:
                name = f"{val} {man}"
            else:
                name = f"{val}, {man}"
    return [name, name2, name3]


def create_names_vp(df: pd.DataFrame, index: int, shift_treshold: int):
    """
    Формирует наименование компонента и переносит его, если тот не будет влезат в рамки шаблона

    :param df: Словарь с сортированными по обозначениям компонентами
    :param index: Индекс редактирумеого элемента
    :param shift_treshold: Порог длины строки, привышая который будет производиться перенос наимнования
    :return: Готовое наименование элемента, [name2, name3] - переносы
    """
    desig = df['Designator'][index]
    rem = df['Rem'][index]
    body = df['Корпус'][index]
    tke = df['TKE'][index]
    val = df['Value'][index]
    pv = df['Power/Voltage'][index]
    tol = df['Tolerance'][index]
    man = df['Manufacturer'][index]
    manpnb = df['ManufacturerPartNumber'][index]
    # Для R и C Номинал и Погрешность переносятся вместе, поэтому сохраняю их в одну строку
    val_tol = val + " " + tol

    name = ''
    name2 = ''

    if '0' not in body \
            and '1' not in body \
            and body != '':
        body = "\"" + body + "\""

    # Формирование наименований
    if 'C' in desig:
        if "ТУ" in man:
            if tke:
                # Костыль для некоторых корпусов
                if '0' not in body and 'М' not in body:
                    body = ''
                if len(f"{rem} {body} - {pv}") > shift_treshold:
                    name = f"{rem} {body} -"
                    name2 = f"- {pv} - {val_tol} - {tke}"
                elif len(f"{rem} {body} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol} - {tke}"
                elif len(f"{rem} {body} - {pv} - {val_tol} - {tke}") > shift_treshold:
                    name = f"{rem} {body} - {pv} - {val_tol} -"
                    name2 = f"- {tke}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol} - {tke}"
            else:
                if len(f"{rem} {body} - {pv}") > shift_treshold:
                    name = f"{rem} {body} -"
                    name2 = f"- {pv} - {val_tol}"
                elif len(f"{rem} {body} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol}"
        else:
            if tke:
                if len(f"{rem} {body} {tke} - {pv}") > shift_treshold:
                    name = f"{rem} {body} {tke} -"
                    name2 = f"- {pv} - {val_tol}"
                elif len(f"{rem} {body} {tke} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} {tke}  - {pv} -"
                    name2 = f"- {val_tol}"
                else:
                    name = f"{rem} {body} {tke}  - {pv} - {val_tol}"
            else:
                if len(f"{rem} {body} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol}"
    elif 'R' in desig:
        if 'ТУ' in man:
            if tke:
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} - {tke}"
                elif len(f"{rem} - {pv} - {val_tol} - {tke}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol} -"
                    name2 = f"- {tke}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
            else:
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} - {tke}"
                elif len(f"{rem} - {pv} - {val_tol} {tke}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol}"
                    name2 = f"- {tke}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
        else:
            if len(f"{rem} {body}") > shift_treshold:
                name = f"{rem} {body}"
                name2 = f"- {val_tol}"
            elif len(f"{rem} {body} - {val_tol}") > shift_treshold:
                name = f"{rem} {body} -"
                name2 = f"- {val_tol}"
            else:
                name = f"{rem} {body} - {val_tol}"
    # Для прочих компонентов
    if "C" not in desig and "R" not in desig:
        if 'ТУ' in man:
            name = manpnb
        else:
            name = manpnb
        if manpnb == '':
            name = val
    return [name, name2]


def only_dir(path: str) -> str:
    '''Из пути к файлу берёт только путь без имени'''
    return path[:path.rfind('\\') + 1]


def only_name(path: str) -> str:
    '''Из пути к файлу берёт только имя файла'''
    return path[path.rfind('/') + 1:]


def path_for_win(path: str) -> str:
    '''Заменяет слэши в строке с адресом на бекслэши'''
    path = path.replace('/', '\\')
    return path


