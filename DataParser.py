import copy
import os
import re
import string

import numpy
import pandas as pd


import DataModel
import Editors as ed
import DataModel as dm

# Программа для создания промежуточного перечня элементов, составленного из BOM'ом всех плат

no_perechen = 1  # 0 - ПЭ, 1 - Сп, 2 - ВП, 3 - ТЭО

prim_cats = {}  # Категории примечаний
prim_not_install = {}  # Категории для "Не устанавливать"
prog_dir = os.path.dirname(os.path.abspath(__file__))

pd.set_option('mode.chained_assignment', None)

test = False


def get_char(desig: str):
    """Убирает из десигнатора циферные обозначения"""
    new_char = ''
    for c in desig:
        if c.isalpha():
            new_char += c
    return new_char


def get_dfs(dm: DataModel, file_index: int):
    if dm.files:
        i = 0
        dm.dfs.clear()
        dm.dfs.append(pd.read_excel(dm.files[file_index],
                                    sheet_name='BOM', na_filter=False, usecols="A:K", skiprows=13))

        if no_perechen:
            dm.dfs[file_index] = dm.dfs[file_index][dm.dfs[file_index]['Примечание'] != 'Не устанавливать']

        dm.files[file_index] = ed.path_for_win(dm.files[file_index])
        i += 1
    else:
        print("Не найдено BOM'ов в каталоге...")
        return

    # Достать все буквенные обозначения с листа
    for dataframe in dm.dfs:
        for index, row in dataframe.iterrows():
            char = get_char(row['Designator'])
            if char == "":
                continue

            if char not in dm.dict_chars.keys():
                # Объединение разных обозначений одной и той же категории

                for c in 'XDHKS':
                    if row['Designator'][0] == c:
                        char = c

                if test == 1:
                    dm.dict_chars[char] = []
                else:
                    dm.dict_chars[char] = pd.DataFrame(columns=["Designator", "Rem", "Корпус", "TKE", "Value",
                                                                "Power/Voltage", "Manufacturer",
                                                                "ManufacturerPartNumber",
                                                                "Module", "Quantity"])
    # return dm.dfs, dm.files


def get_components(dm: DataModel):
    """
    Формирует промежуточный DataFrame

    :param dict_chars: Словарь с сортированными по обозначениям компонентами
    :param dfs: Лист несортированных перечней компонентов из разных BOM'ов
    :param files: Лист путей к файлам BOM
    :return: Сортированный перечень элементов по их позиционным обозначениям
    """
    errors = []

    for char in sorted(dm.dict_chars.keys()):
        dict_chars_index = 0
        for ind, df in enumerate(dm.dfs):
            df_index = 0



            module1 = pd.read_excel(dm.files[ind], sheet_name='BOM', header=None)
            module2 = module1.loc[7, 10]
            module = module2.split(' ')[0]


            while df_index < len(df.index):
                component = ed.Element(df.iloc[df_index])
                component.module = module

                if '*' in component.desig:
                    if 'Регулирование' not in dm.comment_categories.values():
                        dm.comment_categories['*' * (len(dm.comment_categories) + 1)] = 'Регулирование'

                if component.desig.find(char, 0) == 0:
                    component = add_value_char(component, char)
                    # Пропускаем компоненты с примечанием "Выполнен на плате"
                    if component.comment.find("Выполнен на плате", 0) != -1:
                        df_index += 1
                        continue

                    if not any(letter in component.desig for letter in string.ascii_uppercase):
                        df_index += 1
                        continue

                    # Уведомления о пустых полях
                    if component.man == "":
                        errors.append(f"Не заполнен производитель у элемента {component.desig}")

                    if component.desig.find("С", 0) != -1:
                        if component.pv == "":
                            errors.append(f"Не заполнена мощность у элемента {component.desig}")

                    if component.char == 'C' or component.char == 'C':
                        # У отчеки нет Rem
                        if component.rem == "" and component.man.find("ТУ", 0) == -1:
                            errors.append(f"Не заполнен Rem у элемента {component.desig}")
                        if component.val == "":
                            errors.append(f"Не заполнен номинал у элемента {component.desig}")
                        if component.body == "":
                            errors.append(f"Не заполнен корпус у элемента {component.desig}")
                        if component.tol == "":
                            errors.append(f"Не заполнена погрешность у элемента {component.desig}")

                    add_to_comments(component, char)
                    component, df_index = combine_following_chips(component, df, df_index)
                    dm.dict_chars[char].append(component)

                    dict_chars_index += 1
                else:
                    df_index += 1

        # Проверить, общий ли производитель у всех компонентов категории
        common_manufacture = True
        if len(dm.dict_chars[char]) == 1:
            common_manufacture = False
        else:
            i = iter(dm.dict_chars[char])
            while True:
                try:
                    next_component = next(i)
                except StopIteration:
                    break

                if dm.dict_chars[char][0].man != next_component.man:
                    common_manufacture = False
                    break
        if common_manufacture:
            dm.components_one_manuf_categories[char] = dm.dict_chars[char][0].man
            for i, component in enumerate(dm.dict_chars[char]):
                component.man = ''

        if no_perechen:
            if char == 'R' or char == 'C':
                dm.dict_chars[char] = sorted(dm.dict_chars[char], key=lambda x: [x.desig, x.rem, x.body, x.pv, x.tol,
                                                                                 x.val, x.module])
            else:
                dm.dict_chars[char] = sorted(dm.dict_chars[char], key=lambda x: [x.manpnb, x.module])

    if len(errors) != 0:
        print("\n*** *** ***\n"
              "Найдены проблемы:")
        for error in errors:
            print(error)
        print("*** *** ***\n")


def add_to_comments(chip: ed.Element, char):
    global prim_not_install
    if chip.comment != '' and "*" not in chip.comment:
        if chip.comment == 'Не устанавливать':
            if 'Не устанавливать' not in prim_cats.values():
                prim_cats['*' * (len(prim_cats) + 1)] = 'Не устанавливать'
            if char not in prim_not_install.keys():
                prim_not_install[char] = []
            if chip.desig not in prim_not_install[char]:
                prim_not_install[char].append(chip.desig)


def combine_chips_in_module(d_chars):
    """
    Группирует все одинаковые компоненты в рамках одного модуля

    :param d_chars: Словарь с сортированными по обозначениям компонентами
    :return: Словарь с сортированными по обозначениям компонентами и группированными элементами в нем
    """

    for char in sorted(d_chars.keys()):
        char_df_index = 0
        while char_df_index < len(d_chars[char]):
            edit_index = char_df_index + 1
            while edit_index < len(d_chars[char]):
                # Если список компонентов закончился, то перейти к следующему списку
                chip = d_chars[char][edit_index]
                same = 0  # Кол-во одинаковых компонентов

                # Проверка схожести компонента на те, что имеются в промежуточном листе
                if chip.rem == d_chars[char][char_df_index].rem \
                        and chip.body == d_chars[char][char_df_index].body \
                        and chip.pv == d_chars[char][char_df_index].pv \
                        and chip.tol == d_chars[char][char_df_index].tol \
                        and chip.val == d_chars[char][char_df_index].val \
                        and chip.tke == d_chars[char][char_df_index].tke \
                        and chip.man == d_chars[char][char_df_index].man \
                        and chip.manpnb == d_chars[char][char_df_index].manpnb \
                        and chip.module == d_chars[char][char_df_index].module:
                    # Складываем имеющиеся кол-во с тем, что у компонента
                    d_chars[char][char_df_index].quantity = \
                        d_chars[char][char_df_index].quantity + chip.quantity
                    same = 1

                # Добавление нового элемента, если тот не похож на имеющийся
                if same != 0:
                    d_chars[char][char_df_index].desig = \
                        d_chars[char][char_df_index].desig + ", " + chip.desig
                    del d_chars[char][edit_index]
                    edit_index -= 1

                edit_index += 1
            char_df_index += 1

        if no_perechen == 1:
            if char == 'R' or char == 'C':
                d_chars[char] = sorted(d_chars[char], key=lambda x: [x.rem, x.body, x.pv, x.tol,
                                                                     convert_to_simple_value_for_sort(x.val), x.module])
            else:
                d_chars[char] = sorted(d_chars[char], key=lambda x: [x.manpnb, x.module])

    return d_chars


def combine_following_chips(component: ed.Element, df: pd.DataFrame, df_index: int):
    same = 0
    # Проверка схожести компонента на те, что имеются в промежуточном листе
    while True:
        df_index += 1
        if df_index >= len(df.index):
            return component, df_index
        next_component = ed.Element(df.iloc[df_index])
        next_component = add_value_char(next_component, next_component.char)
        if all((next_component.body == component.body, next_component.val == component.val,
                next_component.tke == component.tke, next_component.pv == component.pv,
                next_component.tol == component.tol, next_component.man == component.man,
                next_component.manpnb == component.manpnb, next_component.module != component.module)):
            pass
        else:
            break
        # Складываем имеющиеся кол-во с тем, что у компонента
        component.quantity = component.quantity + next_component.quantity
        next_desig = next_component.desig
        add_to_comments(next_component, next_component.char)
        same += 1

    # Добавление нового элемента, если тот не похож на имеющийся
    if same == 1:
        df_index -= 1
        component.desig = f"{component.desig}, {next_desig}"
    if same > 1:
        if df_index + 1 < len(df.index):
            df_index -= 1
        component.desig = f"{component.desig}-{next_desig}"
    if same != 0:
        df_index += 1
    return component, df_index


def add_value_char(chip: ed.Element, char: str):
    """
    Добавляет буквенные обозначения к номиналу в соотвествии с видом компонента
    :param chip: Экземпляр компонента
    :param char: Буквенное обозначение компонента
    :return:
    """

    if chip.man.find("ТУ") != - 1 and chip.man[0:-3].find(" ") != -1:
        chip.man = chip.man[0:-3].replace(" ", "") + " ТУ"
    else:
        chip.man = chip.man.split(",")[0]
    if chip.man.find("ТУ") != - 1 and chip.man[0:-3].find(" ") != -1:
        if char == 'C':
            print("В поле элемента " + chip.desig +
                  " в модуле " + chip.module + " обнаружен пробел в ТУ")
        chip.man = chip.man[0:-3].replace(" ", "") + " ТУ"
    if char == 'C':
        if chip.val.find("Ф") == -1:
            if chip.val.find(" ") == -1:
                chip.val += ' '
            chip.val += 'Ф'
    if char == 'R':
        if chip.val.find("Ом") == -1:
            if chip.val.find(" ") == -1:
                chip.val += ' '
            chip.val += 'Ом'
    if char == 'L':
        if chip.val.find("Гн") == -1:
            if chip.val.find(" ") == -1:
                chip.val += ' '
            chip.val += 'Гн'
    return chip


def split_to_adjustable(dm: DataModel):
    for char in sorted(dm.dict_chars.keys()):
        d_chars_index = 0
        while d_chars_index < len(dm.dict_chars[char]):
            chip = dm.dict_chars[char][d_chars_index]

            if chip.val == '':
                d_chars_index += 1
                continue
            splited_values = chip.val.split(", ")
            if len(splited_values) != 1:
                for ind, value in enumerate(splited_values):
                    if ind == 0:
                        dm.dict_chars[char][d_chars_index].val = add_value_char(dm.dict_chars[char][d_chars_index],
                                                                                char).val
                        dm.dict_chars[char][d_chars_index].comment = '*'
                    else:
                        line = copy.copy(chip)
                        line.val = value
                        line.comment = '*'
                        line = add_value_char(line, line.char)
                        dm.dict_chars[char].insert(d_chars_index + 1, line)

            d_chars_index += 1


def combine_modules(dm: DataModel):
    """
    Комбинирует компоненты из разных модулей

    :param d_chars: Словарь с сортированными по обозначениям компонентами с группированными компонентами в одном модуле
    :return: Словарь с сортированными по обозначениям компонентами с группированными компонентами в разных модулях
    """
    for char in sorted(dm.dict_chars.keys()):
        new = True
        for ind, chip in enumerate(dm.dict_chars[char]):
            i = 0
            for i in range(0, len(dm.dict_chars[char])):
                next_chip = dm.dict_chars[char][i]
                if chip.rem == next_chip.rem and \
                        chip.body == next_chip.body and \
                        chip.val == next_chip.val and \
                        chip.tke == next_chip.tke and \
                        chip.pv == next_chip.pv and \
                        chip.tol == next_chip.tol and \
                        chip.man == next_chip.man and \
                        chip.manpnb == next_chip.manpnb and \
                        ((chip.module != '' and next_chip.module != '')
                         and chip.module != next_chip.module):
                    new = False
                    break

            # Добавление нового элемента, если тот не похож на имеющийся
            if new:
                pass
            # Комбинирование и пересортировка
            else:
                # Cделал комбинирование полей в виде списка для первого элемента, потому что так будет проще вставлять
                # в функции вставки в Word
                if chip.quantity is not list and chip.module is not list:
                    chip.quantity = [chip.quantity, next_chip.quantity, chip.quantity + next_chip.quantity]
                    chip.module = [chip.module, next_chip.module, '']

                del dm.dict_chars[char][i]
            new = True

        if no_perechen:
            if char == 'R' or char == 'C':
                dm.dict_chars[char] = sorted(dm.dict_chars[char], key=lambda x: [x.rem, x.body, x.pv, x.tol,
                                                                                 convert_to_simple_value_for_sort(
                                                                                     x.val), x.module])
            else:
                dm.dict_chars[char] = sorted(dm.dict_chars[char], key=lambda x: [x.manpnb, x.module])


def convert_to_simple_value_for_sort(v: str):
    """
    Приводит номинал к общему значению и изменяет его сокращенное название без приставки

    :param v: Строка с номиналом компонента
    :return: Строка с измененным номиналом компонента
    """

    if len(v.split()) > 1:
        value = float(v.split()[0].replace(',', '.', 1))
        char = v.split()[1]
        if 'п' in char:
            value = value * 10 ** -12
        if 'н' in char:
            value = value * 10 ** -9
        if 'мк' in char:
            value = value * 10 ** -6
        elif 'к' in char:
            value = int(value * 10 ** 3)

        return str(value)
    else:
        return v
