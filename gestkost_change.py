

# Программа для замены жесткостей в Лире САПР. Заменяет жесткости в одноузловых конечных элементах.
# Таблица Base значения на которые происходит замена,
# Таблица должна содержать Столбец с названием: "Жесткость	кэ	Цвет	Наименование	Значение жесткостей		Элемент	кэ	Узел		Элемент	Жесткость		Узел	x	y	z"
# Первую таблицу Base делаем из схемы с сборкой вторую делаем из схемы в которой будем менять сваи

import pandas as pd
import xlwings as xw
import re
pd.options.display.max_rows = 10
pd.options.display.max_columns = 100
pd.set_option('expand_frame_repr', False)

def fun_un_table(sheet_name):
    sheet=sheet_name

    table_gestkosti = sheet.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    table_gestkosti = table_gestkosti.reset_index()

    table_elemet_uzel = sheet.range("G1").options(pd.DataFrame, expand='table', index_col=True).value
    table_elemet_uzel = table_elemet_uzel.reset_index()

    table_elemet_gestk = sheet.range("K1").options(pd.DataFrame, expand='table', index_col=True).value
    table_elemet_gestk = table_elemet_gestk.reset_index()

    table_uzel_coordinat = sheet.range("N1").options(pd.DataFrame, expand='table', index_col=True).value
    table_uzel_coordinat = table_uzel_coordinat.reset_index()

    new_table1 = pd.merge(table_elemet_uzel, table_uzel_coordinat, on="Узел")
    new_table1 = pd.merge(new_table1, table_elemet_gestk, on="Элемент")

    new_table1 = pd.merge(new_table1, table_gestkosti, on="Жесткость")
    return new_table1


def fun_Perebor(table_change, table_base, pred_l, rastoinie_z):
    """
    rast_zad-аданное расстояние между сваями (допустимое)
    Функция перебора точек по координатам
    :return:
    """
    for i in table_change.index:
        list_rast = []
        for k in table_base.index:
            rast = (
                pow(((table_change.loc[i, "x"] - table_base.loc[k, "x"]) ** 2) + (
                            (table_change.loc[i, "y"] - table_base.loc[k, "y"]) ** 2), 0.5))
            # print(rast)
            table_change.loc[i, "Test_rast"] = rast
            list_rast.append(rast)
            table_change.loc[i, "Test_rast"] = min(list_rast)  # Нахождение минимального расстояния от одной точки до другой
            coeff_heigt = rastoinie_z  # Разность координаты
            if rast < pred_l and table_change.loc[i, "z"] == table_base.loc[k, "z"] + coeff_heigt:
                table_change.loc[i, "Значение жесткостей"] = table_base.loc[k, "Значение жесткостей"]
                table_change.loc[i, "Наименование"] = table_base.loc[k, "Наименование"]
                table_change.loc[i, "Жесткость"] = table_base.loc[k, "Жесткость"]
                table_change.loc[i, "Цвет"] = table_base.loc[k, "Цвет"]
                break
            else:
                table_change.loc[i, "Значение жесткостей"] = "None"

def fun_start(base_name, change_name, pred_l, rastoinie_z):
    print(base_name,change_name)
    """
    base_name таблица на которую происходит замена
    change_name- таблица изменяемая 
    pred_l- расстояние между сваями по горизонтали допустимое
    rastoinie_z- расстояние между сваями по вертикали допустимое
    Функция управления программой
    :return:
    """
    sheet_base = book.sheets[base_name]  # Таблица на которую меняем
    sheet_change = book.sheets[change_name]# Изменяемая таблица исходная
    table_base=fun_un_table(sheet_base)
    table_base["z"] = table_base["z"].apply(lambda x: round(x, 2))
    table_change=fun_un_table(sheet_change)
    #table_change=table_change.drop(["Наименование","Жесткость"],axis=1)
    table_change["z"]=table_change["z"].apply(lambda x: round(x,2))
    print(f"Table chanhge \n {table_change} \n----")
    print(f"Table base \n {table_base} ")
    print(f"Проверка длины таблицы исходной{len(table_change)}")
    print(f"Проверка длины таблицы base {len(table_base)}")
    table_change["Test_rast"] = (pow(((table_change["x"] - table_base["x"]) ** 2) + ((table_change["y"] - table_base["y"]) ** 2), 0.5))
    fun_Perebor(table_change, table_base, pred_l, rastoinie_z)#Функция перебора
    table_change=table_change.sort_values(by=["Наименование"]).reset_index()

    table_change=table_change.drop(["index","кэ_x"],axis=1)
    table_change= table_change.iloc[:,[1, 2,  3, 4,0, 5, 6, 7, 8, 9, 10]]
    pd.options.display.max_rows = 10
    print(table_change)
    xlsheet = sheet_change
    xlsheet2 =sheet_base
    xlsheet2.range("S1").options(index=False).value = table_base#Вставка Базовая таблица на которую меняем значения
    xlsheet.range("S1").options(index=False).value = table_change#вставка Исходная
    print("Успешно")



if __name__ == '__main__':
    book = xw.books
    book = book.active
    sheet_param=book.sheets["Parameter"]# Таблица с параметрами для работы программы
    table_param:pd.DataFrame
    table_param = sheet_param.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    table_param = table_param.reset_index()
    print(table_param)
    table_param=table_param.apply(lambda x:fun_start(x["New_shema"],x["Ishodnay_shema"],x["Pred_L"],x["Rastoinie_z"]),axis=1)






