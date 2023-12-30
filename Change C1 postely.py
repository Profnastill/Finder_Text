# Программа замены коэффициента постели в пластинах



import pandas as pd
import xlwings as xw

import numpy as np
from shapely.geometry import Point, Polygon


import re
import re

pd.options.display.max_rows = 50
pd.options.display.max_columns = 100
pd.set_option('expand_frame_repr', False)


def inPolygon(x, y,koordinate_poly):
    c = 0
    xp=koordinate_poly[0:len(koordinate_poly):2]
    yp=koordinate_poly[1:len(koordinate_poly):2]
    for i in range(len(xp)):
        if (((yp[i] <= y and y < yp[i - 1]) or (yp[i - 1] <= y and y < yp[i])) and
                (x > (xp[i - 1] - xp[i]) * (y - yp[i]) / (yp[i - 1] - yp[i]) + xp[i])): c = 1 - c
    return c


print(inPolygon(100, 0, (-100, 100, 100, -100), (100, 100, -100, -100)))
def fun_Skleyka(table_elemet_uzel: pd.DataFrame, table_uzel_coordinat: pd.DataFrame):
    """
    Функция склейки функции с переименованием столбцов и выкидыванием лишних
    :param table_elemet_uzel:
    :param table_uzel_coordinat:
    :return:
    """
    table_uzel_coordinat_column_name = table_uzel_coordinat.columns  # Сюда сохраняем исходное названиеи столбцов таблицы
    a = 1
    for x in ["Узел1", "Узел2", "Узел3", "Узел4"]:
        print(table_elemet_uzel.columns)
        table_uzel_coordinat.columns = list(
            map(lambda x: re.sub(r"['x']|['y']|['z']", x + "_" + str(a), x), table_uzel_coordinat_column_name))
        table_elemet_uzel = pd.merge(table_elemet_uzel, table_uzel_coordinat, left_on=x,
                                     right_on=["Узел"], how="left")
        table_elemet_uzel = table_elemet_uzel.drop(["Узел"], axis=1)
        print(table_elemet_uzel)
        a += 1
    return table_elemet_uzel


def is_point_inside(point:list, triangle:list):
    """
    Функция детерминант находит лежит ли точка внутри треугольника
    :param point:
    :param triangle:
    :return:
    """
    print(triangle)
    x1, y1, x2, y2, x3, y3 = triangle

    x, y = point
    det_triangle = x1 * (y2 - y3) - y1 * (x2 - x3) + (x2 * y3 - x3 * y2)
    det_p1 = x * (y2 - y3) - y * (x2 - x3) + (x2 * y3 - x3 * y2)
    det_p2 = x1 * (y - y3) - y1 * (x - x3) + (x * y3 - x3 * y)
    det_p3 = x1 * (y2 - y) - y1 * (x2 - x) + (x2 * y - x * y2)
    print((det_p1 >= 0 and det_p2 >= 0 and det_p3 >= 0))
    return (det_p1 >= 0 and det_p2 >= 0 and det_p3 >= 0) or (det_p1 <= 0 and det_p2 <= 0 and det_p3 <= 0) and (
                det_triangle == det_p1 + det_p2 + det_p3)


def fun_un_table(sheet_name):
    """
        Функция объединения таблицы из эксель

    :param sheet_name: имя лста
    :return:
    """

    sheet = sheet_name
    table_gestkosti = sheet.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    table_gestkosti = table_gestkosti.reset_index()
    table_elemet_uzel: pd.DataFrame
    table_elemet_uzel = sheet.range("G1").options(pd.DataFrame, expand='table', index_col=True).value  # Таблица с координатами узлов
    table_elemet_uzel = table_elemet_uzel.reset_index()
    table_elemet_uzel["Узел1"] = table_elemet_uzel["Узел"].apply(lambda x: x.split(", ")[0])
    table_elemet_uzel["Узел2"] = table_elemet_uzel["Узел"].apply(lambda x: x.split(", ")[1])
    table_elemet_uzel["Узел3"] = table_elemet_uzel["Узел"].apply(lambda x: x.split(", ")[2])
    table_elemet_uzel["Узел4"] = table_elemet_uzel["Узел"].apply(lambda x: x.split(", ")[3] if len(x.split(", ")) >= 4 else None)
    table_elemet_uzel = table_elemet_uzel.drop(["Узел"], axis=1)

    table_uzel_coordinat: pd.DataFrame
    table_uzel_coordinat = sheet.range("K1").options(pd.DataFrame, expand='table', index_col=True).value
    table_uzel_coordinat = table_uzel_coordinat.reset_index()
    table_uzel_coordinat["Узел"] = table_uzel_coordinat.loc[:, "Узел"].apply(lambda x: str(int(x)))  # Тиражирование столбцов таблицы
    table_elemet_uzel = fun_Skleyka(table_elemet_uzel, table_uzel_coordinat)  # запуск соединителя таблиц
    table = pd.merge(table_elemet_uzel, table_gestkosti, on=["Элемент"])

    print(table_elemet_uzel)
    print(table)

    return table

def fun_refind_point(x, y, table:pd.DataFrame):
    """
    Функция перебора точек в таблице базовой и новой, для соотнощения координат попаданий в треугольники
    :param x:
    :param y:
    :param table:
    :return:
    """
    for i in range(len(table)):#для старой таблицы
        if is_point_inside([x, y] ,table[["x_1", "y_1", "x_2", "y_2", "x_3", "y_3"]].iloc[i].tolist()):
            return float(table.C1z.iloc[i])
    return None

def find_point_poly(x,y,polygon:list):
    """
    :param x:  кордината проверяемой точки
    :param y: кордината проверяемой точки
    :param polygon: список с координатами точкее [x1,y1,x2,y2,x3,y3]
    :return:
    """


    poly = Polygon([(polygon[0],polygon[1]), (polygon[2],polygon[3]), (polygon[4],polygon[5])])
    contains = np.vectorize(lambda p: poly.contains(Point(p)), signature='(n)->()')
    return contains

    #array([ True, False, False])




if __name__ == '__main__':
    base_name = "С1 HSL "
    change_name = "С1 HSL  change"
    book = xw.books
    book = book.active
    table_param: pd.DataFrame
    sheet_base = book.sheets[base_name]  # Таблица на которую меняем
    sheet_change = book.sheets[change_name]  # Таблица на которую меняем
    base_table: pd.DataFrame
    change_table: pd.DataFrame
    base_table = fun_un_table(sheet_base)
    change_table = fun_un_table(sheet_change)
    l, x = [0, 3]
    print(l, x)
    trt = base_table.apply(lambda row2: [[row2.x_1, row2.y_1], [row2.x_2, row2.y_2], [row2.x_3, row2.y_3]], axis=1)

    def _new():
        """
        перебор но не рабочая
        :return:
        """
        for i in range(len(base_table)):
            # print(base_table[["x_1","y_1","x_2","y_2","x_3","y_3"]].iloc[i].tolist())
            koord_treug = base_table[["x_1", "y_1", "x_2", "y_2", "x_3", "y_3"]].iloc[i].tolist()

            print(base_table["C1z"])
            change_table["Test"]= change_table.apply(lambda row: base_table.C1z.iloc[i] if (is_point_inside([row.x_1, row.y_1] ,koord_treug)) else False, axis = 1)


            change_table["Test"]= change_table.apply(lambda row: base_table.C1z.iloc[i] if (is_point_inside([row.x_1, row.y_1] ,koord_treug)) else False, axis = 1)
            """"
            Пробегаем по каждой точке таблицы change_table, если точка лежит выносим значение
            
            """
    change_table["C1z_new1"]=change_table.apply(lambda row:  find_point_poly(row.x_1, row.y_1, base_table), axis=1)
    change_table["C1z_new2"] = change_table.apply(lambda row: find_point_poly(row.x_2, row.y_2, base_table), axis=1)
    change_table["C1z_new3"] = change_table.apply(lambda row: find_point_poly(row.x_3, row.y_3, base_table), axis=1)


    #change_table["C1z_new1"]=change_table.apply(lambda row:  fun_refind_point(row.x_1, row.y_1, base_table), axis=1)
    #change_table["C1z_new2"] = change_table.apply(lambda row: fun_refind_point(row.x_2, row.y_2, base_table), axis=1)
    #change_table["C1z_new3"] = change_table.apply(lambda row: fun_refind_point(row.x_3, row.y_3, base_table), axis=1)





    print(change_table[["C1z_new1","C1z_new2","C1z_new3"]])
    change_table=change_table.reset_index()


    change_table["C1z_newwddd"] = change_table.loc[:,["C1z_new1","C1z_new2","C1z_new3"]].mean(skipna=True,axis=1)




    print(change_table)
    #change_table=change_table.query("C1z_new!=False")
    print(change_table)
    sheet_change.range("R1").options(index=False).value = change_table#Вставка Базовая таблица на которую меняем значения
