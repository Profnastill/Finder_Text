# Программа замены коэффициента постели в пластинах


import pandas as pd
import xlwings as xw
import threading
import asyncio

import numpy as np
from shapely.geometry import Point, Polygon
import re

pd.options.display.max_rows = 50
pd.options.display.max_columns = 100
pd.set_option('expand_frame_repr', False)


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

def is_point_inside(point: list, triangle: list):
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
    table_elemet_uzel = sheet.range("G1").options(pd.DataFrame, expand='table',
                                                  index_col=True).value  # Таблица с координатами узлов
    table_elemet_uzel = table_elemet_uzel.reset_index()
    table_elemet_uzel["Узел1"] = table_elemet_uzel["Узел"].apply(lambda x: x.split(", ")[0])
    table_elemet_uzel["Узел2"] = table_elemet_uzel["Узел"].apply(lambda x: x.split(", ")[1])
    table_elemet_uzel["Узел3"] = table_elemet_uzel["Узел"].apply(lambda x: x.split(", ")[2])
    table_elemet_uzel["Узел4"] = table_elemet_uzel["Узел"].apply(
        lambda x: x.split(", ")[3] if len(x.split(", ")) >= 4 else None)
    table_elemet_uzel = table_elemet_uzel.drop(["Узел"], axis=1)

    table_uzel_coordinat: pd.DataFrame
    table_uzel_coordinat = sheet.range("K1").options(pd.DataFrame, expand='table', index_col=True).value
    table_uzel_coordinat = table_uzel_coordinat.reset_index()
    table_uzel_coordinat["Узел"] = table_uzel_coordinat.loc[:, "Узел"].apply(
        lambda x: str(int(x)))  # Тиражирование столбцов таблицы
    table_elemet_uzel = fun_Skleyka(table_elemet_uzel, table_uzel_coordinat)  # запуск соединителя таблиц
    table = pd.merge(table_elemet_uzel, table_gestkosti, on=["Элемент"])

    print(table_elemet_uzel)
    print(table)

    return table


def fun_refind_point_OLD(x, y, table: pd.DataFrame):#Старая функция
    """
    Функция перебора точек в таблице базовой и новой, для соотнощения координат попаданий в треугольники
    :param x:
    :param y:
    :param table:
    :return:
    """
    for i in range(len(table)):  # для старой таблицы
        if is_point_inside([x, y], table[["x_1", "y_1", "x_2", "y_2", "x_3", "y_3"]].iloc[i].tolist()):
            return float(table.C1z.iloc[i])
    return None


def fun_refind_point2(x, y, table: pd.DataFrame):
    """
    Функция перебора точек в таблице базовой и новой, для соотнощения координат попаданий в треугольники
    :param x:
    :param y:
    :param table:
    :return:
    """
    for i in range(len(table)):  # для старой таблицы
        old_coord_list = table[["x_1", "y_1", "x_2", "y_2", "x_3", "y_3"]].iloc[i].tolist()
        if find_point_poly(point_coord=[x, y], polygon=old_coord_list):
            return float(table.C1z.iloc[i])
    return None


def find_point_poly(point_coord, polygon: list):
    """
    :point_coord координаты x,y
    Через библиотеку поиск точки в полигоне
    :param x:  кордината проверяемой точки
    :param y: кордината проверяемой точки
    :param polygon: список с координатами точкее [x1,y1,x2,y2,x3,y3]
    :return:
    """
    #print("-----point \n", point_coord)
    #print("-----polygon\n", polygon)
    #print([(polygon[0], polygon[1]), (polygon[2], polygon[3]), (polygon[4], polygon[5])])
    poly = Polygon([(polygon[0], polygon[1]), (polygon[2], polygon[3]), (polygon[4], polygon[5])])
    #print(poly.area)
    pt1 = Point(point_coord)  # создаем объект точка
    intersect = poly.intersection(pt1)
    if pt1 == intersect:
        # breakpoint()
        #print(True, pt1, intersect)
        return (True)
    else:
        #print(False, pt1, intersect)
        return False

    # array([ True, False, False])


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

    #e1=threading.Event()
    #e2 = threading.Event()
    #e3 = threading.Event()

    async def event1(task):
        print(task)
        await asyncio.sleep(0)
        change_table["C1z_new1"] = change_table.apply(lambda row: fun_refind_point2(row.x_1, row.y_1, base_table), axis=1)
    async def event2(task):
        print(task)
        await asyncio.sleep(0)
        change_table["C1z_new2"] = change_table.apply(lambda row: fun_refind_point2(row.x_2, row.y_2, base_table), axis=1)
    async def event3(task):
        print(task)
        await asyncio.sleep(0)
        change_table["C1z_new3"] = change_table.apply(lambda row: fun_refind_point2(row.x_3, row.y_3, base_table), axis=1)

    async def main():

        taskA = loop.create_task(event1('taskA'))
        taskB = loop.create_task(event2('taskB'))
        taskC = loop.create_task(event3('taskC'))

        #taskA = loop.create_task(event1('taskA'))
        #taskB = loop.create_task(event2('taskB'))
        #taskC = loop.create_task(event3('taskC'))
        await asyncio.sleep(0.5)

    try:
        loop = asyncio.new_event_loop()
        loop.run_until_complete(main())
        loop.close()

    except:
        pass


    print(change_table[["C1z_new1", "C1z_new2", "C1z_new3"]])
    change_table = change_table.reset_index()
    change_table["C1z_newwddd"] = change_table.loc[:, ["C1z_new1", "C1z_new2", "C1z_new3"]].mean(skipna=True, axis=1)

    print(change_table)
    # change_table=change_table.query("C1z_new!=False")
    print(change_table)
    sheet_change.range("R1").options(
        index=False).value = change_table  # Вставка Базовая таблица на которую меняем значения
