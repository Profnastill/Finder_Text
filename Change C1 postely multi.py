# Программа замены коэффициента постели в пластинах
import multiprocessing

# Мультипроцессор


import pandas as pd
import xlwings as xw
from multiprocessing import Process,Pool
import time
from tqdm import tqdm

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

def fun_refind_point2(x, y, table: pd.DataFrame):
    """
    Функция перебора точек в таблице базовой и новой, для соотнощения координат попаданий в треугольники
    :param x:
    :param y:
    :param table:
    :return:
    """

    for i, row in table.iterrows():  # для старой таблицы
        old_coord_list = table[["x_1", "y_1", "x_2", "y_2", "x_3", "y_3","x_4","y_4"]].iloc[i].tolist()
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
    if pd.isnull(polygon[6]):
        poly = Polygon([(polygon[0], polygon[1]), (polygon[2], polygon[3]), (polygon[4], polygon[5])])
    else:
        poly = Polygon([(polygon[0], polygon[1]), (polygon[2], polygon[3]), (polygon[4], polygon[5]),(polygon[6], polygon[7])])
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


class SboRka:
    def __init__(self,change_table:pd.DataFrame,base_table:pd.DataFrame):
        self.change_table=change_table
        self.base_table=base_table
    def one_proc_main(self):
        """
        Старое
        :return:
        """
        self.change_table["C1z_new1"] = self.event("task1", "x_1", "y_1")
        self.change_table["C1z_new2"] = self.event("task2", "x_2", "y_2")
        self.change_table["C1z_new3"] = self.event("task3", "x_3", "y_3")

    def event(self,task,change_name,x,y):
        """
        Функция перебора
    
        :param task: 
        :param x: Координата
        :param y:  Координата
        :return: столбец дата фрейма

        """
        print(task)
        sheet_change:pd.DataFrame
        self.change_table=self.change_table.sort_values(by=[x,y])
        self.base_table = self.base_table.sort_values(by=[x, y])
        self.change_table[change_name] = self.change_table.progress_apply(lambda row: fun_refind_point2(row[x], row[y], self.base_table), axis=1)
        print(self.change_table)
        return self.change_table

    def main(self):
        """
        Многопроцессорность
        Реализация через Process
        :return:
        """
        # event1("task1")
        # event2("task2")
        # event3("task3")
        print("старт мультипроцессора")
        process1 = Process(target=self.event, args=("task1","C1z_new1", "x_1", "y_1",))
        process2 = Process(target=self.event, args=("task2","C1z_new2", "x_2", "y_2",))
        process3 = Process(target=self.event, args=("task3","C1z_new3", "x_3", "y_3",))
        process1.start()
        process2.start()
        process3.start()

        process1.join()
        process2.join()
        process3.join()
        print("f-----fff\n",self.change_table)
    def main_Pool(self):
        """
        Многопроцессорность
        Реализация через pool  полностью рабочая
        :return:
        """
        pool=Pool(4)
        process1 = pool.apply_async(self.event,("task1","C1z_new1", "x_1", "y_1",))
        process2 = pool.apply_async(self.event, ("task2","C1z_new2", "x_2", "y_2",))
        process3 = pool.apply_async(self.event, ("task3","C1z_new3", "x_3", "y_3",))
        process4 = pool.apply_async(self.event, ("task4", "C1z_new4", "x_4", "y_4",))
        pool.close()
        pool.join()
        self.change_table=pd.concat([process1.get(),process2.get()["C1z_new2"],process3.get()["C1z_new3"],process4.get()["C1z_new4"]],axis=1)
        #print(f"jghkdhgkd\n{process1.get()}\n {process2.get()}\n {process3.get()}")

    def insert_to_excel(self):
        print(self.change_table[["C1z_new1", "C1z_new2", "C1z_new3","C1z_new4"]])
        self.change_table = self.change_table.reset_index()
        filter1=self.change_table.query('y_4!=y_4')#Находим значение c Nan
        filter2=self.change_table.query('y_4==y_4')#Находим остальные значения
        filter1.loc[:,"C1z_new4"]=np.NAN
        self.change_table=pd.concat([filter1,filter2],axis=0)
        self.change_table["C1z_mean"] = self.change_table.loc[:, ["C1z_new1", "C1z_new2", "C1z_new3","C1z_new4"]].mean(skipna=True,axis=1)
        print(self.change_table)
        # change_table=change_table.query("C1z_new!=False")
        print(self.change_table)
        sheet_change.range("R1").options(
            index=False).value = self.change_table  # Вставка Базовая таблица на которую меняем значения



base_name = "С1 HSL"# Таблица на которую меняем
change_name = "С1 HSL  change"# Таблица которую будем менять

book = xw.books
book = book.active
table_param: pd.DataFrame
sheet_base = book.sheets[base_name]  # Таблица на которую меняем
sheet_change = book.sheets[change_name]  # Таблица на которую меняем
base_table: pd.DataFrame
change_table: pd.DataFrame
base_table = fun_un_table(sheet_base)
change_table = fun_un_table(sheet_change)
tqdm.pandas(desc="power DataFrame 1M to 100 random int!")



if __name__ == '__main__':
    n_cores=4
    # with multiprocessing.Pool(4) as p:
    #     rezults=p.map(main())
    # main()
    # insert_to_excel()
    a=SboRka(change_table,base_table)
    # a.main()
    a.main_Pool()
    print(a.change_table)
    a.insert_to_excel()


