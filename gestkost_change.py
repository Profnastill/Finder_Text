# Программа для замены жесткостей в Лире САПР. Заменяет жесткости в одноузловых конечных элементах.
# Таблица Base значения на которые происходит замена,
# Таблица должна содержать Столбец с названием: "Жесткость	кэ	Цвет	Наименование	Значение жесткостей		Элемент	кэ	Узел		Элемент	Жесткость		Узел	x	y	z"
# Первую таблицу Base делаем из схемы с сборкой вторую делаем из схемы в которой будем менять сваи

import pandas as pd
import xlwings as xw
import re
pd.options.display.max_rows = 20
pd.options.display.max_columns = 100
pd.set_option('expand_frame_repr', False)

def fun_un_table(sheet_name):
    sheet=sheet_name

    table_gestkosti = sheet.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    print(table_gestkosti)
    table_gestkosti = table_gestkosti.reset_index()

    table_elemet_uzel = sheet.range("G1").options(pd.DataFrame, expand='table', index_col=True).value
    table_elemet_uzel = table_elemet_uzel.reset_index()

    table_elemet_gestk = sheet.range("K1").options(pd.DataFrame, expand='table', index_col=True).value
    table_elemet_gestk = table_elemet_gestk.reset_index()

    table_uzel_coordinat = sheet.range("N1").options(pd.DataFrame, expand='table', index_col=True).value
    table_uzel_coordinat = table_uzel_coordinat.reset_index()

    new_table1 = pd.merge(table_elemet_uzel, table_uzel_coordinat, on="Узел")
    new_table1 = pd.merge(new_table1, table_elemet_gestk, on="Элемент")
    print(new_table1)
    new_table1 = pd.merge(new_table1, table_gestkosti, on="Жесткость")
    return new_table1

if __name__ == '__main__':
    book = xw.books
    book = book.active
    #sheet=sheet.active
    sheet_base = book.sheets["Base s1(Mor)"]#Таблица на которую меняем
    sheet_change = book.sheets["change s1(Mor)"]# Изменяемая таблица исходная
    #sheet_base = book.sheets["Base s1(HSL)"]#Таблица на которую меняем
    #sheet_change = book.sheets["change s1(HSL)"]# Изменяемая таблица исходная
    table_base=fun_un_table(sheet_base)
    print(f"Длина таблицы 1{(table_base.axes[0])}")


    table_base["z"] = table_base["z"].apply(lambda x: round(x, 2))

    max_=table_base['z'].max()
    min_ = table_base['z'].min()

    table_change=fun_un_table(sheet_change)
    print(f"Длина таблицы 2{table_change.axes[0]}")
    #table_change=table_change.drop(["Наименование","Жесткость"],axis=1)
    table_change["z"]=table_change["z"].apply(lambda x: round(x,2))

    print(table_change)
    print(table_base)

    table_base_up=table_base.query("z==@max_")
    print(len(table_base_up))
    table_base_down = table_base.query("z==@min_")
    print(len(table_base_down))

    print(max_,min_)
    print(table_change['z'].max())
    print(table_change['z'].min())

    table_change_up=table_change.query("z==@max_")
    table_change_down = table_change.query("z==@min_")
    print(f"Проверка длины таблицы исходной{len(table_change),len(table_change_up),len(table_change_down)}")
    print(f"Проверка длины таблицы base {len(table_base),len(table_base_up),len(table_base_down)}")

    table_base_up=table_base_up.sort_values(by=["x","y"]).reset_index()
    table_change_up=table_change_up.sort_values(by=["x","y"]).reset_index()



    if max_!=min_:
        table_base_down=table_base_down.sort_values(by=["x","y"]).reset_index()
        table_change_down=table_change_down.sort_values(by=["x", "y"]).reset_index()

        table_base =pd.concat([table_base_up,table_base_down]).reset_index()
        table_change = pd.concat([table_change_up,table_change_down]).reset_index()
    else:
        table_base = table_base.sort_values(by=["x", "y"]).reset_index()
        table_change=table_change.sort_values(by=["x","y"]).reset_index()

    
    table_change["Test_rast1"]=(pow(((table_change["x"]-table_base["x"])**2)+((table_change["y"]-table_base["y"])**2),0.5))#Для проверки расстояние между новыми и старыми узлами
    table_change["Test_rast"] = (pow(((table_change["x"] - table_base["x"]) ** 2) + ((table_change["y"] - table_base["y"]) ** 2), 0.5))

#print(table_change)

for i in table_change.index:
    list_rast=[]
    for k in table_base.index:

        rast= (
            pow(((table_change.loc[i,"x"]- table_base.loc[k,"x"]) ** 2) + ((table_change.loc[i,"y"] - table_base.loc[k,"y"]) ** 2), 0.5))
        #print(rast)
        table_change.loc[i,"Test_rast"] = rast

        list_rast.append(rast)
        table_change.loc[i,"Test_rast"] = min(list_rast)#Нахождение минимального расстояния от одной точки до другой
        if rast<1:
            table_change.loc[i,"Значение жесткостей"] = table_base.loc[k,"Значение жесткостей"]
            table_change.loc[i,"Жесткость"] = table_base.loc[k,"Жесткость"]
            table_change.loc[i,"Цвет"] = table_base.loc[k,"Цвет"]
            break
        else:
            table_change.loc[i,"Значение жесткостей"]="None"


    pd.options.display.max_rows = 10
    print(table_change)
    xlsheet = sheet_change
    xlsheet2 =sheet_base
    xlsheet2.range("S1").options(index=False).value = table_base#Вставка Базовая таблица на которую меняем значения
    xlsheet.range("S1").options(index=False).value = table_change#вставка Исходная

    print("Успешно")
    table_change=table_change.sort_values(by=["Жесткость"])





