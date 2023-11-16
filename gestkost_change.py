# Программа для замены жесткостей в Лире САПР
# Таблица должна содержать Столбец с названием: "Жесткость	кэ	Цвет	Наименование	Значение жесткостей		Элемент	кэ	Узел		Элемент	Жесткость		Узел	x	y	z"
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
    sheet_base = book.sheets["base s1"]
    sheet_change = book.sheets["change s1"]

    table_base=fun_un_table(sheet_base)
    print(f"Длина таблицы 1{(table_base.axes[0])}")
    table_base["z"] = table_base["z"].apply(lambda x: round(x, 2))

    max_=table_base['z'].max()
    min_ = table_base['z'].min()



    table_change=fun_un_table(sheet_change)
    print(f"Длина таблицы 2{table_change.axes[0]}")
    table_change=table_change.drop(["Значение жесткостей","Цвет","Наименование"],axis=1)
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
    print(f"Проверка длины таблицы{len(table_base),len(table_change_up),len(table_change_down)}")
    print(f"Проверка длины таблицы {len(table_base),len(table_base_up),len(table_base_down)}")

    table_base_up=table_base_up.sort_values(by=["x","y"]).reset_index()
    table_change_up=table_change_up.sort_values(by=["x","y"]).reset_index()

    if max_!=min_:
        table_base_down=table_base_down.sort_values(by=["x","y"]).reset_index()
        table_change_down=table_change_down.sort_values(by=["x", "y"]).reset_index()

        table_base =pd.concat([table_base_up,table_base_down]).reset_index()
        table_change = pd.concat([table_change_up,table_change_down]).reset_index()

    table_change["Значение жесткостей"] =table_base["Значение жесткостей"]


    xlsheet = sheet_change
    xlsheet.range("S1").options(index=False).value = table_base
    xlsheet.range("AI1").options(index=False).value = table_change

    table_change["Значение жесткостей"] =table_base["Значение жесткостей"]
    table_change=table_change.sort_values(by=["Жесткость"])

    xlsheet.range("BC1").options(index=False).value = table_change

