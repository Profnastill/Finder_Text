# Программа для замены жесткостей в Лире САПР


import pandas as pd
import xlwings as xw
import re
pd.options.display.max_rows = 1000
pd.options.display.max_columns = 100




if __name__ == '__main__':
    book = xw.books
    sheet = book.active.sheets
    sheet=sheet.active

    table:pd.DataFrame

    table = sheet.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    print(table)
    table=table.reset_index()
    pattern="\d{3,}"

    #pattern = "[-+]?\d+"
    match=  table["G"].apply(lambda x: re.sub(pattern,r"A",x))
    #match = table["G"].apply(lambda x: str(int(re.findall(pattern, x)[2])/2))
    table = table["G"].apply(lambda x: re.sub(pattern, str(int(re.findall(pattern, x)[0])/2), x))
    print(table)
    xlsheet = sheet
    xlsheet.range("C1").options(index=False).value = table
