# -*- coding: utf-8 -*-
import datetime
import pandas as pd


def printnumdefault():
    excel_dir = ['/Users/moreoronce/Documents/WeeklyPrintNum/1.xlsx',
                 '/Users/moreoronce/Documents/WeeklyPrintNum/2.xlsx']
    tablename = ['未出票数', '已出票数']
    i = 0
    excel = dict()
    date = datetime.datetime.now().strftime("%Y-%m-%d")
    columnstitle = ['', '', '']
    while i < 2:
        try:
            df = pd.read_excel(excel_dir[i], skiprows=3)
        except IOError:
            print('表格缺失,请核对')
        else:
            #过滤内部销售数据
            filteresult = df[(df.销售渠道 != columnstitle[0]) & (df.销售渠道 != columnstitle[1]) & (df.销售渠道 != columnstitle[2])]
            #清洗后数据进行统计并添加至字典
            filteresult = filteresult['项目名称']
            filteresult = pd.DataFrame(filteresult.value_counts())
            filteresult = filteresult.rename(columns={'项目名称': tablename[i]})
            filteresult = filteresult.to_dict()
            i += 1
            excel.update(filteresult)
            print(excel)

    excel = pd.DataFrame.from_dict(excel)
    excel.to_excel('/Users/moreoronce/Documents/WeeklyPrintNum/' + date + '出票统计.xlsx', sheet_name='Sheet1')
    print('出票统计表格生成完毕')

printnumdefault()
