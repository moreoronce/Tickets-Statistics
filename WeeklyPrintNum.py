# -*- coding: utf-8 -*-
import datetime
import pandas as pd
from openpyxl import load_workbook

def printnumdefault():
    resource_dir = ['/Users/moreoronce/Documents/WeeklyPrintNum/1.xlsx',
                    '/Users/moreoronce/Documents/WeeklyPrintNum/2.xlsx']
    tablename = ['已出票数', '未出票数']
    i = 0
    excel = dict()
    totalSheet = dict()

    while i < 2:
        try:
            df = pd.read_excel(resource_dir[i], skiprows=3)
        except IOError:
            print('表格缺失,请核对')
        else:
            # 清洗后数据进行统计并添加至字典
            filteresult = df.子场次.str.split('\ ', 1, True)
            filteresult = pd.DataFrame(filteresult[0].value_counts()).rename(columns={0: tablename[i]}).to_dict()
            totalNum = pd.DataFrame(df.项目名称.value_counts()).rename(columns={'项目名称': tablename[i]}).to_dict()
            i += 1
            totalSheet.update(totalNum)
            excel.update(filteresult)

    # 输出合计数据
    date = datetime.datetime.now().strftime("%Y-%m-%d")
    excelDir = '/Users/moreoronce/Documents/WeeklyPrintNum/' + date + '出票统计.xlsx'
    totalSheet = pd.DataFrame(totalSheet).fillna(0)
    totalSheet.index.rename('项目名称', inplace=True)
    totalSheet = totalSheet.reset_index().sort_values('项目名称')
    totalSum = {'项目名称': '总计',
                '已出票数': totalSheet.已出票数.sum(),
                "未出票数": totalSheet.未出票数.sum()
                }
    totalSheet = totalSheet.append(totalSum, ignore_index=True)
    totalSheet = pd.DataFrame(totalSheet).set_index(totalSheet.项目名称)
    totalSheet = totalSheet.drop('项目名称', axis=1)
    totalSheet.to_excel(excelDir, sheet_name='合计')

    # 按照日期进行数据过滤
    excel = pd.DataFrame.from_dict(excel)
    excel.index.rename("项目", inplace=True)
    excel = excel.reset_index()
    excelInfo = excel.项目.str.split('\|', 1, True)
    excel.项目 = excelInfo[0]
    excel.insert(0, "日期", excelInfo[1])
    excel = excel.fillna(0)

    # 定义赛期时间
    begin = datetime.date(2016, 9, 25)
    end = datetime.date(2016, 10, 10)
    d = begin
    delta = datetime.timedelta(days=1)

    # 生成赛期每日数据
    while d < end:
        projectDate = (d.strftime("%Y-%m-%d"))
        sheet = excel[excel.日期 == projectDate]
        sheet = pd.DataFrame(sheet)

        sheetsum = {'项目': '总计',
                    "日期": '',
                    '已出票数': sheet.已出票数.sum(),
                    '未出票数': sheet.未出票数.sum()
                    }
        sheet = sheet.sort_values('项目')
        sheet = sheet.append(sheetsum, ignore_index=True)
        sheet = sheet.set_index(sheet.项目)
        sheet = sheet.drop(['项目', '日期'], axis=1)
        book = load_workbook(excelDir)
        writer = pd.ExcelWriter(excelDir, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        sheet.to_excel(writer, sheet_name=projectDate)
        writer.save()
        d +=delta

    print('出票统计表格生成完毕')

printnumdefault()
