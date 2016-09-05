import pandas as pd
import numpy as np
import datetime
import re

def getEvent():
    excel_dir = ['/Users/moreoronce/Documents/WeeklyPrintNum/1.xlsx',
                 '/Users/moreoronce/Documents/WeeklyPrintNum/2.xlsx']
    tablename = ["未出票数", "已出票数"]
    i = 0