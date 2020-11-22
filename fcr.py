import pandas as pd
import numpy as np
import math


def is_number(string):
    try:
        float(string)
        return True
    except ValueError:
        return False


xls = pd.ExcelFile('.\periodic.xlsx')
fcr = pd.read_excel(xls, sheet_name='4.FCR', header=None)
loop_index = 2
sum_val=0
count=0

for x in fcr[0]:

    if (is_number(str(fcr.loc[4, loop_index]))):

        if (math.isnan(fcr.loc[4, loop_index]) is False):
            sum_val+=fcr.loc[4,loop_index]
            count+=1


    loop_index += 1
    if (loop_index == 15):
        break
fcr.loc[4,14] = sum_val/count

loop_index = 2
sum_val=0
count=0

for x in fcr[0]:

    if (is_number(str(fcr.loc[5, loop_index]))):

        if (math.isnan(fcr.loc[5, loop_index]) is False):
            sum_val+=fcr.loc[5,loop_index]
            count+=1


    loop_index += 1
    if (loop_index == 15):
        break
fcr.loc[5,14] = sum_val/count
loop_index = 1
sum_val=0
count=0

for x in fcr[0]:

    if (is_number(str(fcr.loc[4, loop_index]))):

        if (math.isnan(fcr.loc[4, loop_index]) is False):
            fcr.loc[6,loop_index] = (fcr.loc[5,loop_index] + fcr.loc[4,loop_index])/2



    loop_index += 1
    if (loop_index == 15):
        break
fcr.to_excel("modified_fcr.xlsx",index=False)