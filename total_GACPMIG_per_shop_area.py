import pandas as pd
import numpy as np
import math
import selenium

xls = pd.ExcelFile('.\periodic.xlsx')
total_ga = pd.read_excel(xls, sheet_name='7. Total GACPMIG per shop Area', header=None)
loop_index=1


for x in total_ga[1]:

    if str(total_ga.loc[5, loop_index]).isnumeric():
        if math.isnan(total_ga.loc[5,loop_index]) is False:
            total_ga.loc[7,loop_index]= total_ga.loc[5,loop_index]+total_ga.loc[6,loop_index]

    loop_index+=1
    if(loop_index==14):
        break
loop_index=1
sum_val=0
count=0
for x in total_ga[1]:
    if str(total_ga.loc[7,loop_index]).isnumeric():

        if math.isnan(total_ga.loc[7,loop_index]) is False:

            sum_val+=total_ga.loc[8,loop_index]

            count+=1
    loop_index+=1
    if loop_index==13:
        break
total_ga.loc[8,13]= sum_val/count
total_ga.loc[9,13] = total_ga.loc[7,13]/total_ga.loc[8,13]


total_ga.to_excel("modified_total_ga.xlsx",index=False)

selenium