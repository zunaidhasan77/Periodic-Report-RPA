import pandas as pd
import numpy as np
import math

xls = pd.ExcelFile('.\periodic.xlsx')
selfcarelogin = pd.read_excel(xls, sheet_name='6.Self Care Login', header=None)
loop_index=1
for x in selfcarelogin[1]:
    if str(selfcarelogin.loc[5, loop_index]).isnumeric():
        if math.isnan(selfcarelogin.loc[5,loop_index]) is False:
            selfcarelogin.loc[24,loop_index] = (selfcarelogin.loc[5,loop_index]+selfcarelogin.loc[10,loop_index])/(selfcarelogin.loc[16,loop_index]+selfcarelogin.loc[20,loop_index])
            selfcarelogin.loc[25, loop_index] = (selfcarelogin.loc[6, loop_index] + selfcarelogin.loc[
                11, loop_index]) / (selfcarelogin.loc[17, loop_index] + selfcarelogin.loc[21, loop_index])
            selfcarelogin.loc[26, loop_index] = (selfcarelogin.loc[5, loop_index] + selfcarelogin.loc[
                6, loop_index]+selfcarelogin.loc[10, loop_index] + selfcarelogin.loc[
                11, loop_index]) / (selfcarelogin.loc[16, loop_index] + selfcarelogin.loc[17, loop_index]+selfcarelogin.loc[20, loop_index] + selfcarelogin.loc[21, loop_index])


    loop_index+=1
    if(loop_index==25):
        break
selfcarelogin.to_excel("modified_self_care_login.xlsx",index=False)
# print(math.isnan(own_shop_cost.loc[0, 0]))
# exit(1)
