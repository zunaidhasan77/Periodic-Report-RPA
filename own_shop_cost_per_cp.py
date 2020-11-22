import pandas as pd
import numpy as np
import math

xls = pd.ExcelFile('.\periodic.xlsx')
own_shop_cost = pd.read_excel(xls, sheet_name='1. Own Shop Cost Per GA cp', header=None)
loop_index = 3
eur_val = own_shop_cost.loc[7, 1]
print()
# print(math.isnan(own_shop_cost.loc[0, 0]))
# exit(1)
sum_val=0
for x in own_shop_cost[2]:
    if str(own_shop_cost.loc[6, loop_index]).isnumeric():
        if math.isnan(own_shop_cost.loc[6, loop_index]) is False:
            # print(own_shop_cost.loc[6, loop_index])
            sum_val+= own_shop_cost.loc[6,loop_index]

    loop_index += 1
    if loop_index == 15:
        own_shop_cost.loc[6,loop_index]=sum_val
        break


loop_index=3
for x in own_shop_cost[2]:
    if str(own_shop_cost.loc[6, loop_index]).isnumeric():
        if math.isnan(own_shop_cost.loc[6, loop_index]) is False:
            # print(own_shop_cost.loc[6, loop_index])
            own_shop_cost.loc[5, loop_index] = own_shop_cost.loc[6, loop_index] / 10 ** 6
    loop_index += 1
    if loop_index == 16:
        break
loop_index=3
for x in own_shop_cost[2]:
    if str(own_shop_cost.loc[6, loop_index]).isnumeric():
        if math.isnan(own_shop_cost.loc[6, loop_index]) is False:
            # print(own_shop_cost.loc[6, loop_index])
            own_shop_cost.loc[7, loop_index] = own_shop_cost.loc[6, loop_index] *eur_val
    loop_index += 1
    if loop_index == 16:
        break
sum_val=0
loop_index=3
for x in own_shop_cost[2]:
    if str(own_shop_cost.loc[8, loop_index]).isnumeric():
        if math.isnan(own_shop_cost.loc[8, loop_index]) is False:
            # print(own_shop_cost.loc[6, loop_index])
            sum_val+= own_shop_cost.loc[8,loop_index]

    loop_index += 1
    if loop_index == 15:
        own_shop_cost.loc[8,loop_index]=sum_val
        break
sum_val=0
loop_index=3
for x in own_shop_cost[2]:
    if str(own_shop_cost.loc[9, loop_index]).isnumeric():
        if math.isnan(own_shop_cost.loc[9, loop_index]) is False:
            # print(own_shop_cost.loc[6, loop_index])
            sum_val+= own_shop_cost.loc[9,loop_index]

    loop_index += 1
    if loop_index == 15:
        own_shop_cost.loc[9,loop_index]=sum_val
        break

loop_index=3
for x in own_shop_cost[2]:
    if str(own_shop_cost.loc[6, loop_index]).isnumeric():
        if math.isnan(own_shop_cost.loc[6, loop_index]) is False:
            # print(own_shop_cost.loc[6, loop_index])
            own_shop_cost.loc[10, loop_index] = own_shop_cost.loc[7, loop_index] / (own_shop_cost.loc[8,loop_index]+own_shop_cost.loc[9,loop_index])
    loop_index += 1
    if loop_index == 16:
        break
own_shop_cost.to_excel("modified_own_shop.xlsx",index=False)
