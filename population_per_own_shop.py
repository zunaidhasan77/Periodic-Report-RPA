import pandas as pd
import numpy as np
import math

xls = pd.ExcelFile('.\periodic.xlsx')
population_per_own_shop = pd.read_excel(xls, sheet_name='2.PopulationPer Own ShopArea m2', header=None)
loop_index = 1
ft_to_meter = population_per_own_shop.loc[2, 20]

for x in population_per_own_shop[1]:
    if str(population_per_own_shop.loc[5, loop_index]).isnumeric():
        if math.isnan(population_per_own_shop.loc[5, loop_index]) is False:
            # print(population_per_own_shop.loc[4,loop_index])
            population_per_own_shop.loc[6, loop_index] = population_per_own_shop.loc[4, loop_index] + \
                                                         population_per_own_shop.loc[5, loop_index]

    loop_index += 1
    if loop_index == 13:
        break
loop_index = 1
for x in population_per_own_shop[1]:
    if str(population_per_own_shop.loc[7, loop_index]).isnumeric():
        if math.isnan(population_per_own_shop.loc[7, loop_index]) is False:
            # print(population_per_own_shop.loc[4,loop_index])
            population_per_own_shop.loc[9, loop_index] = population_per_own_shop.loc[8, loop_index] * ft_to_meter

    loop_index += 1
    if loop_index == 13:
        break
loop_index = 1
count = 0
for x in population_per_own_shop[1]:
    if str(population_per_own_shop.loc[5, loop_index]).isnumeric():
        if math.isnan(population_per_own_shop.loc[5, loop_index]) is False:
            # print(population_per_own_shop.loc[4,loop_index])
            population_per_own_shop.loc[11, loop_index] = population_per_own_shop.loc[6, loop_index] / \
                                                          population_per_own_shop.loc[9, loop_index]
            count += 1

    loop_index += 1
    if loop_index == 13:
        break

sum_val = 0

loop_index = 1

for x in population_per_own_shop[1]:
    if str(population_per_own_shop.loc[5, loop_index]).isnumeric():

        if math.isnan(population_per_own_shop.loc[11, loop_index]) is False:
            # print(population_per_own_shop.loc[4,loop_index])

            sum_val+=population_per_own_shop.loc[11,loop_index]

    loop_index += 1

    if loop_index == 13:
        population_per_own_shop.loc[11,loop_index] = sum_val/count
        break

population_per_own_shop.to_excel("modified_pop_per_own_shop.xlsx", index=False)
# print(population_per_own_shop)
