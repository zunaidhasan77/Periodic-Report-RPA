import pandas as pd
import numpy as np
import math

xls = pd.ExcelFile('.\periodic.xlsx')
app_user_customer_base = pd.read_excel(xls, sheet_name='3.App Users Customer Base (RGB', header=None)
loop_index = 1
print(str(app_user_customer_base.loc[4,1]).isdigit())
print(type(app_user_customer_base.loc[4,1]))
for x in app_user_customer_base[1]:
    if (str(app_user_customer_base.loc[4, loop_index]).isnumeric()):
        if (math.isnan(app_user_customer_base.loc[4, loop_index]) is False):
            app_user_customer_base.loc[6, loop_index] = app_user_customer_base.loc[4, loop_index] / \
                                                        app_user_customer_base.loc[5, loop_index]

    loop_index += 1
    if (loop_index == 12):
        break
app_user_customer_base.to_excel("modified_app_user_customer_base.xlsx",index=False)
