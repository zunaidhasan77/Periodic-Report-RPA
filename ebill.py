import pandas as pd
import numpy as np
import math

xls = pd.ExcelFile('.\periodic.xlsx')
ebill = pd.read_excel(xls, sheet_name='5.e-bill% Auto Payment  Overdue', header=None)
ebill.to_excel("modified_ebill_auto_payment.xlsx",index=False)