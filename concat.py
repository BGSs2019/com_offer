import pandas as pd
import numpy as np
import os

#concat function
def concat(calc1, calc2):
    #set doc name
    xlsx_name_list = calc1.split(sep = ".")
    xlsx_name_list = xlsx_name_list[:-1]
    xlsx_name_str = ""

    for elem in xlsx_name_list:
        xlsx_name_str = xlsx_name_str + elem
    xlsx_name = xlsx_name_str + "_1.xlsx"
    
    excel1 = pd.read_excel(calc1)
    excel2 = pd.read_excel(calc2)
    result = pd.concat([excel1, excel2.iloc[:,5:]], axis=1)
    result.set_index(result.iloc[:,0], inplace = True)
    result.iloc[:,1:].to_excel(xlsx_name)


# choose file
from tkinter.filedialog import askopenfilenames
initdir=os.getcwd()
calculations1 = []
calculations1 = askopenfilenames(initialdir=initdir, title="Choose files")
calculations1_names = []
for calc in calculations1:
    calc = calc.split(sep = "/")[-1]
    calculations1_names.append(calc)
print(calculations1_names)
calculations2 = []
calculations2 = askopenfilenames(initialdir=initdir, title="Choose files")
calculations2_names = []
for calc in calculations2:
    calc = calc.split(sep = "/")[-1]
    calculations2_names.append(calc)
print(calculations2_names)


# main cycle
for i, calc in enumerate(calculations1_names):
    concat(calculations1_names[i], calculations2_names[i])