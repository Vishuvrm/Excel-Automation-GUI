from type_extract import TypeExtract

cond = """
# value, column, data, showall(Bool)
showall = True
x1,x2,x3,x4 = 0,0,0,0

if column == "Index" or column == "Emp ID": 
    for val3 in data[column]:
        if 50 < val3 < 100:
            if val3 == value:
                x1 = True
                break

if x1:
"""

cond1 = """
showall = False
if type(value)==str:
    if len(value) >= 6:
"""
cond1_a = """
showall = True
if type(value)==str:
    if len(value) >= 6:
"""

cond2 = """
column=column
data=data
value=value
showall =True

x = 0

if column == "Emp ID":
    if type(value) == int:
        if (value) % 10 == 0:

"""

cond3 = """
if type(value) == str or type(value)==int:
"""


extr_str = TypeExtract("E:\MS EXCEL\Controlling EXCEL with PYTHON\Success.xlsx", sheet="advanced").copy_data(personal_data=[[1,2,3],[4,5,6],[7,8,9]])
data = extr_str.get_data
# print(extr_str.get_header)
# print(data)
# extr_int = TypeExtract("E:\MS EXCEL\Controlling EXCEL with PYTHON\Modified_data.xlsx", "advanced").copy_data(copy_to_different_wb="E:\MS EXCEL\Controlling EXCEL with PYTHON\Success.xlsx",
#                                                                                                               type_=str, first_row_header=True )#.append_to_cols(data, sheet="string_data")
# data = extr_int.get_data

# extr1, extr2 = TypeExtract("E:\MS EXCEL\Controlling EXCEL with PYTHON\Success.xlsx", sheet="advanced"), TypeExtract("E:\MS EXCEL\Controlling EXCEL with PYTHON\Success.xlsx", sheet="advanced")
# data = extr1.copy_data(type_=int, use_header=False).get_data
for r in data:
    print(r)

print("================================================================================================")

# data2 = extr2.copy_data(type_=str).get_data
# # data2 = extr_str.copy_data(custom_condition=cond1_a).get_data
#
# for r in data2:
#     print(r)
# extr_str.append_to_cols(data=data2, sheet="Abhi")
# extr_str.save_()

execute = """
# DATA, COLUMNS, SHEET, CELL

print(sheet.title)
column = "Emp ID"
index = 1
vals = column_values[column]

columns[column] = [# if i==None else i for i in vals]

"""

execute1 = """
print(SHEET.title)
for col in COLUMNS:
    COLUMNS[col] = ['#' if i==None else i for i in COLUMNS[col]]
"""

execute2 = """
import numpy as np
print(SHEET.title)
cols = COLUMNS.copy()
for col in cols:
    # print(col)
    COLUMNS[col+"_length"] = [len(val) if val != None else 0 for val in COLUMNS[col]]
"""

execute3 = """
print(SHEET.title)
index = 0
for col in COLUMNS:
    if col.endswith("length"):
        break
    index += 1

avg = []
for row in DATA:
    avg.append(sum(row[index:])/len(row[index:]))    

COLUMNS["Average_of_string_length"] = avg
"""
# final_data = extr_str.execute(execute2).get_data#.copy_data("E:\MS EXCEL\Controlling EXCEL with PYTHON\Success.xlsx",)
# final_data = extr_str.execute(execute3).get_data.copy_data("E:\MS EXCEL\Controlling EXCEL with PYTHON\Success.xlsx",)
# final_data = extr_str.execute(execute3).get_data

# for data in final_data:
#     print(data)

# extr_str.append_to_cols(data=final_data, sheet="Abhi").save_()
# print(modified_data.get_header)
# data = modified_data.get_data
# for r in data:
#     print(r)

# modified_data.copy_data(copy_to_different_wb="E:\MS EXCEL\Controlling EXCEL with PYTHON\Success.xlsx", sheet = "Advance_filter_Ultimate").save_()
# print(data)

# header = extr_str.get_header
# for k,v in header.items():
#     print(k,":",v)

# print(len(header["Last Name"]))
# print(len(header["Location"]))
# extr_int.save_()
# data = extr.get_data
# for r in data:
#     print(r)
# extr_str.save_()

