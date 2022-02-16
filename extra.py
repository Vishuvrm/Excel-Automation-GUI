from type_extract import TypeExtract

data = TypeExtract("E:\Desktop\Python assignments basic\Python conceptual\Python-SQL\HW\carbon_nanotubes_xl.xlsx", sheet=0).copy_data().get_data

for r in data:
    print(r)