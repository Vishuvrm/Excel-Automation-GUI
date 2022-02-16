from tkinter import *
from type_extract import TypeExtract
import tkinter as tk

class Interface:
    def __init__(self):
        self.root = Tk()
        self.root.minsize(800, 600)
        self.root.title("Excel Automation")

        self.basic_buildup()
        self.root.mainloop()

    def basic_buildup(self):
        self.frame_left = Frame(self.root)
        self.fram_right = Frame(self.root)


if __name__ == "__main__":
    Interface()

<F>
if type(value) == int:
</F>

<C>
for col in COLUMNS:
    data = []
    for d in COLUMNS[col]:
        if type(d) == int:
            if d % 2 == 0:
                data.append("DivBy2")

            if d % 3 == 0:
                data.append("DivBy3")

            if d % 5 == 0:
                data.append("DivBy5")

            if d % 7 == 0:
                data.append("DivBy7")

            else:
                data.append("NaN")

    COLUMNS[col] = data

</C>

############################################################################
<F>
if type(value) == int:
</F>

<C>

alpha_num = {0:'a', 1:'b', 2:'c', 3:'d', 4:'e', 5:'f'}

for col in COLUMNS:
	data = []
	for d in COLUMNS[col]:
		if type(d) == int:
			x = ""
			for i in str(d):
				if int(i) in alpha_num:
					x += alpha_num[int(i)]
				else:
					x += i
			data.append(x)
	COLUMNS[col] = data

</C>

