import tkinter as tk
from collapse import Collapse
from functools import partial

class Framework:
    def __init__(self):
        self.root = tk.Tk()
        self.root.minsize(width=self.root.winfo_screenwidth(), height=self.root.winfo_screenheight())
        pw = tk.PanedWindow()
        pw.pack(fill="both", expand=True)

        f1 = tk.Frame(width=200, height=200, background=self._from_rgb((240, 240, 232)))
        f2 = tk.Frame(width=200, height=200, background=self._from_rgb((240, 240, 232)))

        pw.add(f1)
        pw.add(f2)

        # adding some widgets to the left...

        result = tk.Text(f1, height=20, width=125, wrap="none")
        ysb_result = tk.Scrollbar(f1, orient="vertical", command=result.yview)
        xsb_result = tk.Scrollbar(f1, orient="horizontal", command=result.xview)
        result.configure(yscrollcommand=ysb_result.set, xscrollcommand=xsb_result.set)


        f1.grid_rowconfigure(1, weight=1)
        f1.grid_columnconfigure(1, weight=1)

        xsb_result.grid(row=2, column=0, sticky="ew")
        ysb_result.grid(row=1, column=1, sticky="ns")
        result.grid(row=1, column=0, sticky="nsew")

        collapsable1 = Collapse(f1, result)
        collapsable1.load_data()  # 0 x 0

        # result = collapsable1.result
        ###################################################################################################################
        # and to the right...
        # custom_frame = tk.Frame(f2)
        # self.custom_cond = tk.Text(f2, height=16, width=80, wrap="none")
        # ysb_cond = tk.Scrollbar(f2, orient="vertical", command=self.custom_cond.yview)
        # xsb_cond = tk.Scrollbar(f2, orient="horizontal", command=self.custom_cond.xview)
        # self.custom_cond.configure(yscrollcommand=ysb_cond.set, xscrollcommand=xsb_cond.set)
        #
        # f2.grid_rowconfigure(10, weight=1)
        # f2.grid_columnconfigure(10, weight=1)
        #
        # xsb_cond.grid(row=1, column=0, sticky="ew")
        # ysb_cond.grid(row=0, column=1, sticky="ns")
        # self.custom_cond.grid(row=0, column=0, sticky="nsew")
        #
        # b1 = tk.Button(f2, text="Execute as Filter", bg="blue", fg="white", command=collapsable1.custom_filter)
        # b2 = tk.Button(f2, text="Execute as Command", bg="blue", fg="white", command=collapsable1.custom_command)
        # b1.grid(row=3, column=0, sticky="nesw")
        # b2.grid(row=4, column=0, sticky="nesw")

        ################################### Pane window ########################################
        collapsable1.root = f2
        # collapsable1.custom_cond = self.custom_cond.get("1.0", "end")
        # print(self.custom_cond.get("1.0", "end").strip(" ").strip("\n") == "",'-----',self.custom_cond.get("1.0", "end").strip(" "))
        collapsable1.custom_condition()
        collapsable1.copy_data()
        collapsable1.append_data()
        collapsable1.save_data()

        self.root.mainloop()

    def _from_rgb(self, rgb):
        """translates an rgb tuple of int to a tkinter friendly color code
        """
        return "#%02x%02x%02x" % rgb

Framework()


