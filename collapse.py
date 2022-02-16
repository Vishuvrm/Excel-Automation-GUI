# Importing tkinter and ttk modules
import re
from tkinter import *
from tkinter.ttk import *
from collapsable_pane import CollapsiblePane as cp
from type_extract import TypeExtract
from functools import partial
import tkinter as tk

# Importing Collapsible Pane class that we have
# created in separate file
class Collapse:
	def __init__(self, root, result):
		self.root = root
		self.data = None
		self.custom_cond = None
		self.result = result
		self.load_from_wb = None
		self.sheet_src = None
		self.type_ext_load = None
		self.type_ext_copy = None
		self.type_ext_row_append = None
		self.type_ext_col_append = None
		self.type_ext_row_col_append = None
		self.custom_altered_type_extr = None
		self.custom_altered_copy = None
		self.custom_filter_type_extr = None
		self.custom_filter_copy = None
		self.filter_data = None
		self.altered_data = None

		self.copy_sheet_name = None
		self.append_row_col_sheet = None
		self.copy_type_ = None
		self.copy_sheet_index = None
		self.copy_min_row_src = None
		self.copy_min_col_src = None
		self.copy_max_row_src = None
		self.copy_max_col_src = None
		self.copy_min_row_dst = None
		self.copy_min_col_dst = None
		self.copy_new_wb = None
		self.copy_use_header = None
		self.copy_first_row_header = None
		self.appended_row_col_data = []

	def load_data(self):
		# Creating Object of Collapsible Pane Container
		# If we do not pass these strings in
		# parameter the the defalt strings will appear
		# on button that were, expand >>, collapse <<

		cpane = cp(self.root, 'Load Data', 'Load Data')
		cpane.grid(row=0, column=0, sticky="nesw")

		# Button and checkbutton, these will
		# appear in collapsible pane container
		self.load_from_wb = Entry(cpane.frame, font="Calibri 10", width=20)
		self.load_from_wb.insert(0, "Load from WB")
		self.load_from_wb.grid(row=0, column=0)

		self.sheet_src = Entry(cpane.frame, font="Calibri 10", width=20)
		self.sheet_src.insert(0, "Sheet")
		self.sheet_src.grid(row=0, column=1)

		load_data = Button(cpane.frame, text="Load Data", command=partial(self.load, self.load_from_wb, self.sheet_src))
		load_data.grid(row=1, column=0, sticky="nesw")

	def load(self, load_from_wb, sheet):
		self.custom_cond = None
		self.altered_data = None
		self.custom_altered_copy = None
		self.custom_filter_copy = None
		self.type_ext_col_append = None
		self.type_ext_row_append = None
		self.type_ext_load = None
		self.type_ext_copy = None
		self.appended_row_col_data = []

		self.result.config(state=NORMAL)
		if self.result.get("1.0","end"):
			self.result.delete("1.0", "end")
		load_from_wb = load_from_wb.get()
		sheet = sheet.get()

		if sheet.isnumeric():
			sheet = int(sheet)

		self.type_ext_load = TypeExtract(load_from_wb, sheet)
		self.load_data = self.type_ext_load.copy_data().get_data
		self.data = self.load_data

		for data in self.load_data:
			self.result.insert(END, str(data)+"\n\n")
		self.result.config(state=DISABLED)


	def custom_condition(self):
		condition = Text(self.root, height=16, width=80, wrap="none")
		ysb_cond = Scrollbar(self.root, orient="vertical", command=condition.yview)
		xsb_cond = Scrollbar(self.root, orient="horizontal", command=condition.xview)
		condition.configure(yscrollcommand=ysb_cond.set, xscrollcommand=xsb_cond.set)

		self.root.grid_rowconfigure(10, weight=1)
		self.root.grid_columnconfigure(10, weight=1)

		xsb_cond.grid(row=1, column=0, sticky="ew")
		ysb_cond.grid(row=0, column=1, sticky="ns")
		condition.grid(row=0, column=0, sticky="nsew")

		condition.insert(END, "# Filter F.../F:    value, column, data, showall(Bool), showall = True" + "\n")
		condition.insert(END, "# Transform C.../C: DATA, COLUMNS, SHEET, CELL")
		condition.insert(END, "\n<F>\n\n</F>\n\n")
		condition.insert(END, "<C>\n\n</C>")

		b1 = tk.Button(self.root, text="Execute as Filter", bg="blue", fg="white", command=partial(self.custom_filter, condition))
		b2 = tk.Button(self.root, text="Execute as Command", bg="blue", fg="white", command=partial(self.custom_command, condition))
		b1.grid(row=3, column=0, sticky="nesw")
		b2.grid(row=4, column=0, sticky="nesw")

	def custom_filter(self, condition):
		self.result.config(state=NORMAL)
		if self.result.get("1.0","end"):
			self.result.delete("1.0", "end")

		code = condition.get("1.0","end")

		if re.search(r"<f>.+</f>", code, flags=re.DOTALL):
			code = re.findall(r"<f>.+</f>", code, flags=re.DOTALL)[0].strip("<f>").strip("</f>")
		elif re.search(r"<F>.+</F>", code, flags=re.DOTALL):
			code = re.findall(r"<F>.+</F>", code, flags=re.DOTALL)[0].strip("<F>").strip("</F>")
		else:
			code = None

		if self.type_ext_copy:
			self.custom_filter_type_extr = TypeExtract(self.load_from_wb.get(), self.sheet_src.get())
			self.custom_filter_copy = self.custom_filter_type_extr.copy_data(personal_data = self.filter_data,
																			 sheet=self.copy_sheet_name,
																			 type_=self.copy_type_,
																			 index=self.copy_sheet_index,
																			 min_row_source=self.copy_min_row_src,
																			 min_col_source=self.copy_min_col_src,
																			 max_row_source=self.copy_max_row_src,
																			 max_col_source=self.copy_max_col_src,
																			 min_row_dst=self.copy_min_row_dst,
																			 min_col_dst=self.copy_min_col_dst,
																			 new_wb=self.copy_new_wb,
																			 use_header=self.copy_use_header,
																			 first_row_header=self.copy_first_row_header,
																			 copy_to_different_wb=False,
																			 custom_condition=code)
			self.filter_data = self.custom_filter_copy.get_data
			self.type_ext_copy = None

		elif self.type_ext_load:
			self.custom_filter_copy = TypeExtract(self.load_from_wb.get(), self.sheet_src.get())
			self.filter_data = self.custom_filter_copy.copy_data(custom_condition = code).get_data

		for data in self.filter_data:
			self.result.insert(END, str(data)+"\n\n")
		self.result.config(state=DISABLED)

		self.custom_cond = code

		self.custom_altered_copy = None

	def custom_command(self, condition):
		self.result.config(state=NORMAL)
		if self.result.get("1.0", "end"):
			self.result.delete("1.0", "end")

		code = condition.get("1.0", "end")

		if re.search(r"<c>.+</c>", code, flags=re.DOTALL):
			code = re.findall(r"<c>.+</c>", code, flags=re.DOTALL)[0].strip("<c>").strip("</c>")
		elif re.search(r"<C>.+</C>", code, flags=re.DOTALL):
			code = re.findall(r"<C>.+</C>", code, flags=re.DOTALL)[0].strip("<C>").strip("</C>")
		else:
			code = None
		# self.custom_altered_type_extr = TypeExtract(self.load_from_wb.get(), self.sheet_src.get())

		if self.type_ext_copy:
			if self.filter_data:
				data = self.filter_data
			elif self.data:
				data = self.copied_data
			self.custom_altered_type_extr = TypeExtract(self.load_from_wb.get(), self.sheet_src.get()).copy_data(personal_data=data)
			self.custom_altered_copy = self.custom_altered_type_extr.execute(code)
			self.altered_data = self.custom_altered_copy.get_data
			# self.custom_altered_copy = self.type_ext_copy.execute(code)
			# self.altered_data = self.custom_altered_copy.get_data

		elif self.custom_filter_copy:
			self.custom_altered_copy = self.custom_filter_copy.execute(code)
			self.altered_data = self.custom_filter_copy.get_data
		else:
			# self.type_ext_load:
			self.altered_data = self.custom_altered_type_extr.execute(code).get_data

		for data in self.altered_data:
			self.result.insert(END, str(data) + "\n\n")

		self.result.config(state=DISABLED)
		# self.custom_cond = code

	def copy_data(self):
		# Creating Object of Collapsible Pane Container
		# If we do not pass these strings in
		# parameter the the defalt strings will appear
		# on button that were, expand >>, collapse <<

		cpane = cp(self.root, 'Copy data', 'Copy data')
		cpane.grid(row = 5, column = 0, sticky = "nesw")

		# Button and checkbutton, these will
		# appear in collapsible pane container

		type_ = Entry(cpane.frame, font="Calibri 10", width=20)
		type_.insert(0, "type")
		type_.grid(row=0, column=0)

		sheet_name = Entry(cpane.frame, font="Calibri 10", width=20)
		sheet_name.insert(0, "Sheet")
		sheet_name.grid(row=0, column=1)

		sheet_index = Entry(cpane.frame, font="Calibri 10", width=20)
		sheet_index.insert(0, "Sheet-Index")
		sheet_index.grid(row=0, column=2)

		new_wb = Entry(cpane.frame, font="Calibri 10", width=20)
		new_wb.insert(0, "new-workbook")
		new_wb.grid(row=0, column=3)

		min_row_col_src = Entry(cpane.frame, font="Calibri 10", width=20)
		min_row_col_src.insert(0, "src(min_row x min_col)")
		min_row_col_src.grid(row=2, column=0)

		max_row_col_src = Entry(cpane.frame, font="Calibri 10", width=20)
		max_row_col_src.insert(0, "src(max_row x max_col)")
		max_row_col_src.grid(row=2, column=1)

		min_row_col_dst = Entry(cpane.frame, font="Calibri 10", width=20)
		min_row_col_dst.insert(0, "dst(min_row x min_col)")
		min_row_col_dst.grid(row=3, column=0)

		bool_first_row_header = BooleanVar()
		bool_use_header = BooleanVar()
		bool_first_row_header.set(False)
		bool_use_header.set(False)

		first_row_header = Checkbutton(cpane.frame, text="first-row-header", variable=bool_first_row_header)
		first_row_header.grid(row=4, column=0)
		use_header = Checkbutton(cpane.frame, text="Use-header", variable=bool_use_header)
		use_header.grid(row=4, column=1)

		copy = Button(cpane.frame, text="COPY", command=partial(self.copy, type_, sheet_name, sheet_index, new_wb,
																min_row_col_src, max_row_col_src, min_row_col_dst,
																bool_first_row_header, bool_use_header))
		copy.grid(row=5, column=1, sticky = "nesw")

	def copy(self, type_, sheet_name, sheet_index, new_wb, min_row_col_src, max_row_col_src, min_row_col_dst, first_row_header, use_header):

		self.result.config(state=NORMAL)
		if self.result.get("1.0","end"):
			self.result.delete("1.0", "end")

		if not sheet_name.get():
			self.copy_sheet_name = "Sheet"
		else:
			self.copy_sheet_name = sheet_name.get()

		if not type_.get() or type_.get() == "type":
			self.copy_type_ = None
		elif type_.get() == "str":
			self.copy_type_ = str
		elif type_.get() == "int":
			self.copy_type_ = int

		if not sheet_index.get() or sheet_index.get() == "Sheet-Index":
			self.copy_sheet_index = 0
		else:
			self.copy_sheet_index = int(sheet_index.get())

		min_row_col_src = min_row_col_src.get().replace(' ', '')
		max_row_col_src = max_row_col_src.get().replace(' ', '')
		min_row_col_dst = min_row_col_dst.get().replace(' ', '')

		if re.search(r"[\d]+x[\d]+", min_row_col_src):
			self.copy_min_row_src, self.copy_min_col_src = [int(a) for a in min_row_col_src.split('x')]
		else:
			self.copy_min_row_src, self.copy_min_col_src = None, None

		if re.search(r"[\d]+x[\d]+", max_row_col_src):
			self.copy_max_row_src, self.copy_max_col_src = [int(a) for a in max_row_col_src.split('x')]
		else:
			self.copy_max_row_src, self.copy_max_col_src = None, None

		if re.search(r"[\d]+x[\d]+", min_row_col_dst):
			self.copy_min_row_dst, self.copy_min_col_dst = [int(a) for a in min_row_col_dst.split('x')]
		else:
			self.copy_min_row_dst, self.copy_min_col_dst = 1, 1

		if not new_wb.get() or new_wb.get() == "new-workbook":
			self.copy_new_wb = None
		else:
			self.copy_new_wb = self.copy_new_wb.get()

		self.copy_first_row_header = first_row_header.get()
		self.copy_use_header = use_header.get()

		self.type_ext_copy = TypeExtract(self.load_from_wb.get(),
										 self.sheet_src.get()).copy_data(personal_data=self.altered_data,
																		type_=self.copy_type_,
																		custom_condition=self.custom_cond,
																		min_row_source=self.copy_min_row_src,
																		min_col_source=self.copy_min_col_src,
																		max_row_source=self.copy_max_row_src,
																		max_col_source=self.copy_max_col_src,
																		use_header=self.copy_use_header,
																		first_row_header=self.copy_first_row_header,
																		new_wb=self.copy_new_wb,
																		copy_to_different_wb=False,
																		sheet=self.copy_sheet_name,
																		index=self.copy_sheet_index,
																		min_row_dst=self.copy_min_row_dst,
																		min_col_dst=self.copy_min_col_dst)

		self.copied_data = self.type_ext_copy.get_data

		self.data = self.copied_data

		for data in self.data:
			self.result.insert(END, str(data) + "\n\n")
		self.result.config(state=DISABLED)

		self.custom_filter_type_extr = None
		self.custom_altered_type_extr = None

	def append_data(self):
		# Creating Object of Collapsible Pane Container
		# If we do not pass these strings in
		# parameter the the defalt strings will appear
		# on button that were, expand >>, collapse <<
		cpane = cp(self.root, 'Append data', 'Append data')
		cpane.grid(row=6, column=0, sticky = "nesw")

		# Button and checkbutton, these will
		# appear in collapsible pane container
		wb_name = Entry(cpane.frame, font="Calibri 10", width=20)
		wb_name.insert(0, "Workbook-name")
		wb_name.grid(row=0, column=0)

		sheet_name = Entry(cpane.frame, font="Calibri 10", width=20)
		sheet_name.insert(0, "Sheet-name")
		sheet_name.grid(row=1, column=0)

		bool_append_to_rows = BooleanVar()
		bool_append_to_rows.set(False)
		append_rows = Checkbutton(cpane.frame, text="Append-to-rows", variable=bool_append_to_rows)
		append_rows.grid(row=4, column=0, )

		bool_append_to_cols = BooleanVar()
		bool_append_to_cols.set(False)
		append_cols = Checkbutton(cpane.frame, text="Append-to-cols", variable=bool_append_to_cols)
		append_cols.grid(row=5, column=0)

		append = Button(cpane.frame, text="APPEND", command=partial(self.append_row_col, wb_name, sheet_name, bool_append_to_rows, bool_append_to_cols))
		append.grid(row=6, column=0, sticky = "nesw")

	def append_row_col(self, wb, sheet_name, append_rows, append_cols):
		self.result.config(state=NORMAL)
		if self.result.get("1.0", "end"):
			self.result.delete("1.0", "end")

		self.append_row_col_sheet = sheet_name.get()

		if not wb.get() or wb.get()=="Workbook-name":
			wb = None
		else:
			wb = wb.get()

		if not sheet_name.get() or sheet_name.get()=="Sheet-name":
			sheet = None
		else:
			sheet = sheet_name.get()
		append_to_rows = append_rows.get()
		append_to_cols = append_cols.get()

		data = []
		if self.custom_altered_copy:
			data = self.altered_data

		elif self.custom_filter_copy:
			data = self.filter_data

		if self.type_ext_copy:
			type_ext_append = self.type_ext_copy
			if not data:
				data = type_ext_append.get_data

		elif self.type_ext_load:
			type_ext_append = self.type_ext_load
			if not data:
				data = type_ext_append.get_data

		if append_to_rows:
			if not self.type_ext_row_append:
				self.type_ext_row_append = type_ext_append.append_to_rows(data[1:], wb, sheet)
			else:
				self.type_ext_row_append.append_to_rows(data[1:], wb, sheet)

			min_row = self.type_ext_row_append.final_sheet.min_row
			max_row = self.type_ext_row_append.final_sheet.max_row
			min_col = self.type_ext_row_append.final_sheet.min_column
			max_col = self.type_ext_row_append.final_sheet.max_column

			for row in self.type_ext_row_append.final_sheet.iter_rows(min_col=min_col, max_col=max_col, min_row=min_row, max_row=max_row, values_only=True):
				self.result.insert(END, str(row) + "\n\n")

		elif append_to_cols:
			if not self.type_ext_col_append:
				self.type_ext_col_append = type_ext_append.append_to_cols(data, wb, sheet)
			else:
				self.type_ext_col_append.append_to_cols(data, wb, sheet)

			min_row = self.type_ext_col_append.final_sheet.min_row
			max_row = self.type_ext_col_append.final_sheet.max_row
			min_col = self.type_ext_col_append.final_sheet.min_column
			max_col = self.type_ext_col_append.final_sheet.max_column

			for row in self.type_ext_col_append.final_sheet.iter_rows(min_col=min_col, max_col=max_col, min_row=min_row,
													   max_row=max_row, values_only=True):
				self.result.insert(END, str(row) + "\n\n")

		else:
			raise Exception("Data not properly mentioned!")

		self.result.config(state=DISABLED)

	def save_data(self):
		# Creating Object of Collapsible Pane Container
		# If we do not pass these strings in
		# parameter the the defalt strings will appear
		# on button that were, expand >>, collapse <<
		cpane = cp(self.root, 'Save', 'Save')
		cpane.grid(row=7, column=0, sticky="nesw")

		# Button and checkbutton, these will
		# appear in collapsible pane container

		save_to_new_wb = BooleanVar()
		save_to_new_wb.set(False)
		Checkbutton(cpane.frame, text="Save to a new WB", variable=save_to_new_wb).grid(row=0, column=0)

		to_wb = Entry(cpane.frame, font="Calibri 10", width=20)
		to_wb.insert(0, "WB-complete-name")
		to_wb.grid(row=1, column=0)

		save = Button(cpane.frame, text="SAVE" ,command=partial(self.save, save_to_new_wb, to_wb))
		save.grid(row=2, column=0, sticky = "nesw")

	def save(self, save_to_new_wb, new_wb):
		if not new_wb.get() or new_wb.get() == "WB-complete-name":
			new_wb = self.load_from_wb
		else:
			new_wb = new_wb.get()
		sheet = self.copy_sheet_name

		if self.type_ext_load and not(self.custom_filter_copy) and not(self.custom_altered_copy) and not(self.type_ext_copy) and (not(self.type_ext_row_append) or not(self.type_ext_col_append)):
			print("Load ==> Save")
			self.type_ext_load.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and not(self.custom_altered_copy) and not(self.type_ext_copy) and not(not(self.type_ext_row_append) or not(self.type_ext_col_append)):
			print("Load data ==> Execute as filter ==> Save")
			self.custom_filter_copy.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and self.custom_altered_copy and not(self.type_ext_copy) and not(not(self.type_ext_row_append) or not(self.type_ext_col_append)):
			print("Load data ==> filter ==>Command ==> Save")
			self.custom_altered_copy.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and self.type_ext_copy and not(self.custom_altered_copy) and not(not(self.type_ext_row_append) or not(self.type_ext_col_append)):
			print("Load data ==> filter ==> copy ==> Save")
			self.type_ext_copy.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and self.custom_altered_copy and self.type_ext_copy and not(not(self.type_ext_row_append) or not(self.type_ext_col_append)):
			print("Load data ==> filter ==> command ==> copy ==> Save")
			self.type_ext_copy.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and (self.type_ext_row_append or self.type_ext_col_append) and not(self.custom_altered_copy) and not(self.type_ext_copy):
			print("Load data ==> filter ==> append ==> Save")
			if self.append_row_col_sheet:
				sheet = self.append_row_col_sheet
			if self.type_ext_row_append:
				self.type_ext_row_append.save_(save_to_new_wb.get(), new_wb, sheet)
			elif self.type_ext_col_append:
				self.type_ext_col_append.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and (self.type_ext_row_append or self.type_ext_col_append) and self.type_ext_copy and not(self.custom_altered_copy):
			print("Load data ==> filter ==> append ==> copy ==> Save")
			self.type_ext_copy.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and self.custom_altered_copy and (self.type_ext_row_append or self.type_ext_col_append) and not(self.type_ext_copy):
			print("Load data ==> filter ==> command ==> append ==> Save")
			if self.append_row_col_sheet:
				sheet = self.append_row_col_sheet
			if self.type_ext_row_append:
				self.type_ext_row_append.save_(save_to_new_wb.get(), new_wb, sheet)
			elif self.type_ext_col_append:
				self.type_ext_col_append.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and self.custom_filter_copy and self.custom_altered_copy and (self.type_ext_row_append or self.type_ext_col_append) and self.type_ext_copy:
			print("Load data ==> filter ==> command ==> append ==> copy ==> save")
			self.type_ext_copy.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and (self.type_ext_row_append or self.type_ext_col_append) and not(self.custom_filter_copy) and not(self.custom_altered_copy) and not(self.type_ext_copy):
			print("Load data ==> append ==> Save")
			if self.append_row_col_sheet:
				sheet = self.append_row_col_sheet
			if self.type_ext_row_append:
				self.type_ext_row_append.save_(save_to_new_wb.get(), new_wb, sheet)
			elif self.type_ext_col_append:
				self.type_ext_col_append.save_(save_to_new_wb.get(), new_wb, sheet)

		elif self.type_ext_load and (self.type_ext_row_append or self.type_ext_col_append) and self.type_ext_copy and not(self.custom_filter_copy) and not(self.custom_altered_copy):
			print("Load data ==> append ==> copy ==> Save")
			self.type_ext_copy.save_(save_to_new_wb.get(), new_wb, sheet)

		else:
			print("Condition not satisfied!")



		# if self.type_ext_row_col_append:
		# 	self.type_ext_row_col_append.save_(save_to_new_wb.get(), new_wb, sheet)
		#
		# if self.type_ext_copy:
		# 	self.type_ext_copy.save_(save_to_new_wb.get(), new_wb, sheet)
		#
		# if self.custom_altered_type_extr:
		# 	self.custom_altered_type_extr.save_(save_to_new_wb.get(), new_wb, sheet)
		#
		# if self.custom_altered_copy:
		# 	self.custom_altered_copy.save_(save_to_new_wb.get(), new_wb, sheet)
		#
		# if self.custom_filter_type_extr:
		# 	self.custom_filter_type_extr.save_(save_to_new_wb.get(), new_wb, sheet)
		#
		# if self.custom_filter_copy:
		# 	self.custom_filter_copy.save_(save_to_new_wb.get(), new_wb, sheet)



