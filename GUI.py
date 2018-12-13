from tkinter import *
from new_client import *


def set_file_name():
    excel_file_name = file_name.get()
    xls = Xls(excel_file_name)
    voltage_list = xls.get_voltage()
    for item in voltage_list:
        voltage_list_box.insert(END, item)

    model_list = xls.get_voltage(var='型号')
    for item in model_list:
        model_list_box.insert(END, item)

    spe_list = xls.get_voltage(var='规格')
    for item in spe_list:
        spe_list_box.insert(END, item)


def get_voltage(event):
    return voltage_list_box.get(voltage_list_box.curselection())


def get_model(event):
    return model_list_box.get(model_list_box.curselection())


def get_specification(event):
    return spe_list_box.get(spe_list_box.curselection())


root = Tk()
root.title("Excel tool")
root.geometry("700x500")
root.resizable(width=True, height=True)

# 显示电压列表
Label(root, text="电压列表", font=('Arial', 10)).grid(row=0, column=0)
voltage_list_display = StringVar()
voltage_list_box = Listbox(root, listvariable=voltage_list_display)
voltage_list_box.bind('<ButtonRelease-1>', get_voltage)
voltage_list_box.grid(row=1, column=0)

# 显示型号列表
Label(root, text="型号列表", font=('Arial', 10)).grid(row=0, column=1)
model_list_display = StringVar()
model_list_box = Listbox(root, listvariable=model_list_display)
model_list_box.bind('<ButtonRelease-1>', get_model)
model_list_box.grid(row=1, column=1)

# 显示滚动条
scrl_model = Scrollbar(root)
scrl_model.grid(row=1, column=2, sticky=N + S)
model_list_box.configure(yscrollcommand=scrl_model.set)
scrl_model['command'] = model_list_box.yview

# 显示规格列表
Label(root, text="规格列表", font=('Arial', 10)).grid(row=0, column=3)
spe_list_display = StringVar()
spe_list_box = Listbox(root, listvariable=spe_list_display)
spe_list_box.bind('<ButtonRelease-1>', get_specification)
spe_list_box.grid(row=1, column=3)

# 显示滚动条
scrl_spe = Scrollbar(root)
scrl_spe.grid(row=1, column=4, sticky=N + S)
spe_list_box.configure(yscrollcommand=scrl_spe.set)
scrl_spe['command'] = spe_list_box.yview

# 显示铜价列表
Label(root, text="铜价列表", font=('Arial', 10)).grid(row=0, column=3)
spe_list_display = StringVar()
spe_list_box = Listbox(root, listvariable=spe_list_display)
spe_list_box.bind('<ButtonRelease-1>', get_specification)
spe_list_box.grid(row=1, column=3)

# 显示滚动条
scrl_spe = Scrollbar(root)
scrl_spe.grid(row=1, column=4, sticky=N + S)
spe_list_box.configure(yscrollcommand=scrl_spe.set)
scrl_spe['command'] = spe_list_box.yview

# 显示文件名输入框
file_name = StringVar()
file_name_entry = Entry(root, textvariable=file_name)
file_name.set("test.xls")

file_name_entry.grid(row=1, column=5)

# 显示加载按钮
Button(root, text="Load", command=set_file_name).grid(row=1, column=6)

root.grid()
root.mainloop()
