from tkinter import *
from new_client import *


def set_file_name():
    global xls
    global voltage, row_dict, price, multi_selection
    voltage = None
    row_dict = None
    price = None
    multi_selection = None
    voltage_list_box.delete(0, END)
    model_list_box.delete(0, END)
    price_list_box.delete(0, END)
    spe_list_box.delete(0, END)
    excel_file_name = file_name.get()
    xls = Xls(excel_file_name)
    voltage_list = xls.get_voltage()
    for item in voltage_list:
        voltage_list_box.insert(END, item)
    voltage_list_box.insert(0, '无')

    model_list = xls.get_voltage(var='型号')
    for item in model_list:
        model_list_box.insert(END, item)

    price_list = xls.get_price()
    for item in price_list:
        price_list_box.insert(END, item)
    return xls


def set_multi_selections():
    try:
        global row_dict
        multi_list_box.delete(0, END)
        if voltage and voltage != '无':
            multi_selections_list, row_dict = xls.get_row(model, specification, voltage)
            if multi_selections_list:
                for item in multi_selections_list:
                    multi_list_box.insert(END, item)
        elif voltage == '无':
            multi_selections_list, row_dict = xls.get_row(model, specification)
            if multi_selections_list:
                for item in multi_selections_list:
                    multi_list_box.insert(END, item)
        else:
            result_txt.insert(END, '{}\n'.format('请选择电压.'))
    except TypeError:
        result_txt.insert(END, '{}\n'.format('该组合无效!请使用Excel中正确的"电压","型号","规格"组合.'))
    except NameError:
        result_txt.insert(END, '{}\n'.format('该组合无效!请使用Excel中正确的"电压","型号","规格"组合.'))
    except Exception as e:
        result_txt.insert(END, '{}\n'.format(e))


def get_voltage(event):
    global voltage
    voltage = voltage_list_box.get(voltage_list_box.curselection())
    spe_list_box.select_clear(0, END)
    # result_txt.insert(END, '{}\n'.format(voltage))


def get_model(event):
    global model
    model = model_list_box.get(model_list_box.curselection())
    spe_list = xls.get_spe(model)
    spe_list_box.delete(0, END)
    for item in spe_list:
        spe_list_box.insert(END, item)
    # result_txt.insert(END, '{}\n'.format(model))


def get_specification(event):
    global specification
    specification = spe_list_box.get(spe_list_box.curselection())
    set_multi_selections()


def get_price(event):
    global price
    price = price_list_box.get(price_list_box.curselection())
    # result_txt.insert(END, '{}\n'.format(price))


def get_multi_selections(event):
    global multi_selection
    multi_selection_list = []
    select_tuple = multi_list_box.curselection()
    for i in select_tuple:
        multi_selection_list.append(multi_list_box.get(i))
    multi_selection = tuple(multi_selection_list)
    # result_txt.insert(END, '{}\n'.format(multi_selections))


def get_unit_price():
    try:
        if row_dict and multi_selection and price:
            unit_price = xls.get_unit_price(row_dict, price, multi_selection)
            result_txt.insert(END, '计算结果{}万元/吨\n'.format(unit_price))
            result_txt.see(END)  # 一直显示最新的一行
            result_txt.update()
        elif row_dict and price and not multi_selection:
            unit_price = xls.get_unit_price(row_dict, price)
            result_txt.insert(END, '计算结果{}万元/吨\n'.format(unit_price))
            result_txt.see(END)  # 一直显示最新的一行
            result_txt.update()
        elif not row_dict:
            result_txt.insert(END, '{}\n'.format('该组合无效!请使用Excel中正确的"电压","型号","规格"组合.'))
            result_txt.see(END)  # 一直显示最新的一行
            result_txt.update()
        elif not price:
            result_txt.insert(END, '{}\n'.format('请从铜价列表中选择铜价.'))
            result_txt.see(END)  # 一直显示最新的一行
            result_txt.update()
        else:
            result_txt.insert(END, '{}\n'.format('该组合无效!请使用Excel中正确的"电压","型号","规格"组合.'))
            result_txt.see(END)  # 一直显示最新的一行
            result_txt.update()
    except NameError:
        result_txt.insert(END, '{}\n'.format('请先加载Excel文件.'))
        result_txt.see(END)  # 一直显示最新的一行
        result_txt.update()
    except Exception as e:
        result_txt.insert(END, '{}\n'.format(e))
        result_txt.see(END)  # 一直显示最新的一行
        result_txt.update()
    # result_txt.insert(END, '{}\n'.format(unit_price))


def clear_result():
    result_txt.delete('1.0', 'end')


root = Tk()
root.title("Excel tool")
root.geometry("1000x500")
root.resizable(width=False, height=False)

# 显示电压列表
Label(root, text="电压列表", font=('Arial', 10)).grid(row=0, column=0)
voltage_list_display = StringVar()
voltage_list_box = Listbox(root, listvariable=voltage_list_display, exportselection=False)
voltage_list_box.bind('<ButtonRelease-1>', get_voltage)
voltage_list_box.grid(row=1, column=0, sticky=E)

# 显示电压滚动条
scrl_voltage = Scrollbar(root)
scrl_voltage.grid(row=1, column=1, sticky=W + N + S)
voltage_list_box.configure(yscrollcommand=scrl_voltage.set)
scrl_voltage['command'] = voltage_list_box.yview

# 显示型号列表
Label(root, text="型号列表", font=('Arial', 10)).grid(row=0, column=2)
model_list_display = StringVar()
model_list_box = Listbox(root, listvariable=model_list_display, exportselection=False)
model_list_box.bind('<ButtonRelease-1>', get_model)
model_list_box.grid(row=1, column=2, sticky=E)

# 显示型号滚动条
scrl_model = Scrollbar(root)
scrl_model.grid(row=1, column=3, sticky=W + N + S)
model_list_box.configure(yscrollcommand=scrl_model.set)
scrl_model['command'] = model_list_box.yview

# 显示规格列表
Label(root, text="规格列表", font=('Arial', 10)).grid(row=0, column=4)
spe_list_display = StringVar()
spe_list_box = Listbox(root, listvariable=spe_list_display, exportselection=False)
spe_list_box.bind('<ButtonRelease-1>', get_specification)
spe_list_box.grid(row=1, column=4, sticky=E)

# 显示规格滚动条
scrl_spe = Scrollbar(root)
scrl_spe.grid(row=1, column=5, sticky=W + N + S)
spe_list_box.configure(yscrollcommand=scrl_spe.set)
scrl_spe['command'] = spe_list_box.yview

# 显示铜价列表
Label(root, text="铜价列表", font=('Arial', 10)).grid(row=0, column=6)
price_list_display = StringVar()
price_list_box = Listbox(root, listvariable=price_list_display, exportselection=False)
price_list_box.bind('<ButtonRelease-1>', get_price)
price_list_box.grid(row=1, column=6, sticky=E)

# 显示铜价滚动条
scrl_price = Scrollbar(root)
scrl_price.grid(row=1, column=7, sticky=W + N + S)
price_list_box.configure(yscrollcommand=scrl_price.set)
scrl_price['command'] = price_list_box.yview

# 显示附加列表
Label(root, text="附加选项列表", font=('Arial', 10)).grid(row=2, column=0)
multi_list_display = StringVar()
multi_list_box = Listbox(root, listvariable=multi_list_display, selectmode=MULTIPLE)
multi_list_box.bind('<ButtonRelease-1>', get_multi_selections)
multi_list_box.grid(row=3, column=0, sticky=E)

# 显示附加列表滚动条
scrl_multi = Scrollbar(root)
scrl_multi.grid(row=3, column=1, sticky=E + N + S)
multi_list_box.configure(yscrollcommand=scrl_multi.set)
scrl_multi['command'] = multi_list_box.yview

# 显示附加列表横向滚动条
scrl_multi = Scrollbar(root, orient=HORIZONTAL)
scrl_multi.grid(row=4, column=0, sticky=E + W + N)
multi_list_box.configure(xscrollcommand=scrl_multi.set)
scrl_multi['command'] = multi_list_box.xview

# 显示获取附加项按钮
# Button(root, text="获取附加选项", command=set_multi_selections).grid(row=5, column=0)

# 显示文件名输入框
file_name = StringVar()
file_name_entry = Entry(root, textvariable=file_name)
file_name.set("test.xls")

file_name_entry.grid(row=1, column=8, sticky=N)

# 显示加载按钮
Button(root, text="Load", command=set_file_name).grid(row=1, column=9, sticky=N)

# 显示结果显示框
Label(root, text="计算结果", font=('Arial', 10)).grid(row=2, column=2, columnspan=7, sticky=W)
result_txt = Text(root, height=15)
result_txt.grid(row=3, column=2, columnspan=6, rowspan=2, sticky=W + E + N + S)

# 显示结果滚动条
scrl_result = Scrollbar(root)
scrl_result.grid(row=3, column=7, rowspan=2, sticky=E + S + N)
result_txt.configure(yscrollcommand=scrl_result.set)
scrl_result['command'] = result_txt.yview

# 显示获取单价
Button(root, text="获取单价", command=get_unit_price).grid(row=3, column=8, sticky=W + N)

# 显示清除结果按钮
Button(root, text="清除结果", command=clear_result).grid(row=5, column=7, sticky=W + N)

# 显示自定义铜价
Label(root, text="自定义铜价", font=('Arial', 10)).grid(row=4, column=8, sticky=W)

# 显示自定义铜价输入框
customer_copper_price = StringVar()
customer_copper_price_entry = Entry(root, textvariable=customer_copper_price)
# customer_copper_price.set("")

customer_copper_price_entry.grid(row=5, column=8, sticky=W)

root.grid()
root.mainloop()
