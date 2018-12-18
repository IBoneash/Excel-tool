import xlrd
import re
import time
import logging.handlers

# Log File sort by time
tm = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
log_file = str(tm) + '.log'

# Log File handler
handler = logging.handlers.RotatingFileHandler(log_file, maxBytes=1024 * 1024, backupCount=5)
fmt = '%(asctime)s - %(message)s'
formatter = logging.Formatter(fmt)
handler.setFormatter(formatter)

log = logging.getLogger('test')
log.addHandler(handler)
log.setLevel(logging.DEBUG)


class Xls(object):
    def __init__(self, __name='test.xls'):
        self.xls_name = __name
        self.xlrd_object = None

    def open(self):
        try:
            self.xlrd_object = xlrd.open_workbook(self.xls_name)
            return self.xlrd_object
        except Exception as e:
            log.info('Open {} failed, {}\n'.format(self.xls_name, e))

    def get_voltage(self, var='电压'):
        voltage_list = []
        sheet_no = self.open().nsheets
        for sheet in range(0, sheet_no):
            table = self.open().sheet_by_index(sheet)
            nrows = table.nrows  # 行数
            ncols = table.ncols  # 列数
            for col_no in range(0, ncols):
                cols = table.col_values(col_no)
                for col in cols:
                    if var in str(col).replace(' ', ''):
                        for cell_no in range(0, len(cols)):
                            if cols[cell_no] and cols[cell_no] not in voltage_list and str(cols[cell_no]).replace(' ',
                                                                                                                  '') != var:
                                voltage_list.append(cols[cell_no])
        log.debug("获取{}列表为:{}".format(var, voltage_list))
        return voltage_list

    def get_price(self, var='铜价'):
        price_list = []
        sheet_no = self.open().nsheets
        for sheet in range(0, sheet_no):
            table = self.open().sheet_by_index(sheet)
            nrows = table.nrows  # 行数
            ncols = table.ncols  # 列数
            for row_no in range(0, nrows):
                rows = table.row_values(row_no)
                for row in rows:
                    if var in str(row).replace(' ', ''):
                        for cell_no in range(0, len(rows)):
                            if var in rows[cell_no].replace(" ", "") and rows[cell_no] not in price_list:
                                price_list.append(rows[cell_no])
        log.debug("获取{}列表为:{}".format(var, price_list))
        return price_list

    def get_row(self, model, specification, voltage=None):
        row_key = []
        row_value = []
        sheet_no = self.open().nsheets
        for sheet in range(0, sheet_no):
            multi_selections = []
            table = self.open().sheet_by_index(sheet)
            nrows = table.nrows  # 行数
            ncols = table.ncols  # 列数
            for row_no in range(0, nrows):
                row_0 = []
                rows = table.row_values(row_no)
                for item in rows:
                    row_0.append(str(item).replace(" ", ""))
                if '型号' in row_0 and '规格' in row_0:
                    row_key = row_0
                    for select in row_0:
                        if select and '执行价格' not in select and '序号' not in select and '供应商' not in select and '型号' not in select and '电压' not in select and '规格' not in select and '铜价' not in select:
                            multi_selections.append(select)
                if voltage:
                    if model in rows and voltage in rows and specification in rows:
                        row_value = rows
                        row_dict = dict(zip(row_key, row_value))
                        log.debug(
                            "获取model={},voltage={},specification={} ,整行数据为:{},复选项为{}".format(model, voltage,
                                                                                             specification,
                                                                                             rows, multi_selections))
                        return multi_selections, row_dict
                elif model in rows and specification in rows and '电压' not in row_key:
                    row_value = rows
                    row_dict = dict(zip(row_key, row_value))
                    log.debug(
                        "没有电压,获取model={},specification={} ,整行数据为:{},复选项为{}".format(model, specification, rows,
                                                                                   multi_selections))
                    return multi_selections, row_dict
                else:
                    log.debug('没有该组合数据!')

    def get_unit_price(self, row_dict, copper_price, multi_selections=None):
        unit_price = row_dict[copper_price]
        log.info('{} 对应单价为 {}'.format(copper_price, unit_price))
        if multi_selections:
            for i in multi_selections:
                log.info('{} 对应单价为 {}'.format(i, row_dict[i]))
                unit_price += row_dict[i]
            log.info('最后统计总单价为 {}'.format(unit_price))
        return '{:.4f}'.format(unit_price)

    def get_spe(self, model):
        spe_key = []
        spe_value = []
        spe_list = []
        sheet_no = self.open().nsheets
        for sheet in range(0, sheet_no):
            table = self.open().sheet_by_index(sheet)
            nrows = table.nrows  # 行数
            ncols = table.ncols  # 列数
            for col_no in range(0, ncols):
                cols = table.col_values(col_no)
                for col in cols:
                    if "型号" == str(col).replace(' ', ''):
                        for cell_no in range(0, len(cols)):
                            if cell_no:
                                spe_key.append(cols[cell_no])
                    if "规格" == str(col).replace(' ', ''):
                        for cell_no in range(0, len(cols)):
                            if cell_no:
                                spe_value.append(cols[cell_no])
        for i in range(0, len(spe_key)):
            if spe_key[i] == model and spe_value not in spe_list:
                spe_list.append(spe_value[i])
        log.debug("获取型号,规格对应列表为:{}".format(spe_list))
        return spe_list

    def get_avg_price(self, custom_price, row_dict):
        price_dict = []
        price_list = custom_price.split(',')
        if price_list and len(price_list) == 1:
            return cal_avg_price(price_list[0], row_dict)
        elif price_list and len(price_list) == 2:
            if float(price_list[0]) < float(price_list[1]):
                for item in range(int(float(price_list[0]) * 10), int(float(price_list[1]) * 10 + 1)):
                    if item % 10 == 0:
                        item = int(item / 10)
                    else:
                        item = item / 10
                    price_dict.append("{}".format(cal_avg_price(item, row_dict)[0]))
                return price_dict


def cal_avg_price(price, row_dict, multi_selections=None):
    price_dict = []
    if 0 < float(price) % 1 < 0.5:
        price_low = "铜价{}万元/吨".format(int(float(price)))
        price_high = "铜价{:.1f}万元/吨".format(int(float(price)) + 0.5)
        price_avg = (row_dict[price_high] - row_dict[price_low]) * (float(price) - int(float(price))) / 5 + row_dict[
            price_low]
        price = "铜价{:.1f}万元/吨".format(price)
    elif float(price) % 1 > 0.5:
        price_low = "铜价{:.1f}万元/吨".format(int(float(price)) + 0.5)
        price_high = "铜价{}万元/吨".format(int(float(price)) + 1)
        price_avg = (row_dict[price_high] - row_dict[price_low]) * (int(float(price)) + 1 - float(price)) / 5 + \
                    row_dict[price_low]
        price = "铜价{:.1f}万元/吨".format(price)
    elif float(price) % 1 == 0:
        price_str = float(price)
        price = "铜价{}万元/吨".format(int(price))
        price_avg = row_dict[price]
        price = "铜价{:.1f}万元/吨".format(int(price_str))
    else:
        price = "铜价{:.1f}万元/吨".format(price)
        price_avg = row_dict[price]
    price_dict.append('[{}]:{:.4f}元/米'.format(price, price_avg))
    return price_dict


# if __name__ == '__main__':
#     xls = Xls()
#     dic = {'序号': 2.0, '框架协议供应商': '保定京阳立津线缆制造有限公司', '型号': 'KYJV', '规格': '10×1.5', '铜价2.5万元/吨': 6.2216,
#            '铜价3万元/吨': 7.1628, '铜价3.5万元/吨': 8.0148, '铜价4万元/吨': 8.8271, '铜价4.5万元/吨': 9.6692, '铜价5万元/吨': 10.4519,
#            '铜价5.5万元/吨': 11.1057, '铜价6万元/吨': 11.829, '铜价6.5万元/吨': 12.5423, '铜价7.0万元/吨': 13.3051, '铜价7.5万元/吨': 14.0679,
#            '耐火加价\n(元/米）\n（FF型填0）': 1.7535, 'A级阻燃加价\n(元/米）': 0.4656, 'B级阻燃加价\n(元/米）': 0.3269, 'C级阻燃加价\n(元/米）': 0.1783,
#            '低烟无卤加价\n(元/米）（FF型填0）': 0.6638, '软线R加价\n(元/米)': 0.5647, '铠装22加价\n(元/米）': 1.0303, '铠装23加价\n(元/米）': 1.07,
#            '铠装32加价\n(元/米）': 2.447, '铠装33加价\n(元/米）': 4.0817, '防白蚁加价（元/米）': 0.4954,
#            '耐低温（－40℃）聚氯乙烯护套加价(元/米）（GG型、FF型填0）': 0.5152, '铜带屏蔽P2减价（元/米）（无屏蔽填0）': 0.0, '铝塑屏蔽P3减价（元/米）（无屏蔽填0）': 0.0}
#     #     # print(xls.get_voltage(var='型号'))
#     #     # print(xls.get_row('KYJV', '10×1.5'))
#     #     print(xls.get_unit_price(dict, '铜价4.5万元/吨', '耐低温（－40℃）聚氯乙烯护套加价(元/米）（GG型、FF型填0）'))
#     # xls.get_spe('KGG')
#     print(xls.get_avg_price('4', dic))
