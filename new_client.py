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
            for col_no in range(1, ncols):
                cols = table.col_values(col_no)
                for col in cols:
                    if var in str(col).replace(' ', ''):
                        for cell_no in range(2, len(cols)):
                            if cols[cell_no] not in voltage_list:
                                voltage_list.append(cols[cell_no])
        return voltage_list

    def get_price(self, var='铜价'):
        price_list = []
        sheet_no = self.open().nsheets
        for sheet in range(0, sheet_no):
            table = self.open().sheet_by_index(sheet)
            nrows = table.nrows  # 行数
            ncols = table.ncols  # 列数
            for row_no in range(1, nrows):
                rows = table.row_values(row_no)
                for row in rows:
                    if var in str(row).replace(' ', ''):
                        for cell_no in range(1, len(rows)):
                            if rows[cell_no] not in price_list:
                                price_list.append(rows[cell_no])
        return price_list


if __name__ == '__main__':
    xls = Xls()
    print(xls.get_price(var='铜价'))
