import xlrd
import re
import time
import logging.handlers

# Log File sort by time
tm = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
log_file = str(tm)+'.log'

# Log File handler
handler = logging.handlers.RotatingFileHandler(log_file, maxBytes=1024 * 1024, backupCount=5)
fmt = '%(asctime)s - %(message)s'
formatter = logging.Formatter(fmt)
handler.setFormatter(formatter)

logger = logging.getLogger('test')
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)

# Load Excel File
file_name = 'test.xls'
ob = xlrd.open_workbook(file_name)

# class input
class Input(object):
    def __init__(self, xh='', dy='', gg='', ll='', tj='', tjv=0):
        self.xh = xh
        self.dy = dy
        self.gg = gg
        self.ll = ll
        self.tj = tj
        self.tjv = tjv

    # calculate the gap of copper price
    def get_gap(self):
        self.__gap = self.tjv % 0.5
        return self.__gap

    # get the price cell of excel
    def get_price(self):
        self.__pricez = self.tjv - self.tjv % 0.5
        if self.__pricez % 1 == 0:
            self.__pricez = int(self.__pricez)
        self.__prices_after = u'\u94dc\u4ef7' + str(self.__pricez) + u'\u4e07\u5143/\u5428'
        return self.__prices_after

    # get the price cell +1 of excel
    def get_price_plus(self):
        self.__price_plus = self.tjv - self.tjv % 0.5 + 0.5
        if self.__price_plus % 1 == 0:
            self.__price_plus = int(self.__price_plus)
        self.__price_plus_after = u'\u94dc\u4ef7' + str(self.__price_plus) + u'\u4e07\u5143/\u5428'
        return self.__price_plus_after

# get int number
def get_int(num):
    if num % 1 == 0:
        num = str(int(num))
    else:
        num = str(num)
    return num


if __name__ == '__main__':
    try:
        while True:
            data = Input()
            data.xh = raw_input('Input xh: ').upper()
            logger.info('Input xh: %s' % (data.xh))
            data.dy = raw_input('Input dy: ')
            logger.info('Input dy: %s' % (data.dy))
            data.gg = raw_input('Input gg: ')
            logger.info('Input gg: %s' % (data.gg))
            data.gg = data.gg.replace('*', u'\xd7')
            data.tj = float(raw_input('Input price of copper: '))
            data.tjv = data.tj
            data.tj = get_int(data.tj)
            logger.info('Input price of copper: %f' % (data.tjv))
            data.tj = u'\u94dc\u4ef7' + data.tj + u'\u4e07\u5143/\u5428'
            data.ll = raw_input('Input letters: ')
            logger.info('Input letters: %s' % (data.ll))

            key_list = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's',
                        't', 'u',
                        'v', 'w', 'x', 'y', 'z', 'aa', 'ab', 'ac', 'ad', 'ae', 'af', 'ag']
            row_data = []
            price_data = []
            result = 0
            for x in range(0, ob.nsheets):
                sh = ob.sheet_by_index(x)
                price_data = sh.row_values(1)
                for i in range(0, sh.nrows):
                    row_data = sh.row_values(i)
                    if data.xh in row_data and data.gg in row_data and (
                                    data.dy.decode('utf-8') in row_data or data.dy == ''):
                        key_list2 = key_list[0:len(row_data)]
                        dic = dict(zip(key_list2, row_data))
                        price_dic = dict(zip(price_data, row_data))
                        # unit price = (1-percent) * x + percent * y
                        per = (1 - ((data.get_gap() / 0.1) * 0.2)) * price_dic[data.get_price()] + ((
                                                                                                        data.get_gap() / 0.1) * 0.2) * \
                                                                                                   price_dic[
                                                                                                       data.get_price_plus()]
                        print 'Price per meter of %s is %f' % (data.tj, per)
                        logger.info('Price per meter of %s is %f' % (data.tj, per))
                        letter_list = re.findall(r'a\w|\w', data.ll)
                        for a in letter_list:
                            result += dic.get(a)
                            print 'letter ' + a + ' = ' + str(dic.get(a))
                            logger.info('letter ' + a + ' = ' + str(dic.get(a)))
                        print 'Unit price = %f + %f = %f' % (per, result, per + result)
                        logger.info('Unit price = %f + %f = %f' % (per, result, per + result))
                        length = int(raw_input('Enter length = '))
                        logger.info('Enter length = %d' % (length))
                        print 'Result = %f X %d = %f' % (per + result, length, (per + result) * length)
                        logger.info('Result = %f X %d = %f' % (per + result, length, (per + result) * length))
                        break
                    elif i == sh.nrows - 1:
                        print 'No such model in sheet ' + str(x + 1) + '!'
                        logger.info('No such model in sheet ' + str(x + 1) + '!')
                        break
            raw_input('Press Enter to continue...\n')
            logger.info('Press Enter to continue...\n')
    except Exception, e:
        print e
        logger.info(e)
