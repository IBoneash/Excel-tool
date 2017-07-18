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

logger = logging.getLogger('test')
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


class Xls(object):
    def __init__(self, __name='test.xls'):
        self.xls_name = __name
        self.xlrd_object = None
        self.isopenfailed = True

    def open(self):
        try:
            self.xlrd_object = xlrd.open_workbook(self.xls_name)
            self.isopenfailed = False
            pass
        except:
            self.isopenfailed = True
            self.xlrd_object = None
            print ('Open %s faile \n' % self.xls_name)
            pass
        finally:
            pass
        return self.xlrd_object

    def get_price_data(self, x):
        self.__sh = self.open().sheet_by_index(x)
        for y in range(0, self.__sh.nrows):
            if u'\u5e8f\u53f7' in self.__sh.row_values(y):
                return self.__sh.row_values(y)

    def get_dy(self):
        dy_list = []
        for __x in range(0, self.open().nsheets):
            self.__sh = self.open().sheet_by_index(__x)
            dy_row = None
            for __y in range(0, self.__sh.nrows):
                if u'\u7535\u538b' in self.__sh.row_values(__y):
                    for z in range(0, len(self.__sh.row_values(__y))):
                        if u'\u7535\u538b' in self.__sh.row_values(__y)[z]:
                            dy_row = __y
                            dy_line = z
                            break
            if dy_row is not None:
                for __i in range(dy_row + 1, self.__sh.nrows):
                    if self.__sh.row_values(__i)[dy_line] not in dy_list:
                        dy_list.append(self.__sh.row_values(__i)[dy_line])
        return dy_list


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
            xl = Xls()
            dy_dic = dict(zip(range(1, len(xl.get_dy())+1), xl.get_dy()))
            print dy_dic
            logger.info('dy_list = %s' % (dy_dic))

            data = Input()
            data.xh = raw_input('Input xh: ').upper()
            logger.info('Input xh: %s' % (data.xh))

            data.dy = raw_input('Input dy: ')
            logger.info('Input dy: %s' % (data.dy))
            if data.dy != '':
                data.dy = xl.get_dy()[int(data.dy) - 1]
            print data.dy

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

            for x in range(0, xl.open().nsheets):
                sh = xl.open().sheet_by_index(x)
                price_data = xl.get_price_data(x)
                for i in range(0, sh.nrows):
                    row_data = sh.row_values(i)
                    if data.xh in row_data and data.gg in row_data and (
                                    data.dy.decode('utf8') in row_data or data.dy == ''):
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

                if i == sh.nrows - 1:
                    print 'No such model in sheet ' + str(x + 1) + '!'
                    logger.info('No such model in sheet ' + str(x + 1) + '!')
                    continue

            raw_input('Press Enter to continue...\n')
            logger.info('Press Enter to continue...\n')

    except Exception, e:
        print e
        logger.info(e)
