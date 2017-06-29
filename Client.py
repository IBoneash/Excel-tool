import xlrd

file_name = 'test.xls'
ob = xlrd.open_workbook(file_name)


class FromRow(object):
    def __init__(self, letter):
        self.letter = letter


if __name__ == '__main__':
    xh_input = raw_input('Input xh: ')
    dy_input = raw_input('Input dy: ')
    gg_input = raw_input('Input gg: ')

    ll_input = raw_input('Input letters: ')

    key_list = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u',
                'v', 'w', 'x', 'y', 'z', 'aa', 'ab', 'ac', 'ad']
    row_list = []
    result = 0
    try:
        for x in range(0, ob.nsheets):
            sh = ob.sheet_by_index(x)
            for i in range(0, sh.nrows):
                if xh_input.upper() in sh.row_values(i) and gg_input.decode('utf-8') in sh.row_values(
                        i) and (dy_input.decode('utf-8') in sh.row_values(i) or dy_input == ''):
                    row_list = sh.row_values(i)
                    key_list2 = key_list[0:len(row_list)]
                    dic = dict(zip(key_list2, row_list))
                    letter_list = list(ll_input)
                    # for l in range(0, len(letter_list)):
                    #     if letter_list[l] == 'a':
                    #         letter_list[l] = letter_list[l] + letter_list[l + 1]
                    #         del letter_list[l + 1]
                    for a in letter_list:
                        result += dic.get(a)
                        print 'letter ' + a + ' is ' + str(dic.get(a))
                    print 'Result = ', result
                    break
                elif i == sh.nrows - 1:
                    print 'No such model in sheet ' + str(x + 1) + '!'
                    break
    except Exception, e:
        print 'no sheet', e









        # nrows = sh.nrows    # ncols = sh.ncols  #  # print nrows, ncols  #
# cell_value = sh.cell_value(1, 3)
#
# row_list = []
#
# # for i in range(1, 3):
#     row_data = sh.row_values(i)
#     row_list.append(row_data)
#
# print cell_value
