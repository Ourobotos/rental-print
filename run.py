#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
    xijie15.__main__
    Alias for xijie15.run for the command line.

"""
from rental_printer.xlsxio import XlsxIO
from pprint import pprint

if __name__ == '__main__':
    x = XlsxIO('DATA.xlsx')
    pprint(x.wl)
    x.output('TEMPLATE.xlsx')
    # one_list = [0, 0, 5, 2, ' ', ' ', ' ']
    # one_list.reverse()
    # print(one_list)
    # x.write_out_by_template('TEMPLATE.xlsx')
