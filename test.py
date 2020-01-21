# -*- coding: utf-8 -*-
"""
Created on Wed Jan  8 18:49:10 2020

@author: user
"""
import os
from os.path import split
import xlrd
import xlsxwriter


def get_classes(filename):
    clses = []
    sheet = xlrd.open_workbook(filename).sheets()[0]
    i, j = 4, 1
    value = ''
    target_first_name = ['進二技', '碩專', '進四技', '進博雅']
    for i in range(1, 15):
        clses.append([])
        for j in range(4, 25, 3):
            for k in target_first_name:
                if (k in str(sheet.cell_value(i, j+2))):
                    value = str(sheet.cell_value(i, j).split(
                        '(')[0].split('-')[0])[0:5]
                    value += ' ' + str(sheet.cell_value(i,
                                                        j+1).split('(')[0][:-2])
                    value += '\n' + str(sheet.cell_value(i, j+2).split('(')[0])
                    break
                elif(len(str(sheet.cell_value(i, j+2)))):
                    value = '非本處課程'
            clses[i-1].append(value)
            value = ''
    return clses


def set_columns(sh, l, s):
    sh.set_column(0, 0, s)
    sh.set_column(1, 7, l)
    sh.set_column(8, 8, s)
    sh.set_column(9, 15, l)
    sh.set_column(16, 16, s)
    sh.set_column(17, 23, l)


def write_xlsx(clses, filename):
    wb = xlsxwriter.Workbook(filename)
    sh = wb.add_worksheet()

    newLine = 3

    day = ['一', '二', '三', '四', '五', '六', '日']
    cls_num = [1, 2, 3, 4, '', 5, 6, 7, 8, 9, 10, 11, 12, 13]

    # setting of format1
    cell_format = wb.add_format(
        {
            'bold': True,
            'font': '微軟正黑體',
            'align': 'vcenter',
            'font_size': '7'
        }
    )
    cell_format.set_align('center')
    cell_format.set_border(1)

    # setting of format2
    cell_format2 = wb.add_format(
        {
            'bold': True,
            'font_size': '12',
            'font_color': 'white',
            'bg_color': 'gray',
            'font': '微軟正黑體',
            'text_wrap': True
        }
    )
    cell_format2.set_align('vcenter')
    cell_format2.set_align('center')

    # setting of format3
    cell_format3 = wb.add_format(
        {
            'bold': True,
            'font_size': '12',
            'font_color': 'white',
            'bg_color': 'red',
            'font': '微軟正黑體'
        }
    )
    cell_format3.set_align('vcenter')
    cell_format3.set_align('center')

    set_columns(sh, 12, 5)

    for i, el1 in enumerate(clses):
        # write filename on left-top place
        sh.write((i//newLine)*15, i % newLine*8,
                 files[i].split('.')[0], cell_format3)

        # write week on top
        for j in range(7):
            sh.write(i//newLine*15,  i %
                     newLine*8 + j + 1, day[j], cell_format2)

        # write class number on left
        for j in range(14):
            sh.write(i//newLine*15+j+1, i % newLine*8,
                     cls_num[j], cell_format2)

        # write data
        for j, el2 in enumerate(el1):
            for k, el3 in enumerate(el2):
                sh.write(1+(i//newLine)*15+j, 1+i %
                         newLine*8+k, el3, cell_format)
                # y_x

        # blank compression
        if(i % newLine == 0):
            for j in range(15):
                if(j == 5 or j == 10):
                    sh.set_row((i//newLine*15)+j, 1,
                               wb.add_format({'color': 'red'}))
                else:
                    sh.set_row((i//newLine*15)+j, 22)
    # #A3pixel = 4961 * 3508

    wb.close()
    print('Done')


if __name__ == '__main__':
    fp = 'cls\YU'
    files = os.listdir(fp)
    clses = []
    for i in files:
        clses.append(get_classes(fp + '\\' + i))

    write_xlsx(clses, 'yee.xlsx')
