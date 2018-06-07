#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import sys
import os
import json

def excel2json(filename):
    bk = xlrd.open_workbook(filename)
    for sh in bk.sheets():
        if sh.nrows == 0:
            continue
        content = []
        jsonname = sh.name
        print('creating %s.json' % jsonname)
        output = open(jsonname + '.json', 'w')
        firstline = sh.row_values(0)
        for i in range(1, sh.nrows):
            row = {}
            line = sh.row_values(i)
            for j in range(len(firstline)):
                row[firstline[j]] = line[j]
            content.append(row)
        for crange in sh.merged_cells:
            for i in range(crange[0], crange[1]-1):
                for j in range(crange[2], crange[3]):
                    content[i][firstline[j]] = content[crange[0]-1][firstline[crange[2]]]
        output.write(json.dumps(content))
        output.close();

if __name__ == '__main__':
    if len(sys.argv) == 1:
        print('usage: %s excelfilename' % (sys.argv[0]))
        sys.exit()

    filename = sys.argv[1]
    if (not os.path.exists(filename)):
        print('%s does not exist...' % (filename))
        sys.exit()

    excel2json(filename)
