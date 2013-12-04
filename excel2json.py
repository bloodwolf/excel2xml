#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import sys
import os
import json


def parseType(x):
    if type(x) == float and int(x) == x:
        return int(x)
    return x

def excel2json(filename):
    try:
        bk = xlrd.open_workbook(filename)
        for sh in bk.sheets():
            content = {}
            jsonname = sh.name.encode('utf-8')
            print 'creating %s.json' % jsonname
            output = open(jsonname + '.json', 'w')
            firstline = sh.row_values(0)
            for i in xrange(1, sh.nrows):
                row = {}
                line = sh.row_values(i)
                line = [parseType(x) for x in line]
                for j in xrange(len(firstline)):
                    row[firstline[j]] = line[j]
                content[line[0]] = row
            output.write(json.dumps(content))
            output.close();
    except:
        print 'file format error...'
        sys.exit()

if __name__ == '__main__':
    if len(sys.argv) == 1:
        print 'usage: %s excelfilename' % (sys.argv[0])
        sys.exit()

    filename = sys.argv[1]
    if (not os.path.exists(filename)):
        print '%s does not exist...' % (filename)
        sys.exit()

    excel2json(filename)
