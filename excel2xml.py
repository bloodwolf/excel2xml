#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlrd
import sys
import os


def toString(x):
    try:
        if int(x) == x:
            x = int(x)
        return str(x)
    except:
        return x

def excel2xml(filename):
    try:
        bk = xlrd.open_workbook(filename)
        for sh in bk.sheets():
            print 'creating %s.xml' % sh.name
            output = open(sh.name + '.xml', 'w')
            firstline = sh.row_values(0)
            outputstr = '<%s>' % (sh.name)
            output.write(outputstr.encode('utf-8') + "\n")
            line = '<item '
            for i in firstline:
                line += (i + '="%s" ')
            line += '/>'

            for i in xrange(1, sh.nrows):
                args = tuple([toString(x) for x in sh.row_values(i)])
                outputstr = line % args
                output.write(outputstr.encode('utf-8') + "\n")
            outputstr = '</%s>' % (sh.name)
            output.write(outputstr.encode('utf-8'))
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

    excel2xml(filename)
