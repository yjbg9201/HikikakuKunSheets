#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import re
import tkinter
import shutil


def get_clipboard():
    root = tkinter.Tk()
    root.withdraw()
    try:
        return root.clipboard_get()
    finally:
        root.destroy()


def ccopy(dst):
    os.makedirs(dst)
    cl = get_clipboard()
    li = re.split('\n', cl)
    print(li)
    num = 0
    for src in li[:-1]:
        print('copying...%s to %s.' % (src, dst))
        num += 1
        shutil.copy(src, dst)
    print('%d files are copied. done...(^^)' % num)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('usage: python3 ccopy_py3.py target-path')
    else:
        print(sys.argv[1])
        ccopy(sys.argv[1])
