#!/usr/bin/env python3
# -*- coding: utf-8 -*-
""" HikikakuKunSheets: extract game strategies of Shogi games
   and put Excel Sheets for viewing shogi games in convenient ways."""

import sys
import os
import os.path
import re
import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def retrieve_files(directly):
    """ retrieve csa and kif files.
    """
    new_files = []
    csa_files = []

    for root, dirs, files in os.walk(directly):
        dirs = dirs
        for fil in files:
            path = os.path.join(root, fil)
            if fil.endswith('.kif'):
                new_files += [path]
                # print "kif:%s" % path
            elif fil.endswith('.csa'):
                csa_files += [path]
                # print "csa:%s" % path

    remove_list = []
    csa_new_files = []
    for cfi in csa_files:
        root, ext = os.path.splitext(cfi)
        ext = ext
        kiffile = root + '.kif'

        if not os.path.isfile(kiffile):
            print('remove: %s' % kiffile)
            remove_list += [cfi]
        else:
            csa_new_files += [cfi]

    # print "newFiles = %d" % len(newFiles)
    # print "csaFiles = %d" % len(csaFiles)
    # print "csaNewFiles = %d" % len(csaNewFiles)
    # print "removeList = %d" % len(removeList)

    return csa_new_files, new_files


def calc_senkei(kif_file):
    """ main routine of senkei calculation.
    """
    date = re.compile(
        r'開始日時：([0-9]+)/([0-9]+)/([0-9]+) ([0-9]+):([0-9]+):([0-9]+)\r\n')
    sente = re.compile(r'先手：(.+)\r\n')
    gote = re.compile(r'後手：(.+)\r\n')
    senkei_line = re.compile(r'戦型：(.+)\r\n')
    made = re.compile(r'まで([0-9]+)手で([先後]手)の(入玉勝ち|勝ち)\r\n')
    summary = re.compile(
        r"'summary:([a-z_ ]+):([^:]+)\ (lose|win|draw):([^:]+)\ (win|lose|draw)")

    reason_map = {'illegal move': '反則負け',
                  'max_moves': '最大手数',
                  'oute_sennichite': '王手千日手',
                  'sennichite': '千日手',
                  'time up': '時間切れ負け',
                  'oute_kaihimore': '王手回避もれ',
                  'uchifuzume': '打ち歩詰め'}
    num = 1
    senkei_table = {'合計': 0}

    print(','.join(['年', '月', '日', '時', '分', '秒', '先手', '後手', '戦型',
                    '手数', '勝者(先後)', '勝者', '敗者', '結果', '棋譜ファイル', 'リンク']))

    records = []

    for k in kif_file:
        num += 1
        senkei_str = '戦型データなし'
        root, ext = os.path.splitext(k)
        ext = ext
        csafile = root + '.csa'
        print('analyze:' + str(num) + ':' + csafile)
        # newline='' keeps '\r\n' so the regexes below still match
        k_file_obj = open(k, 'r', encoding='cp932', newline='')
        try:
            lines = k_file_obj.readlines()
        except Exception:
            print(csafile + ' is ignored!!!!!!!!!!!!!!!!!!!!!!!!!!!')
            num -= 1
            k_file_obj.close()
            continue

        k_file_obj.close()
        #print li
        kifu = []
        broken = False
        for kif_line in lines:
            if kif_line.startswith('開'):
                ma_obj = date.match(kif_line)
                if ma_obj:
                    #print '日時='.encode('utf-8') + repr(m.group(1,2,3,4,5,6))
                    kifu = list(ma_obj.group(1, 2, 3, 4, 5, 6))
                else:
                    kifu = ['2014', '1', '1', '0', '0',
                            '0']  # dummy data (broken file...)
                    broken = True

            ma_obj2 = senkei_line.match(kif_line)
            if ma_obj2:
                #print '戦型='.encode('utf-8') + m.group(1).encode('utf-8')
                senkei_str = ma_obj2.group(1)[:]
                senkei_table['合計'] += 1
                senkei_item = senkei_table.get(senkei_str, 'えらー')
                if senkei_item == 'えらー':
                    senkei_table[senkei_str] = 1
                else:
                    senkei_table[senkei_str] += 1

            ma_obj3 = sente.match(kif_line)
            if ma_obj3:
                #print '先手='.encode('utf-8') + m.group(1).encode('utf-8')
                sente_str = ma_obj3.group(1)[:]

            ma_obj4 = gote.match(kif_line)
            if ma_obj4:
                #print '後手='.encode('utf-8') + m.group(1).encode('utf-8')
                gote_str = ma_obj4.group(1)[:]

        if broken:
            print(csafile + ' is ignored(broken file)!!!!!!!!!!!!!!!')
            num -= 1
            continue

        if senkei_str == '戦型データなし':
            senkei_table['合計'] += 1
            item = senkei_table.get(senkei_str, 'えらー')
            if item == 'えらー':
                senkei_table[senkei_str] = 1
            else:
                senkei_table[senkei_str] += 1

        shouhai = False
        kifu += [sente_str]
        kifu += [gote_str]
        kifu += [senkei_str]

        if lines[len(lines) - 1].startswith('まで'):
            mat = made.match(lines[len(lines) - 1])
            if mat:
                kifu += [mat.group(1) + '手']
                if mat.group(3) == '入玉勝ち':
                    kachi = '入玉勝ち'
                else:
                    kachi = '投了'

                if mat.group(2) == '先手':
                    kifu += ['先手']
                    kifu += [sente_str]
                    kifu += [gote_str]
                    kifu += [kachi]
                else:
                    kifu += ['後手']
                    kifu += [gote_str]
                    kifu += [sente_str]
                    kifu += [kachi]
                shouhai = True
            else:
                print('xxxxx')
                print(lines[len(lines) - 1])

        else:
            csafile2 = open(csafile, 'r', encoding='cp932', newline='')
            li2 = csafile2.readlines()
            csafile2.close()

            kifu += ['？手']

            for csa_line in li2:
                if csa_line.startswith("'summary:"):
                    m_summary = summary.match(csa_line)
                    if m_summary:
                        # print em
                        # print 'CSA '+m.group(1).encode('utf-8')
                        if m_summary.group(3) == 'draw':
                            kifu += ['引き分け']
                            kifu += ['']
                            kifu += ['']
                            kifu += [reason_map[m_summary.group(1)]]
                            shouhai = True
                        elif m_summary.group(3) == 'win':
                            kifu += ['先手']
                            kifu += [sente_str]
                            kifu += [gote_str]
                            kifu += [reason_map[m_summary.group(1)]]
                            shouhai = True
                        elif m_summary.group(5) == 'win':
                            kifu += ['後手']
                            kifu += [gote_str]
                            kifu += [sente_str]
                            kifu += [reason_map[m_summary.group(1)]]
                            shouhai = True
                    else:
                        pass

            if shouhai is False:
                # print 'まで?手勝敗不明：不明'.encode('utf-8')
                kifu += ['不明']
                kifu += ['']
                kifu += ['']
                kifu += ['不明']

        kifu += [csafile[2:]]  # J
        kifu += ['=HYPERLINK(J' + ('%d' % num) + ')']  # K
        kifu += ['=CELL("filename")']  # L
        unum = ('%d' % num)
        kifu += ['=CONCATENATE(LEFT(L' + unum + ',FIND("★",SUBSTITUTE(L' +
                 unum + ',"/","★",LEN(L' + unum + ')-LEN(SUBSTITUTE(L' + unum
                 + ',"/",""))),1)-1),"/",J' + unum + ')']  # M
        kifu += ['=HYPERLINK(RIGHT(M' +
                 unum + ',LEN(M' + unum + ')-1))'] # N

        # kifu_t = ','.join(kifu)
        # print kifu_t.encode('utf-8')
        # print 'analyze:' + str(num) + ':' + csafile
        records += [kifu[:]]

    return records, senkei_table


def write_row(ws1, row_idx, reco):
    """ write kif record to excel sheet.
    """
    col = get_column_letter(1)
    row = str(row_idx)
    cell_name = '%s%s' % (col, row)
    try:
        ws1[cell_name] = datetime.datetime(
            int(reco[0]), int(reco[1]), int(reco[2]), int(reco[3]),
            int(reco[4]), int(reco[5]))
    except Exception:
        print(reco[0:5])

    for i in range(6, len(reco)):
        col2 = get_column_letter(i + 1 - 5)
        cell_name2 = '%s%s' % (col2, row)
        ws1[cell_name2] = reco[i]


def write_excel_xml(kif_record, dest_filename):
    """ write excel file.
    """
    wbook = Workbook()

    ws2 = wbook.active
    ws2.title = '戦型一覧表'

    ws2.append(['試合開始日時', '先手', '後手', '戦型', '手数', '勝者(先後)', '勝者', '敗者',
                '結果', '棋譜ファイル', 'リンク(Excel)', 'ファイルパス1', 'ファイルパス2',
                'リンク(LibreOffice)'])

    for row in range(2, len(kif_record) + 2):
        write_row(ws2, row, kif_record[row - 2])
        print(row - 1)

    wbook.save(filename=dest_filename)
    print("file=%s: %d records are saved(^^)." % (dest_filename,
                                                  len(kif_record)))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('usage: python3 hikikaku_py3.py sheetname')
    else:
        print(sys.argv[1])
        CSA, KIF = retrieve_files(".")
        RECORD, SENKEI = calc_senkei(KIF)
        write_excel_xml(RECORD, sys.argv[1])
        for x in SENKEI.keys():
            print("%s: %d" % (x, SENKEI[x]))
