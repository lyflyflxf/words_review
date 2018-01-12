#! /usr/bin/env python
# -*- coding: gbk -*-

# inter=[0,1,2,4,7,15,30]
# for i in range(1, len(inter)):
#     day_list.append(sum(inter[:i]))
# print(day_list)

from datetime import *
import pyperclip
import pandas as pd
import numpy as np
import os

# review_day = [1, 3, 7, 14, 29, 59]
interval = [1, 2, 4, 7, 15, 30]
inter_l = len(interval)

# f_dir = r"E:\py\tools\s_plan\\"
# f_name = 't.xlsx'
dir = os.getcwd()+r"\t.xlsx"

init = '初次背诵日期'
content = '内容'
count = '已复习次数'
last = '上次复习日期'
next = '下次复习日期'
dl = 'deadline'
sn = '序号'
erase = [last, next, dl]
heads = ['高中单词', '21天list']

now = pd.Timestamp(date.today())
xls = pd.ExcelFile(dir)
writer = pd.ExcelWriter(dir)
names = xls.sheet_names


def future(dates):
    out = []
    for sn in dates:
        sn = int(sn)
        if sn < inter_l:
            out.append(interval[sn])
        else:
            out.append(60)
    return [pd.Timedelta(days=each) for each in out]


def initial(s):
    def null_bools(col):
        return pd.isnull(s[col])

    def copy(col, value):
        if null_bools(col).any():
            s.loc[null_bools(col), col] = value

    def sum_up(col, plus=0):
        def cells(row, col):
            return s.loc[null_bools(row), col]

        copy(col,
             future(cells(col, count) + plus) + cells(col, last))

    # index
    if pd.isnull(s.index).any():
        s.index = range(len(s))
    # 初次背诵日期
    copy(init, now)
    # 编号
    copy(count, 0)
    # 上次
    copy(last, now)
    # 下次时间
    sum_up(next, 0)
    # deadline时间
    sum_up(dl, 1)


def tasks(day, df):
    length = len(df)
    if length == 0:
        print(day + '无复习任务。')
        return 0
    else:
        print(day + '的复习任务为：')
        return length


def today(name):
    if name not in names:
        return 'error'
    s = xls.parse(name)
    initial(s)
    nexts, dls = s.loc[:, next], s.loc[:, dl]
    today = s[(nexts <= (now)) & (dls >= now)][content]
    today_no = tasks("今天", today)
    if today_no:
        for head in heads:
            out = []
            h_len = len(head)
            for item in list(today):
                if item.startswith(head):
                    read = item[h_len:]
                    out.append(read)
            if len(out) != 0:
                return head, out
    else:
        return None


if __name__ == '__main__':
    for name in names:
        s = xls.parse(name, index_col=0)
        initial(s)

        nexts = s.loc[:, next]
        dls = s.loc[:, dl]
        print('学生：', name)
        # 明天任务
        tomorrow = s[nexts == (now + pd.Timedelta(days=1))]
        tasks('明天', tomorrow)
        tmr_tasks = list(tomorrow[content])

        copy = ''
        for head in heads:
            out = ''
            h_len = len(head)
            for item in tmr_tasks:
                if item.startswith(head):
                    out += item[h_len:] + '、'
            if len(out) != 0:
                copy += head + out
        copy = copy[:-1]
        print(copy)
        pyperclip.copy(copy)
        # 今天任务
        today = s[(nexts <= now) & (dls >= now)]
        today_no = tasks("今天", today)
        if today_no:
            today = pd.DataFrame({sn: pd.Series(np.arange(1, 1 + len(today)),
                                                index=today.index)}
                                 ).join(today)
            print(today[[sn, content]].to_string(index=False))

            warn = today[today[next] < now][[sn, dl]]
            if len(warn) != 0:
                print('其中需要注意的任务为')
                print(warn.to_string(index=False))

            while 1:
                type_in = input('请输入已完成任务的序号，用空格间隔:')
                if type_in == '':
                    break
                try:
                    type_in = [int(i) for i in type_in.split(' ')]
                except ValueError:
                    print('输入错误，请重新输入')
                    continue
                error = False
                for each in type_in:
                    if not (today[sn].isin([2]).any()):
                        print('输入错误，请重新输入')
                        error = True
                        break
                if not error:
                    type_in = today[today[sn] == type_in].index
                    s.loc[type_in, count] += 1
                    s.loc[type_in, erase] = np.nan
                    initial(s)
                    break

        while 1:
            type_in = input('请输入今天新背的内容。参考头：' + ''.join(heads))
            if type_in == '':
                break

            s.loc[len(s.index), content] = type_in
            initial(s)
            break

        s.to_excel(writer, sheet_name=name)
    # writer.save()
