#! /usr/bin/env python
# -*- coding: gbk -*-

# inter=[0,1,2,4,7,15,30]
# for i in range(1, len(inter)):
#     day_list.append(sum(inter[:i]))
# print(day_list)

from datetime import *
import pyperclip
import pandas as pd
from pdb import set_trace as st
import numpy as np
import os

# review_day = [1, 3, 7, 14, 29, 59]
interval = [1, 2, 4, 7, 15, 30]
inter_l = len(interval)

# f_dir = r"E:\py\tools\s_plan\\"
# f_name = 't.xlsx'
dir = r"C:\py\words_review\t.xlsx"

init = '���α�������'
content = '����'
count = '�Ѹ�ϰ����'
last = '�ϴθ�ϰ����'
next = '�´θ�ϰ����'
dl = 'deadline'
sn = '���'
erase = [last, next, dl]
heads = ['���е���', '21��list']

now = pd.Timestamp(date.today())#+pd.Timedelta(days=1)
xls = pd.ExcelFile(dir)
writer = pd.ExcelWriter(dir)

names = xls.sheet_names



def future(l):
    out = []
    for sn in l:
        sn = int(sn)
        if sn < inter_l:
            out.append(interval[sn])
        else:
            out.append(60)
    return [pd.Timedelta(days=each) for each in out]


def initial(s):
    def check_n(col):
        return pd.isnull(s[col]).any()

    def write_n(col, value):
        if check_n(col):
            s.loc[pd.isnull(s[col]), col] = value

    # ���α�������
    write_n(init, now)
    # ���
    write_n(count, 0)
    # �ϴ�
    write_n(last, now)
    # �´�ʱ��
    if check_n(next):
        n = pd.isnull(s[next])
        m = s.loc[n, :]
        s.loc[n, next] = future(m[count]) + m[last]
    # deadlineʱ��
    if check_n(dl):
        n = pd.isnull(s[dl])
        m = s.loc[n, :]
        s.loc[n, dl] = future(m[count] + 1) + m[next]

def tasks(day, df):
    l = len(df)
    if l == 0:
        print(day + '�޸�ϰ����')
        return 0
    else:
        print(day + '�ĸ�ϰ����Ϊ��')
        return l


def today(name):
    if name not in names:
        return 'error'

    s = xls.parse(name)
    initial(s)
    nexts = s.loc[:, next]
    dls = s.loc[:, dl]

    today = s[(nexts <= (now)) & (dls >= now)][content]

    today_no = tasks("����", today)
    if today_no:
        for head in heads:
            out = []
            h_len = len(head)
            for item in list(today):
                if item.startswith(head):
                    read= item[h_len:]
                    out.append()
            if len(out) != 0:
                return head,out
    else:
        return None



if __name__ == '__main__':

    for name in names:
        s = xls.parse(name)
        initial(s)

        nexts = s.loc[:, next]
        dls = s.loc[:, dl]
        print('ѧ����', name)
        # ��������
        tomorrow = s[nexts == (now + pd.Timedelta(days=1))]
        tasks('����', tomorrow)
        l = list(tomorrow[content])

        copy = ''
        for head in heads:
            out = ''
            h_len = len(head)
            for item in l:
                if item.startswith(head):
                    out += item[h_len:] + '��'
            if len(out) != 0:
                copy += head + out
        copy = copy[:-1]
        print(copy)
        pyperclip.copy(copy)
        # ��������
        today = s[(nexts <= (now)) & (dls >= now)]
        today_no = tasks("����", today)
        if today_no:
            today = pd.DataFrame({sn: pd.Series(np.arange(1, 1 + len(today)),
                                                index=today.index)}
                                 ).join(today)
            print(today[[sn, content]].to_string(index=False))

            warn = today[today[next]<now][[sn, dl]]
            if len(warn) != 0:
                print('������Ҫע�������Ϊ')
                print(warn.to_string(index=False))

            while 1:
                type_in = input('������������������ţ��ÿո���:')
                if type_in == '':
                    break

                try:
                    type_in = [int(i) for i in type_in.split(' ')]
                except ValueError:
                    print('�����������������')
                    continue
                error = False
                for each in type_in:
                    if not (today[sn].isin([2]).any()):
                        print('�����������������')
                        error = True
                        break
                if not error:
                    type_in = today[today[sn] == type_in].index
                    s.loc[type_in, count] += 1
                    s.loc[type_in, erase] = np.nan
                    initial(s)
                    break

        while 1:
            type_in = input('����������±������ݡ��ο�ͷ��' + ''.join(heads))
            if type_in == '':
                break

            s.loc[len(s.index), content] = type_in
            initial(s)
            break

        s.to_excel(writer, sheet_name=name)
    writer.save()