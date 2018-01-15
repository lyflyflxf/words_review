#! /usr/bin/env python
# -*- coding: gbk -*-

import pandas as pd
import main
from pdb import set_trace as st
import os
import pyttsx3

files = {"���е���": "high_ran.xlsx",
         '21��list': "21_ran.xls"}
unit = 'Unit'
english = '����'

if __name__ == "__main__":
    while 1:
        student = input('������ѧ��������')  # '2':#
        if student not in main.names or student == '�����':
            print('����������������롣')
            continue
        else:
            break
    if student == '���κ�':
        v_down = 20
    elif student == '���ɭ':
        v_down = 90

    f_name, today = main.today(student)
    xls = pd.ExcelFile(files[f_name])
    s = xls.parse('Sheet1')
    units = s.loc[:, unit]
    w_dict = {}
    for each in today:
        w_dict.update({each: list(s[units == int(each)][english])})

    engine = pyttsx3.init()
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate - v_down)
    voices = engine.getProperty('voices')

    for wl_name, words in w_dict.items():
        word_l = len(words)
        i = 0
        while i < word_l:
            word = words[i]


            def this():
                engine.setProperty('voice', voices[1].id)
                engine.say(word)
                engine.setProperty('voice', voices[2].id)
                engine.say(word)

                engine.runAndWait()


            while 1:
                os.system('cls')
                print('list', wl_name, '�ĵ�', i + 1, '�����ʣ���', word_l, '����\n�س�����һ�ʣ� 1+�س�����һ�ʣ�������+�س������²��ţ�')
                this()
                type_in = input()
                if type_in == '':
                    break
                elif type_in == '1':
                    i -= 2
                    break
                # sys.exit()
                else:
                    continue  # print('�����������������')

            i += 1
        # if type_in==
    # while 1:
    #     tasks=main.today('���ɭ')#input('���������ȫ����'))
    #     print(tasks)
    #     if tasks==None:
    #         print('��ϲ�㣬����û������')
    #         exit()
    #     elif tasks=='error':
    #         os.system('cls')
    #         print('�����������������')
    #         continue
    #     else:
    #         os.system('cls')
    #         print('���ڿ�ʼ��д...')
    #         break
    # st()
    # for k,v in files.items():
    #     xls = pd.ExcelFile(v)
    #     print(k)