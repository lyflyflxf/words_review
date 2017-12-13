#! /usr/bin/env python
# -*- coding: gbk -*-

import pandas as pd
import main
from pdb import set_trace as st
import os
import pyttsx3
files= {"high":"high school.xlsx",
    '21':"21_ran.xls"}
unit='Unit'
english='����'

if __name__=="__main__":
    xls = pd.ExcelFile(files['21'])
    s=xls.parse('Sheet1')
    units=s.loc[:, unit]
    words=list(s[units==1][english])
    # print(words)

    engine = pyttsx3.init()
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate - 70)
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[2].id)

    for i, word in enumerate(words):
        def this():
            engine.say(word)
            engine.runAndWait()
        while 1:
            os.system('cls')
            print('��', i + 1, '�����ʣ���', len(words), '����\n�س�����һ��  ������+�س������²��ţ� ')
            this()
            type_in=input()
            if type_in=='':
                break
            else:
                continue#print('�����������������')
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