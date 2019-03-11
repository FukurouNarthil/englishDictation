# -*- coding:utf-8 -*-
# 英文听写程序

import win32com.client as wc
import pymysql as MySQLdb
import random
import time
import sys


# 存入新单词
def save_word(e, c):
    # 连接数据库
    db = MySQLdb.connect("localhost", "root", "123456", "DICTATION")
    cursor = db.cursor()
    sql_stmt = "INSERT INTO new_words(english, chinese) VALUES ('%s', '%s')" % (e, c)
    cursor.execute(sql_stmt)
    db.commit()
    db.close()


# 听写
def dictation():
    db = MySQLdb.connect("localhost", "root", "123456", "DICTATION")
    cursor = db.cursor()
    sql_stmt = 'SELECT * FROM new_words'
    cursor.execute(sql_stmt)
    db.commit()
    results = cursor.fetchall()
    total = len(results)

    words = []

    for i in range(50):
        while True:
            id = random.randint(0, total)
            if (results[id][1], results[id][2]) not in words:
                speaker = wc.Dispatch("SAPI.SpVoice")
                speaker.Speak(results[id][1])
                words.append((results[id][1], results[id][2]))
                time.sleep(10)
                break

    count = 1
    for i in words:
        print(count, end='')
        print(i)
        count += 1
    db.close()


# 查找单词
def search():
    db = MySQLdb.connect("localhost", "root", "123456", "DICTATION")
    cursor = db.cursor()
    while True:
        word = input("input the word being searched: ")
        if word != 'e':
            sql_stmt = "SELECT * FROM new_words WHERE english = '%s'" % word
            cursor.execute(sql_stmt)
            db.commit()
            result = cursor.fetchall()
            if result:
                print("In.")
            else:
                print("Not in.")
        else:
            break


if __name__ == '__main__':

    while True:
        mode = input("input mode: ")
        if mode == 'new':
            eng = input("input english: ")
            chn = input("input chinese: ")
            save_word(eng, chn)
        elif mode == 'd':
            dictation()
            break
        elif mode == 's':
            search()
        elif mode == 'e':
            sys.exit(0)
        else:
            print('enter new or d or e or s plz!')