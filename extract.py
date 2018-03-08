# -*- coding: UTF-8 -*-

import xlrd
import xlwt
import os
import re
import jieba
import levenshtein

def readTextFile():
    f = []
    files = os.listdir('text')
    for file in files:
        f.append(file)
    return f

def readData():
    files = readTextFile()
    datas = {}
    for file in files:
        workbook = xlrd.open_workbook('text' + '/' + file)
        sheet = workbook.sheet_by_index(0)
        count = sheet.nrows
        for i in range(count):
            if(i + 1 < count):
                text = sheet.cell(i + 1, 1).value
                categ = sheet.cell(i + 1, 2).value
                text = replaceUseless(text)
                if(datas.get(categ) is not None):
                    datas[categ].add(text)
                else:
                    vailes = set([])
                    vailes.add(text)
                    datas[categ] = vailes
    return datas

def cut(datas):
    keys = datas.keys()
    texts = {}
    for key in keys:
        values = datas[key]
        print(key + "___%d" % len(values))
        words = []
        for value in values:
            words.extend(jieba.cut(value))
        texts[key] = words
    return texts

def filter():
    datas = readData()
    keys = datas.keys()
    for key in keys:
        print(key)
        words = list(datas.get(key))
        for i in range(len(words)):
            j = i + 1
            while j < len(words):
                distance = levenshtein.minEditDist(words[i], words[j])
                if distance > 95:
                    print(words[i] + "---" + words[j])
                    words.remove(words[j])
                j = j + 1
    return datas

def statistics(datas):
    keys = datas.keys()
    dicts = {}
    for key in keys:
        values = datas.get(key)
        stat = {}
        for value in values:
            if stat.get(value) is not None:
                stat[value] = stat[value] + 1
            else:
                stat[value] = 1
        dicts[key] = stat
    return dicts

def writeResult(dicts):
    wbk = xlwt.Workbook()
    keys = dicts.keys()
    for key in keys:
        sheet = wbk.add_sheet(key)
        dic = dicts[key]
        words = dic.keys()
        line = 0
        for word in words:
            sheet.write(line, 0, word)
            sheet.write(line, 1, dic[word])
            line += 1
    wbk.save('result1.xls')

def replaceUseless(text):
    text = text.replace('的', '')
    text = text.replace('了', '')
    text = text.replace('吗', '')
    text = text.replace('啊', '')
    text = re.sub(r'[^\u4e00-\u9fa5]', "", text)
    return text

if __name__ == '__main__':
    datas = readData()
    words = cut(datas)
    dics = statistics(words)
    writeResult(dics)