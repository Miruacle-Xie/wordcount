# -*- coding: utf-8 -*-
"""
Created on Tue Aug  9 09:52:43 2022

@author: Administrator
"""
from nltk.probability import FreqDist
from nltk import word_tokenize
from nltk.util import ngrams
import pandas as pd
import os
import time


def create_ngrams(text, number):
    textNgrams = []
    for line in text:
        token = word_tokenize(line)
        tmpText = [" ".join(x) for x in list(ngrams(token, number))]
        textNgrams += tmpText
    return textNgrams


def wordfreq(originalData, ngrams):
    wordgram = create_ngrams(originalData, ngrams)
    fdist = FreqDist(wordgram)
    return fdist


def create_excelwriter(filename):
    excelWriter = pd.ExcelWriter(filename, engine='openpyxl')
    return excelWriter


def saveclose_excelwriter(excelWriter):
    excelWriter.save()
    excelWriter.close()


def wordfreqreport(excelWriter, textValue, filename, ngram, col, tittlename):
    subjectFreq = wordfreq(textValue, ngram)
    df = pd.DataFrame(subjectFreq.items(), columns=[
        tittlename, 'freq']).sort_values(by='freq', ascending=False)
    df["占比"] = df["freq"] / subjectFreq.N()
    df["占比"] = df["占比"].apply(lambda x: format(x, '.2%'))
    # print(type(df["占比"][0]))
    # print(df["占比"][0])
    df.to_excel(excelWriter, sheet_name="词频", index=False, startcol=col)


def getwordfreq(text, filename):
    excelWriter = create_excelwriter(filename)
    text.to_excel(excelWriter, index=False)
    wordfreqreport(excelWriter, text.str.lower(), filename, 1, 0, "1个词")
    wordfreqreport(excelWriter, text.str.lower(), filename, 2, 5, "2个词")
    wordfreqreport(excelWriter, text.str.lower(), filename, 3, 10, "3个词")
    saveclose_excelwriter(excelWriter)


def main():
    filepath = input("\n输入文件路径：\n")
    # filepath = "F:\\JetBrains\\PycharmProjects\\pythonProject\\wordcount\\test.xlsx"
    filepath = filepath.replace("\"", "").replace("\'", "")
    reportpath = os.path.splitext(filepath)[0] + "-词频报告" + ".xlsx"
    # print(reportpath)
    print("正在生成词频报告, 请稍等...")
    originalData = pd.read_excel(filepath)
    time_start = time.time()
    getwordfreq(originalData.iloc[:, 0], reportpath)
    time_end = time.time()
    input("已生成报告, 耗时时间:%f, 按回车键结束" % (time_end - time_start))


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        input(str(e)+"\n\n运行异常,按回车键结束")
