import pandas as pd
import msoffcrypto
import pathlib
import io
import os
from googletrans import Translator, constants


def unlock(filename, passwd):
    temp = open(filename, 'rb')
    excel = msoffcrypto.OfficeFile(temp)
    excel.load_key(passwd)

    with open("out.xlsx", 'wb') as f:
        excel.decrypt(f)
    temp.close()

def trans(df1,sheetname):

    translator = Translator()
    for col in df1.columns:
        for val in df1[col]:
            print(val)
            tmp = translator.translate(val)
            print(tmp.text)
            df1[col]=tmp.text


def translate_to_english(text):
    translator = Translator()
    translated_text = translator.translate(text, src='ja', dest='en')
    return translated_text.text
