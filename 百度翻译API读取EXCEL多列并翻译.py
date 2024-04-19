# -*- coding: utf-8 -*-

import requests
import random
from hashlib import md5
import pandas as pd
import openpyxl
import configparser


# Set your own appid/appkey.  我在2021/3/25日申请的百度翻译API,认证账户
# https://fanyi-api.baidu.com/api/trans/product/desktop?req=developer

def config_read(config_path, section='DingTalkAPP_chatGLM', option1='Client_ID', option2=None):
    """
    option2 = None 时,仅输出第一个option1的值; 否则输出section下的option1与option2两个值
    """
    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')
    option1_value = config.get(section=section, option=option1)
    if option2 is not None:
        option2_value = config.get(section=section, option=option2)
        return option1_value, option2_value
    else:
        return option1_value


# For list of language codes, please refer to `https://api.fanyi.baidu.com/doc/21`
# from_lang = 'auto'
# to_lang =  'en'


# Generate salt and sign
def make_md5(s, encoding='utf-8'):
    return md5(s.encode(encoding)).hexdigest()


# 定义百度翻译ＡＰＩ函数
def BaiduTranslateAPI(text, appid, appkey, from_lang='auto', to_lang='en'):
    """
    appid: 百度openAPI appid
    appkey: 百度openAPI appkey
    """
    salt = random.randint(32768, 65536)
    sign = make_md5(appid + text + str(salt) + appkey)
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload = {'appid': appid, 'q': text, 'from': from_lang, 'to': to_lang, 'salt': salt, 'sign': sign}
    endpoint = 'http://api.fanyi.baidu.com'
    path = '/api/trans/vip/translate'
    url = ''.join([endpoint, path])
    r = requests.post(url, params=payload, headers=headers)
    result = r.json()
    # 第一个方括号为字典的指定key取值,第二个方括号为之前取出的值为列表,列表的第一个元素,
    # 第三个方括号为取出的第一个元素继续为字典,取字典的值
    return result['trans_result'][0]['dst']


# 定义函数:DataFrame单列遍历,调用百度翻译API,并生成翻译后的DataFrame

def ColumnTranslate(DFColumn, from_lang='auto', to_lang='en'):
    """
    翻译单列, 输入DataFrame列,翻译后,返回DataFrame列,包括列名翻译;空白文本自动跳过;
    DFColum: DataFrame列
    from_lang: 源语言
    to_lang:目标语言
    """
    Translate = []
    for key, value in DFColumn.iteritems():
        # print(value)
        if value != '':
            Translate.append(BaiduTranslateAPI(value, from_lang, to_lang))
        else:
            Translate.append(value)
    Column_name = BaiduTranslateAPI(DFColumn.name, from_lang, to_lang)  # 因为单列DataFrame变成了序列,其列名变成了序列的name
    # print(Column_name)
    Translate = pd.DataFrame(Translate, columns=[Column_name])
    return Translate


def WriteColumnTranslate(TranslateDF, WriteDirectory, WriteExcelName, WriteSheetName, StartRow=0,
                         StartCol=None):
    """
    将单列DataFrame翻译结果,写入已经存在的Excel文件中,指定的Sheet表的指定单元格
    WriteNewSheet: bool; True表示写入新建的sheet表;
    """
    Workbook = openpyxl.load_workbook(WriteDirectory + WriteExcelName)  # 读取要写入的workbook
    # '''''''''
    if Write2NewSheet:
        mode = 'a'
        # mode='a'意味着追加,即在存在的Excel文件中,追加写入,但是写入的位置是新增Sheet表(表名可以自定义).
        # 该追加写入的用法,不需要book和sheets复制;
        # with pd.ExcelWriter(WriteDirectory+WriteExcelName,mode='a', engine="openpyxl") as writer:

    else:
        mode = 'w'

    with pd.ExcelWriter(WriteDirectory + WriteExcelName, mode=mode, engine="openpyxl") as writer:
        writer.book = Workbook  # 此时的writer里还只是读写器. 然后将上面读取的Workbook复制给writer
        writer.sheets = dict((ws.title, ws) for ws in Workbook.worksheets)  # 复制存在的表
        TranslateDF.to_excel(writer, WriteSheetName, startrow=StartRow,
                             startcol=StartCol, index=False, )
        writer.save()
        writer.close()


if __name__ == "__main__":
    # 读取需要翻译的Excel文件,装入DataFrame
    from_lang = 'auto'
    to_lang = 'en'  # 中文: zh;文言文: wyw;日本: jp; --> 伊朗语: ir; 波斯语: per;(仅企业认证的尊享版用户可以使用非常见语种)
    # ExcelDirectory = 'E:/Working Documents/汇龙股份/伊朗/新业务/矿产/BitCoinMining/DH管理/报销/'
    ExcelDirectory = r"E:/Working Documents/Eastcom/Russia/Igor/专网/LeoTelecom/发货测试验收/"
    # ExcelName = '个人报销 -翻译测试.xlsx'
    ExcelName = r'R国箱号.xlsx'
    ExcelSheet = 'Sheet1'
    Read_RowHeader = 2
    Read_Column = [5]  # 也可以用列表读取多列
    ExcelData = pd.read_excel(ExcelDirectory + ExcelName, ExcelSheet,
                              header=Read_RowHeader, usecols=Read_Column, nrows=10)
    # 先处理缺失值,填充为''
    ExcelData.fillna('', inplace=True)
    # ExcelData[ExcelData.keys()[2]]

    # 单列DataFrame翻译结果,写入已经存在的Excel文件中,指定的Sheet表的指定单元格
    WriteDirectory = ExcelDirectory
    # WriteExcelName = '个人报销 -翻译测试.xlsx'
    WriteExcelName = ExcelName
    # WriteSheetName = '20200831'
    WriteSheetName = 'Sheet3'
    StartRow = 2
    StartCol = 6

    # 多列循环调用翻译结果写入函数,写入
    # TranslateDF = pd.DataFrame()
    # for i in range(0, 1):
    #     temp = ColumnTranslate(ExcelData[ExcelData.keys()[i]])
    #     TranslateDF[temp.keys()] = temp
    #     # print(temp.keys()[0])
    #     WriteColumnTranslate(TranslateDF, WriteDirectory, WriteExcelName, WriteSheetName, StartRow, StartCol)
    # TranslateDF

    with pd.ExcelWriter(WriteDirectory + WriteExcelName, mode='a', engine="openpyxl",
                        if_sheet_exists="overlay") as writer:
        # Workbook = openpyxl.load_workbook(WriteDirectory + WriteExcelName)  # 读取要写入的workbook
        # writer.book = Workbook  # 此时的writer里还只是读写器. 然后将上面读取的Workbook复制给writer
        # writer.sheets = dict((ws.title, ws) for ws in Workbook.worksheets)  # 复制存在的表
        ExcelData.to_excel(writer, WriteSheetName, startrow=StartRow,
                           startcol=StartCol, index=False, )
        # writer.save()
        # writer.close()
