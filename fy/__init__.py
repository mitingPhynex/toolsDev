import pandas as pd
import requests


def getRu(text,fet):
    url = "https://api-free.deepl.com/v2/translate"
    headers = {
        "Authorization":"DeepL-Auth-Key 745614b8-7a4a-0cc9-0137-26a7263f4713:fx",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.50",
        "Content-Type": "application/x-www-form-urlencoded"
    }

    data = {
        "text":text,
        "target_lang":"RU"
    }

    res = fet.post(url,headers=headers,data=data)
    js = res.json()
    if "translations" in js and len(js['translations'])>0:
        return js['translations'][0]['text']
    return ''

def fy(input,name):
    # 读取原始 Excel 文件
    df = pd.read_excel(input,header=None)
    # 遍历每一行并应用规则匹配
    for index, row in df.iterrows():
        value = row[0]  # 假设原始数据在 'Column' 列中
        s = requests.session()
        df.at[index, 1] = getRu(value,s)
    # 保存为新的 Excel 文件
    newName = name.split(".")[0]+"_over.xlsx"
    df.to_excel("static/{}".format(newName), header=False, index=False)
    return newName