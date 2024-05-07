import re
import pandas as pd


def pp(input,name,rules_str):
    # 读取原始 Excel 文件
    df = pd.read_excel(input,header=None)
    # 创建新的列用于存储匹配后的结果
    # 定义规则字符串
    # rules_str = "S->d||a->w"
    # 解析规则字符串为规则列表
    rules = [rule_str.split('->') for rule_str in rules_str.split('||')]
    # 遍历每一行并应用规则匹配
    for index, row in df.iterrows():
        value = row[0]  # 假设原始数据在 'Column' 列中
        # 应用规则匹配
        for rule in rules:
            value = value.lower().replace(rule[0].lower(), rule[1].lower())
            # print(rule[0],rule[1])
            # value = re.sub(re.escape(rule[0]), rule[1], value,flags=re.IGNORECASE)
        df.at[index, 1] = " ".join(re.split(r'\s+', value)[:5])
        # df.at[index, 2] = " ".join(re.split(r'\s+', value))
    # 保存为新的 Excel 文件
    newName = name.split(".")[0]+"_over.xlsx"
    df.to_excel("static/{}".format(newName), header=False, index=False)
    return newName
# pp("test.xlsx")


def getRule(input):
    r = []
    df = pd.read_excel(input,header=None)
    # 遍历每一行并应用规则匹配
    for index, row in df.iterrows():
        # print(row)
        o = "" if pd.isna(row[0]) else row[0] # 假设原始数据在 'Column' 列中
        n = "" if pd.isna(row[1]) else row[1]
        # print("{}->{}".format(str(o).replace("\xa0"," "),str(n).replace("\xa0"," ")))
        r.append("{}->{}".format(o,n))
    return "||".join(r)

# s = getRule("标题过滤规则.xlsx")
# print(s)
# r = "Boys/Girls/Mens/Womens->||Short-Sleeved->||S,M,L,XL,XXL->||Short-Sleeve->||S-M-L-XL-XXL->||XS,S,M,L,XL->||XS-S-M-L-XL->||Children's->||M-L-XL-XXL->||S/M-L/XL-->||Tee-shirt->||XS,S,L,XL,->||Tee.Shirt->||Thank-You->||S,M,L,XL->||XS,S,M,XL->||XXX-Large->||Sh!tters->AAAA"
# pp("标题过滤(1).xlsx","aa.xlsx",s)