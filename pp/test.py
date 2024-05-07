import pandas as pd

# 读取原始 Excel 文件
df = pd.read_excel('test.xlsx',header=None)

# 创建新的列用于存储匹配后的结果
# 定义规则字符串
rules_str = "S->d||a->w"

# 解析规则字符串为规则列表
rules = [rule_str.split('->') for rule_str in rules_str.split('||')]

# 遍历每一行并应用规则匹配
for index, row in df.iterrows():
    value = row[0]  # 假设原始数据在 'Column' 列中

    # 应用规则匹配
    for rule in rules:
        value = value.replace(rule[0].upper(), rule[1])
        value = value.replace(rule[0].lower(), rule[1])
    df.at[index, 1] = value

# 保存为新的 Excel 文件
df.to_excel('out.xlsx', header=False, index=False)
