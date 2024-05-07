import pandas as pd
from openpyxl import load_workbook
import zipfile
from openpyxl.utils.dataframe import dataframe_to_rows
def cut(obj, sec):
    if sec == 0:
        return [obj]
    return [obj[i:i+sec] for i in range(0,len(obj),sec)]

# 输入最初的文件
def bq2(input,name,num):
    newNameLists = []
    template_path = './bq/mb.xlsx'
    zipName = name.split(".")[0]+"zip_over.zip"
    zipFile = zipfile.ZipFile("./static/{}".format(zipName), 'w')
    # df = pd.read_excel('input.xlsx', sheet_name="Vendor template",header=1)
    df = pd.read_excel(input, sheet_name="Vendor template",header=1)
    russian_sizes = ["46", "48", "50", "52", "54", "56"]
    manufacturer_sizes = ['L', 'XL', '2XL', '3XL', '4XL', '5XL']
    # 将每一行分成六份，并将它们在一起
    new_rows = []
    for i in range(len(df)):
        if i == 0:
            continue
        for j in range(6):
            df.loc[i, 'Manufacturer Size'] = manufacturer_sizes[j]
            df.loc[i, 'Russian Size*'] = russian_sizes[j]
            df.loc[i, 'Size of the Product in the Photo'] = russian_sizes[j]
            if i>=2:
                if df.loc[i, "Product name"] == df.loc[i-1, "Product name"]:
                    df.loc[i, "Combine into One PDP*"] = df.iloc[i-1]["Combine into One PDP*"]
                else:
                    buff =  str(int(df.iloc[i-1]["Combine into One PDP*"][-3:])+1)
                    if len(buff) == 2:
                        buff = "0"+buff
                    elif len(buff) == 1:
                        buff = "00"+buff
                    df.loc[i, "Combine into One PDP*"] = df.iloc[i-1]["Combine into One PDP*"][:-3] +  buff
            print(i,df.iloc[i]["Combine into One PDP*"],df.iloc[i]["Product Colour*"],manufacturer_sizes[j])
            df.loc[i, "article*"] = df.iloc[i]["Combine into One PDP*"] + " " + df.iloc[i]["Product Colour*"] + " " + manufacturer_sizes[j]
            row = df.iloc[i].tolist()
            new_rows.append(row)
    columns = df.columns.tolist()
    new_df = pd.DataFrame(new_rows, columns=columns)
    # 写入数据
    for index,li in enumerate(cut(new_df,num)):
        template_wb = load_workbook(template_path)
        worksheet = template_wb.active
        newName = name.split(".")[0]+ str(index) +"_over.xlsx"
        for row in dataframe_to_rows(li, index=False, header=False):
            worksheet.append(row)
        for cell in worksheet['U']:
            if cell.value is not None:
                try:
                    cell.value = int(cell.value)
                except:
                    pass
        for cell in worksheet['AF']:
            if cell.value is not None:
                try:
                    cell.value = int(cell.value)
                except:
                    pass
        template_wb.save("./static/{}".format(newName))
        newNameLists.append(newName)
        zipFile.write("./static/{}".format(newName), newName, zipfile.ZIP_DEFLATED)
        # new_df.to_excel("../static/{}".format(newName), index=False,header=None)
    zipFile.close()
    return zipName
# n = bq("input.xlsx","input.xlsx",7)
# print(n)