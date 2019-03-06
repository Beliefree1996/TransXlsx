import pandas as pd

excel_path = 'G:\code\Python\demo\demo.xlsx'
rf = pd.read_excel(excel_path)
# 读取第10列
data_rf = rf.ix[0:,10].values
# 除去空数据
data_df = pd.DataFrame(data_rf)
data_df.dropna(inplace=True)
s_sum = data_df.shape[0]

# 写入excel
pstart = excel_path.index('.')
address = excel_path[0:pstart]
address = address + "_extract.xlsx"
writer = pd.ExcelWriter(address)
data_df.to_excel(writer,index=False,header=False)
writer.save()

# # 读取第11列
data_rf1 = rf.ix[1:,11].values
# 除去空数据并分割出号码;
number = []
for index in range(len(data_rf1)):
    data_str = str(data_rf1[index])
    if data_str.count(';') == 1:
        mobile = data_str.replace(';', '')
        number.append(mobile)
    elif data_str.count(';') == 2:
        pstart = data_str.index(';')
        mobile1 = data_str[0:pstart]
        # print(mobile1)
        mobile2 = data_str[pstart + 1:-1]
        number.append(mobile2)
    elif data_str.count(';') == 3:
        pstart = data_str.index(';')
        mobile3 = data_str[0:pstart]
        number.append(mobile3)
        mobile4 = data_str[pstart + 1:]
        pstart = mobile4.index(';')
        mobile5 = mobile4[0:pstart]
        number.append(mobile5)
        mobile6 = mobile4[pstart + 1:-1]
        number.append(mobile6)
# for index in range(len(number)):
data_df1 = pd.DataFrame(number)
data_df1.dropna(inplace=True)
# 写入excel
s_sum = data_df1.shape[0]
data_df1.to_excel(writer,index=False,header=False,startrow=s_sum)
writer.save()