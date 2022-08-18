import pandas as pd
data = pd.read_excel('spimport.xlsx')
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 1000)
#完全相等过程
for i in range (95):
    for k in range(16):
        data.iloc[(i + 1) * 16 + k, 7] = data.iloc[i * 16 + k, 7]
        data.iloc[(i + 1) * 16 + k, 8] = data.iloc[i * 16 + k, 8]
#定义替换递归函数
def repalce(i, j, k):
    data.iloc[i * 16 + k, j] = data.iloc[(i - 1) * 16 + k, j] + (data.iloc[(i-1) * 16 + k, j] - data.iloc[(i - 2) * 16 + k, j])
#第一台箱变递归过程
for i in range(2, 32):
    for j in range(9, 13):
        for k in range(16):
            repalce(i, j, k)
    j = 6
    for k in range(16):
        repalce(i, j, k)

#后面箱变迭代step1:互相相等
for i in range(32):
    for k in range(16):
        j = 6
        data.iloc[i * 16 + k + 32 * 16, j] = data.iloc[i * 16 + k, j]
        data.iloc[i * 16 + k + 32 * 32, j] = data.iloc[i * 16 + k, j]
        for j in range(9, 13):
            data.iloc[i * 16 + k + 32 * 16, j] = data.iloc[i * 16 + k, j]
            data.iloc[i * 16 + k + 32 * 32, j] = data.iloc[i * 16 + k, j]

#箱变迭代step2：不同箱变通道差值
for i in range(32):
    data.iloc[i * 16 + 32 * 16, 9] = data.iloc[i * 16, 9] + 20480
    data.iloc[i * 16 + 32 * 32, 9] = data.iloc[i * 16, 9] + 40960
    data.iloc[i * 16 + 32 * 16, 10] = 2
    data.iloc[i * 16 + 32 * 32, 10] = 3
    for k in range(4, 13):
        data.iloc[i * 16 + k + 32 * 16, 9] = data.iloc[i * 16 + k, 9] + 20480
        data.iloc[i * 16 + k + 32 * 32, 9] = data.iloc[i * 16 + k, 9] + 40960
        data.iloc[i * 16 + k + 32 * 16, 10] = 2
        data.iloc[i * 16 + k + 32 * 32, 10] = 3
    for k in range(14, 16):
        data.iloc[i * 16 + k + 32 * 16, 9] = data.iloc[i * 16 + k, 9] + 2560
        data.iloc[i * 16 + k + 32 * 32, 9] = data.iloc[i * 16 + k, 9] + 5120
        data.iloc[i * 16 + k + 32 * 16, 10] = 2
        data.iloc[i * 16 + k + 32 * 32, 10] = 3
print(data)
data.to_excel("data2.xlsx")
