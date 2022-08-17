import pandas as pd
data = pd.read_excel('spimport.xlsx')
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 1000)
#完全相等过程
for i in range (95):
    for k in range(15):
        data.iloc[(i + 1) * 16 + k, 7] = data.iloc[i * 16 + k, 7]
        data.iloc[(i + 1) * 16 + k, 8] = data.iloc[i * 16 + k, 8]
#定义替换递归函数
def repalce(i, j, k):
    data.iloc[i * 16 + k, j] = data.iloc[(i - 1) * 16 + k, j] + (data.iloc[(i-1) * 16 + k, j] - data.iloc[(i - 2) * 16 + k, j])
#第一台箱变递归过程
for i in range(2, 32):
    for j in range(9, 12):
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
        for j in range(9, 12):
            data.iloc[i * 16 + k + 32 * 16, j] = data.iloc[i * 16 + k, j]
            data.iloc[i * 16 + k + 32 * 32, j] = data.iloc[i * 16 + k, j]
#     repalce(i,j = 9,k = 0,m = 20480)
#     repalce(i, j=11, k=0, m=20480)
# for k in range(4, 12):
#     for i in range(32):
#         data.iloc[i * 16 + k, 9] = data.iloc[k, 9] + i * 42
#         data.iloc[i * 16 + k + 32 * 16, 9] = data.iloc[i * 16 + k, 9] + 20480
#         data.iloc[i * 16 + k + 32 * 32, 9] = data.iloc[i * 16 + k, 9] + 20480 * 2
# for k in range(14, 16):
#     for i in range(32):
#         data.iloc[i * 16 + k, 9] = data.iloc[k , 9] + i * 3
#         data.iloc[i * 16 + k + 32 * 16, 9] = data.iloc[i * 3 + k, 9] + 20480
#         data.iloc[i * 16 + k + 32 * 32, 9] = data.iloc[i * 3 + k, 9] + 20480 * 2
# for i in range(95):
#     data.iloc[16 * i, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 1, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 5, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 8, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 9, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 10, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 12, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 13, 7] = int(i / 24) + 1
#     data.iloc[16 * i + 15, 7] = int(i / 24) + 1
#
# for i in range(120):
#     data.iloc[16 * i, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 1, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 5, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 8, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 9, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 10, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 12, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 13, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#     data.iloc[16 * i + 15, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
#
# for i in range(120):
#     if (i + 1) % 24 == 0:
#         data.iloc[0 + 16 * i, [12, 14]] = data.iloc[0, [12, 14]] + i * (data.iloc[16, [12, 14]] - data.iloc[0, [12, 14]]) + x
#         data.iloc[1 + 16 * i, [12, 14]] = data.iloc[1, [12, 14]] + i * (data.iloc[17, [12, 14]] - data.iloc[1, [12, 14]]) + x
#         data.iloc[5 + 16 * i, [12, 14]] = data.iloc[5, [12, 14]] + i * (data.iloc[21, [12, 14]] - data.iloc[5, [12, 14]]) + x
#         data.iloc[8 + 16 * i, [12, 14]] = data.iloc[8, [12, 14]] + i * (data.iloc[24, [12, 14]] - data.iloc[8, [12, 14]]) + x
#         data.iloc[9 + 16 * i, [12, 14]] = data.iloc[9, [12, 14]] + i * (data.iloc[25, [12, 14]] - data.iloc[9, [12, 14]]) + x
#         data.iloc[10 + 16 * i, [12, 14]] = data.iloc[10, [12, 14]] + i * (data.iloc[26, [12, 14]] - data.iloc[10, [12, 14]]) + x
#         data.iloc[12 + 16 * i, [12, 14]] = data.iloc[12, [12, 14]] + i * (data.iloc[28, [12, 14]] - data.iloc[12, [12, 14]]) + x
#         data.iloc[13 + 16 * i, [12, 14]] = data.iloc[13, [12, 14]] + i * (data.iloc[29, [12, 14]] - data.iloc[13, [12, 14]]) + x
#         data.iloc[15 + 16 * i, [12, 14]] = data.iloc[15, [12, 14]] + i * (data.iloc[31, [12, 14]] - data.iloc[15, [12, 14]]) + x
#     else:
#         data.iloc[0 + 16 * i, [12,14]] = data.iloc[0, [12,14]] + i * (data.iloc[16, [12, 14]] - data.iloc[0, [12, 14]])
#         data.iloc[1 + 16 * i, [12,14]] = data.iloc[1, [12,14]] + i * (data.iloc[17, [12, 14]] - data.iloc[1, [12, 14]])
#         data.iloc[5 + 16 * i, [12,14]] = data.iloc[5, [12,14]] + i * (data.iloc[21, [12, 14]] - data.iloc[5, [12, 14]])
#         data.iloc[8 + 16 * i, [12,14]] = data.iloc[8, [12,14]] + i * (data.iloc[24, [12, 14]] - data.iloc[8, [12, 14]])
#         data.iloc[9 + 16 * i, [12,14]] = data.iloc[9, [12,14]] + i * (data.iloc[25, [12, 14]] - data.iloc[9, [12, 14]])
#         data.iloc[10 + 16 * i, [12,14]] = data.iloc[10, [12,14]] + i * (data.iloc[26, [12, 14]] - data.iloc[10, [12, 14]])
#         data.iloc[12 + 16 * i, [12,14]] = data.iloc[12, [12,14]] + i * (data.iloc[28, [12, 14]] - data.iloc[12, [12, 14]])
#         data.iloc[13 + 16 * i, [12,14]] = data.iloc[13, [12,14]] + i * (data.iloc[29, [12, 14]] - data.iloc[13, [12, 14]])
#         data.iloc[15 + 16 * i, [12,14]] = data.iloc[15, [12,14]] + i * (data.iloc[31, [12, 14]] - data.iloc[15, [12, 14]])
print(data)
data.to_excel("data2.xlsx")
