import  pandas  as pd
data = pd.read_excel('data.xlsx')
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 1000)
x = 22
for i in range(120):
    for j in range(16):
        data.iloc[j + 16 * i, 1:] = data.iloc[j, 1:]

for i in range(120):
    data.iloc[16 * i, 7] = int(i / 24) + 1
    data.iloc[16 * i + 1, 7] = int(i / 24) + 1
    data.iloc[16 * i + 5, 7] = int(i / 24) + 1
    data.iloc[16 * i + 8, 7] = int(i / 24) + 1
    data.iloc[16 * i + 9, 7] = int(i / 24) + 1
    data.iloc[16 * i + 10, 7] = int(i / 24) + 1
    data.iloc[16 * i + 12, 7] = int(i / 24) + 1
    data.iloc[16 * i + 13, 7] = int(i / 24) + 1
    data.iloc[16 * i + 15, 7] = int(i / 24) + 1

for i in range(120):
    data.iloc[16 * i, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 1, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 5, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 8, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 9, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 10, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 12, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 13, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1
    data.iloc[16 * i + 15, 10] = i - (data.iloc[16 * i, 7] - 1) * 24 + 1

for i in range(120):
    if (i + 1) % 24 == 0:
        data.iloc[0 + 16 * i, [12, 14]] = data.iloc[0, [12, 14]] + i * (data.iloc[16, [12, 14]] - data.iloc[0, [12, 14]]) + x
        data.iloc[1 + 16 * i, [12, 14]] = data.iloc[1, [12, 14]] + i * (data.iloc[17, [12, 14]] - data.iloc[1, [12, 14]]) + x
        data.iloc[5 + 16 * i, [12, 14]] = data.iloc[5, [12, 14]] + i * (data.iloc[21, [12, 14]] - data.iloc[5, [12, 14]]) + x
        data.iloc[8 + 16 * i, [12, 14]] = data.iloc[8, [12, 14]] + i * (data.iloc[24, [12, 14]] - data.iloc[8, [12, 14]]) + x
        data.iloc[9 + 16 * i, [12, 14]] = data.iloc[9, [12, 14]] + i * (data.iloc[25, [12, 14]] - data.iloc[9, [12, 14]]) + x
        data.iloc[10 + 16 * i, [12, 14]] = data.iloc[10, [12, 14]] + i * (data.iloc[26, [12, 14]] - data.iloc[10, [12, 14]]) + x
        data.iloc[12 + 16 * i, [12, 14]] = data.iloc[12, [12, 14]] + i * (data.iloc[28, [12, 14]] - data.iloc[12, [12, 14]]) + x
        data.iloc[13 + 16 * i, [12, 14]] = data.iloc[13, [12, 14]] + i * (data.iloc[29, [12, 14]] - data.iloc[13, [12, 14]]) + x
        data.iloc[15 + 16 * i, [12, 14]] = data.iloc[15, [12, 14]] + i * (data.iloc[31, [12, 14]] - data.iloc[15, [12, 14]]) + x
    else:
        data.iloc[0 + 16 * i, [12,14]] = data.iloc[0, [12,14]] + i * (data.iloc[16, [12, 14]] - data.iloc[0, [12, 14]])
        data.iloc[1 + 16 * i, [12,14]] = data.iloc[1, [12,14]] + i * (data.iloc[17, [12, 14]] - data.iloc[1, [12, 14]])
        data.iloc[5 + 16 * i, [12,14]] = data.iloc[5, [12,14]] + i * (data.iloc[21, [12, 14]] - data.iloc[5, [12, 14]])
        data.iloc[8 + 16 * i, [12,14]] = data.iloc[8, [12,14]] + i * (data.iloc[24, [12, 14]] - data.iloc[8, [12, 14]])
        data.iloc[9 + 16 * i, [12,14]] = data.iloc[9, [12,14]] + i * (data.iloc[25, [12, 14]] - data.iloc[9, [12, 14]])
        data.iloc[10 + 16 * i, [12,14]] = data.iloc[10, [12,14]] + i * (data.iloc[26, [12, 14]] - data.iloc[10, [12, 14]])
        data.iloc[12 + 16 * i, [12,14]] = data.iloc[12, [12,14]] + i * (data.iloc[28, [12, 14]] - data.iloc[12, [12, 14]])
        data.iloc[13 + 16 * i, [12,14]] = data.iloc[13, [12,14]] + i * (data.iloc[29, [12, 14]] - data.iloc[13, [12, 14]])
        data.iloc[15 + 16 * i, [12,14]] = data.iloc[15, [12,14]] + i * (data.iloc[31, [12, 14]] - data.iloc[15, [12, 14]])
print(data)
data.to_excel("data2.xlsx")
