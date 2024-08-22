# /**
#  * @file main.py
#  * @author 恒星不见
#  * @brief 四六级查询
#  * @version 0.1
#  * @date 2024-02-27
#  *
#  * @copyright Copyright (c) 2024
#  *
#  */
import time
import requests
import pandas as pd


def cet_query(km, name, id):
    global data
    url = "https://cachecloud.neea.cn/api/latest/results/cet"
    header = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36 Edg/115.0.1901.203",
              "Referer": "https://cjcx.neea.edu.cn/"}
    params = {"km": km, "xm": name, "no": id, "source": "pc"}
    response = requests.get(url, headers=header, params=params)
    json_data = response.json()
    try:
        data.iloc[i, 2] = json_data['score']  # 总分
        data.iloc[i, 3] = json_data['sco_lc']  # 听力
        data.iloc[i, 4] = json_data['sco_rd']  # 阅读
        data.iloc[i, 5] = json_data['sco_wt']  # 写作
        data.iloc[i, 6] = 2*km+2
        print(name, 2*km+2, int(json_data['score']), int(json_data['sco_lc']), int(
            json_data['sco_rd']), int(json_data['sco_wt']))
    except:
        print(name, 2*km+2, 'None')
        return 0


data = pd.read_excel('./data/data.xlsx')  # data.xlsx格式应为: 第一列为姓名，第二列为身份证号

data.insert(2, '成绩', '')
data.insert(3, '听力', '')
data.insert(4, '阅读', '')
data.insert(5, '写作', '')
data.insert(6, '科目', '')
nums = len(data)  # 查询条数
for i in range(nums):
    if cet_query(2, data.iloc[i, 0], data.iloc[i, 1]) == 0:
        cet_query(1, data.iloc[i, 0], data.iloc[i, 1])
    time.sleep(0.5)

data.to_excel('./result/grade.xlsx', index=None, header=True)
print("导出完毕.")
