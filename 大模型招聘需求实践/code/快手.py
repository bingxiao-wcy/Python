import requests, pandas as pd, time, datetime

# --------------------------------------------------
# 1. 和浏览器完全一致的头/ cookie / ctoken
# --------------------------------------------------
import requests

cookies = {
    'aliyungf_tc': 'f31656a694f76991dbc2f37528c8adc1133c78c8c00c5e119c8c3536c72006b7',
    'accessproxy_session': '547be30c-d947-4c90-a926-8b6a7ca17b9c',
    'apdid': 'a6c8924f-abfe-444d-aa8e-3ed4dfd6ac6beec37592f11ddb179b5162e1896af049:1756213461:1',
    '_did': 'web_4371132533C60B6E',
}

headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Connection': 'keep-alive',
    'Referer': 'https://zhaopin.kuaishou.cn/recruit/e/h5/',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    # 'Cookie': 'aliyungf_tc=f31656a694f76991dbc2f37528c8adc1133c78c8c00c5e119c8c3536c72006b7; accessproxy_session=547be30c-d947-4c90-a926-8b6a7ca17b9c; apdid=a6c8924f-abfe-444d-aa8e-3ed4dfd6ac6beec37592f11ddb179b5162e1896af049:1756213461:1; _did=web_4371132533C60B6E',
}



company = "快手"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 30          # 你原始也是 10

    while True:
        params = {
            'name': keyword,
            'pageNum': page_index,
            'pageSize': page_size,
            'positionNatureCode': 'C001',
            'recruitProject': 'socialr',
        }

        resp = requests.get(
            'https://zhaopin.kuaishou.cn/recruit/e/api/v1/open/positions/simple',
            params=params,
            cookies=cookies,
            headers=headers,
        )
        resp.raise_for_status()
        data = resp.json()
        jobs = data.get('result', {}).get('list', [])
        
        if not jobs:
            break
        print(f"关键词 '{keyword}' - 第 {page_index} 页，获取到 {len(jobs)} 条")
        all_jobs.extend(jobs)
        page_index += 1
        time.sleep(0.5)

    return all_jobs

# print(jobs['id','name','categories','publishTime','workLocations','requirement','description'])

# 全量抓取并写 CSV
big_data_jobs   = fetch_all_ant("大数据")
big_model_jobs   = fetch_all_ant("大模型")

# 把结果和关键词一起打包
jobs_with_tag = [("大数据", j) for j in big_data_jobs] + \
                [("大模型", j) for j in big_model_jobs]

# 把字段整理成列表字典

rows = [
    {
        "分类": tag,
        "职位id": '',
        "职位名称": j["name"],
        "工作地点": j["workLocationCode"],
        "职位描述": j.get("description", ""),
        "职位要求": j.get("positionDemand", ""),
        "职位类别ID": '',
        "职位类别名称": '',
        "发布时间": ''
    }
    for tag,j in jobs_with_tag
]

# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
