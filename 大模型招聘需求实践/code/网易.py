import requests, pandas as pd, time, datetime

# --------------------------------------------------
# 1. 和浏览器完全一致的头/ cookie / ctoken
# --------------------------------------------------
import requests

cookies = {
    'hb_MA-8E16-605C3AFFE11F_source': 'cn.bing.com',
    'userName': '',
    'accountType': '',
    'JSESSIONID': '5298900FEB5AD4BD30D4CEC61C58A752',
}

headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'authtype': 'ursAuth',
    'content-type': 'application/json;charset=UTF-8',
    'language': 'zh',
    'origin': 'https://hr.163.com',
    'priority': 'u=1, i',
    'referer': 'https://hr.163.com/job-list.html?currentPage=1&pageSize=10&postType=36&lang=zh',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
    'x-ehr-uuid': '63167186-98f6-42f8-89bf-3a693f4592',
    # 'cookie': 'hb_MA-8E16-605C3AFFE11F_source=cn.bing.com; userName=; accountType=; JSESSIONID=5298900FEB5AD4BD30D4CEC61C58A752',
}

company = "网易"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 30          # 你原始也是 10

    while True:
        json_data = {
            'keyword': keyword,
            'currentPage': page_index,
            'pageSize': page_size,
        }

        resp = requests.post(
            'https://hr.163.com/api/hr163/position/queryPage',
            cookies=cookies,
            headers=headers,
            json=json_data
        )
        resp.raise_for_status()
        data = resp.json()
        jobs = data.get("data", {}).get("list", [])
        
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
        "职位id": j["id"],
        "职位名称": j["name"],
        "工作地点": j["workPlaceNameList"],
        "职位描述": j.get("description", ""),
        "职位要求": j.get("requirement", ""),
        "职位类别ID": '',
        "职位类别名称": j["firstPostTypeName"],
        "发布时间": ''
    }
    for tag,j in jobs_with_tag
]
# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
