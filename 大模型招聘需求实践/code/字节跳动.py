import requests, json, csv
import time
import pandas as pd
import datetime

# 1. 两个关键词各自对应的完整 Cookie 与 token


token = "RRpdaQrvXIUGVoMeLJlVPNdIRUPV18ofKom38o-DZU0="

cookie = {
    'ttwid': '1%7CccjK9OrG4Rgw4e7tL91g129qbXbNmWH6wO2S9Bmm8qY%7C1755779051%7C95f2be7a67b21105e7ad71b1f7e22562ff815201e9e1bace39688c81bd2bfd6c',
    'locale': 'zh-CN',
    's_v_web_id': 'verify_melfm6y1_dG6F7AN1_trKW_4ttV_AXK2_E2xarKflBwLK',
    'device-id': '7541028722909693441',
    'channel': 'office',
    'tea_uid': '7541028591360149038',
    'atsx-csrf-token': token[:-1] + "%3D",
    'platform': 'h5',
}
headers = {
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0',
    'Referer': 'https://jobs.bytedance.com/experienced/position',
    'x-csrf-token': token
}   

company = "字节跳动"

# 2. 统一的 fetch_all，内部根据 keyword 取对应配置
def fetch_all(keyword: str):

    all_jobs = []
    page  = 0          # 从第 0 页开始
    limit = 30         # 每页条数

    while True:
        payload = {
            "keyword": keyword,
            "limit": limit,
            "offset": page * limit,   # 关键点：动态偏移
            "job_category_id_list": [],
            "portal_type": 2,
            "portal_entrance": 1,
        }

        resp = requests.post(
            "https://jobs.bytedance.com/api/v1/search/job/posts",
            json=payload,
            headers=headers,
            cookies=cookie,
            timeout=10
        )
        resp.raise_for_status()
        data = resp.json()
        jobs = data.get("data", {}).get("job_post_list", [])
        if not jobs: 
            break  
        print(f"关键词 '{keyword}' - 第 {page+1} 页，获取到 {len(jobs)} 条")
        all_jobs.extend(jobs)
        page += 1
        time.sleep(0.5)

    return all_jobs

big_data_jobs = fetch_all("大数据")
big_model_jobs = fetch_all("大模型")

# 把结果和关键词一起打包
jobs_with_tag = [("大数据", j) for j in big_data_jobs] + \
                [("大模型", j) for j in big_model_jobs]

# 整理字段，新增“分类”列
rows = [
    {
        "分类": tag,
        "职位id": j["id"],
        "职位名称": j["title"],
        "工作地点": j["city_info"]["name"],
        "职位描述": j.get("description", ""),
        "职位要求": j.get("requirement", ""),
        "职位类别ID": j["job_category"]["id"],
        "职位类别名称": j["job_category"]["name"],
        "发布时间": datetime.datetime.fromtimestamp(j.get("publish_time", "") / 1000).strftime('%Y-%m-%d') 
    }
    for tag,j in jobs_with_tag
]
# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
