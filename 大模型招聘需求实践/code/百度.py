import requests
import pandas as pd
import csv
import time

company = "百度"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 10          # 你原始也是 10

    while True:
        url = 'https://talent.baidu.com/httservice/getPostListNew'
        data = {
            'recruitType': 'SOCIAL',
            'pageSize': page_size,
            'keyWord': keyword,
            'curPage': page_index,
            'projectType': ''
        }
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36...',
            'Referer': 'https://talent.baidu.com/jobs/social-list'
        }

        resp = requests.post(url, data=data, headers=headers)
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
        "职位id": j["jobId"],
        "职位名称": j["name"],
        "工作地点": j["workPlace"],
        "职位描述": j.get("workContent", ""),
        "职位要求": j.get("serviceCondition", ""),
        "职位类别ID": '',
        "职位类别名称": j["postType"],
        "发布时间": j['updateDate']
    }
    for tag,j in jobs_with_tag
]
# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
