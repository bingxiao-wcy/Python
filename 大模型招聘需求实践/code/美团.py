import requests, json, csv
import time
import pandas as pd
import datetime
import requests, urllib.parse, time

# cookies = {
#     '_lxsdk_cuid': '198d9a59f15c8-0122558563eb1e8-26011051-1bcab9-198d9a59f15c8',
#     'weixinType': '1',
#     'com.sankuai.recruitment.official.website_strategy': 'VM4LVFMDrs',
#     'logan_session_token': 'u2uhv6zlev4esgoei4uu',
#     '_lx_utm': 'utm_source%3Dbing%26utm_medium%3Dorganic',
#     '_lxsdk_s': '198e13f8ae6-383-751-a1a%7C%7C12',
# }

# headers = {
#     'accept': 'application/json',
#     'accept-language': 'zh-CN,zh;q=0.9',
#     'content-type': 'application/json',
#     'origin': 'https://zhaopin.meituan.com',
#     'priority': 'u=1, i',
#     'referer': 'https://zhaopin.meituan.com/web/social?keyword=%E5%A4%A7%E6%95%B0%E6%8D%AE',
#     'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
#     'sec-ch-ua-mobile': '?1',
#     'sec-ch-ua-platform': '"Android"',
#     'sec-fetch-dest': 'empty',
#     'sec-fetch-mode': 'cors',
#     'sec-fetch-site': 'same-origin',
#     'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
#     'x-requested-with': 'XMLHttpRequest',
#     # 'cookie': '_lxsdk_cuid=198d9a59f15c8-0122558563eb1e8-26011051-1bcab9-198d9a59f15c8; weixinType=1; com.sankuai.recruitment.official.website_strategy=VM4LVFMDrs; logan_session_token=u2uhv6zlev4esgoei4uu; _lx_utm=utm_source%3Dbing%26utm_medium%3Dorganic; _lxsdk_s=198e13f8ae6-383-751-a1a%7C%7C12',
# }

cookies = {
    '_lxsdk_cuid': '198d9a59f15c8-0122558563eb1e8-26011051-1bcab9-198d9a59f15c8',
    'weixinType': '1',
    'logan_session_token': 'i66teop25hg5oyi0qj6v',
    '_lx_utm': 'utm_source%3Dbing%26utm_medium%3Dorganic',
    '_lxsdk_s': '1995a6731c5-929-45d-63b%7C%7C18',
}

headers = {
    'accept': 'application/json',
    'accept-language': 'zh-CN,zh;q=0.9',
    'content-type': 'application/json',
    'origin': 'https://zhaopin.meituan.com',
    'priority': 'u=1, i',
    'referer': 'https://zhaopin.meituan.com/web/social?keyword=%E5%A4%A7%E6%A8%A1%E5%9E%8B',
    'sec-ch-ua': '"Chromium";v="140", "Not=A?Brand";v="24", "Google Chrome";v="140"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Mobile Safari/537.36',
    'x-requested-with': 'XMLHttpRequest',
    # 'cookie': '_lxsdk_cuid=198d9a59f15c8-0122558563eb1e8-26011051-1bcab9-198d9a59f15c8; weixinType=1; logan_session_token=i66teop25hg5oyi0qj6v; _lx_utm=utm_source%3Dbing%26utm_medium%3Dorganic; _lxsdk_s=1995a6731c5-929-45d-63b%7C%7C18',
}

company = "美团"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 30          # 你原始也是 10

    while True:
        json_data = {
            'page': {
                'pageNo': page_index,
                'pageSize': page_size,
            },
            'jobShareType': '1',
            'keywords': keyword,
            'cityList': [],
            'department': [],
            'jfJgList': [],
            'jobType': [
                {
                    'code': '3',
                    'subCode': [],
                },
            ],
            'u_query_id': '63386b9288b8811eef04317c831c2d42',
            'r_query_id': '175612570957845632955',
        }

        resp = requests.post(
            'https://zhaopin.meituan.com/api/official/job/getJobList',
            cookies=cookies,
            headers=headers,
            json=json_data,
        )
        resp.raise_for_status()
        data = resp.json()
        jobs = data.get('data', {}).get('list', [])

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
        "职位id": j["jobUnionId"],
        "职位名称": j["name"],
        "工作地点": [i.get('name','') for i in j["cityList"]],
        "职位描述": j.get("highLight", ""),
        "职位要求": j.get("jobDuty", ""),
        "职位类别ID": '',
        "职位类别名称": j["jobFamilyGroup"],
        "发布时间": datetime.datetime.fromtimestamp(j.get("refreshTime", "") / 1000).strftime('%Y-%m-%d')
    }
    for tag,j in jobs_with_tag
]

# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")

