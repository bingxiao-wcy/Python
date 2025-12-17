import requests, pandas as pd, time, datetime

# --------------------------------------------------
# 1. 和浏览器完全一致的头/ cookie / ctoken
# --------------------------------------------------
cookies = {
    '__jda': '176729966.1756205851264334398740.1756205851.1756205851.1756205851.1',
    '__jdc': '176729966',
    '__jdv': '176729966|cn.bing.com|-|referral|-|1756205851266',
    '__jdu': '1756205851264334398740',
    '__jdb': '176729966.2.1756205851264334398740|1.1756205851',
    'JSESSIONID': 'CD9762F132A9DE9C67063C2709D25190.s1',
}

headers = {
    'accept': '*/*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'origin': 'https://zhaopin.jd.com',
    'priority': 'u=1, i',
    'referer': 'https://zhaopin.jd.com/web/job/job_info_list/3',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
    'x-requested-with': 'XMLHttpRequest',
    # 'cookie': '__jda=176729966.1756205851264334398740.1756205851.1756205851.1756205851.1; __jdc=176729966; __jdv=176729966|cn.bing.com|-|referral|-|1756205851266; __jdu=1756205851264334398740; __jdb=176729966.2.1756205851264334398740|1.1756205851; JSESSIONID=CD9762F132A9DE9C67063C2709D25190.s1',
}

company = "京东"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    data = {
        'pageIndex': 1,
        'pageSize': 30,
        'workCityJson': '[]',
        'jobTypeJson': '[]',
        'jobSearch': keyword
    }
    resp = requests.post(
        'https://zhaopin.jd.com/web/job/job_list',
        cookies=cookies,
        headers=headers,
        json=data
    )
    resp.raise_for_status()
    data = resp.json()
    jobs = data
    print(f"关键词 '{keyword}' - 获取到 {len(jobs)} 条")
    all_jobs.extend(jobs)
    time.sleep(0.5)

    return all_jobs

# print(jobs['id','name','categories','publishTime','workLocations','requirement','description'])

# 全量抓取并写 CSVd
big_data_jobs   = fetch_all_ant("大数据")
big_model_jobs   = fetch_all_ant("大模型")

# 把结果和关键词一起打包
jobs_with_tag = [("大数据", j) for j in big_data_jobs] + \
                [("大模型", j) for j in big_model_jobs]

# 把字段整理成列表字典
rows = [
    {
        "分类": tag,
        "职位id": j["positionId"],
        "职位名称": j["positionName"],
        "工作地点": j["workCity"],
        "职位描述": j.get("workContent", ""),
        "职位要求": j.get("qualification", ""),
        "职位类别ID": '',
        "职位类别名称": j["jobType"],
        "发布时间": j.get("formatPublishTime", "")
    }
    for tag,j in jobs_with_tag
]

# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
