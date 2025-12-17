import requests, pandas as pd, time, datetime

# --------------------------------------------------
# 1. 和浏览器完全一致的头/ cookie / ctoken
# --------------------------------------------------

cookies = {
    'cna': 'w0ozITlofgECAXJWZhIj6sCW',
    'xlly_s': '1',
    'prefered-lang': 'zh',
    'XSRF-TOKEN': 'ffe7874f-ae9e-45b2-8e56-086d0e81b587',
    'SESSION': 'QkY3N0RCRjlFRDZDMTUyMzM2QkI3QjExMDBERkFCQzY=',
    'tfstk': 'gn3-5kXCAKvljrXPwvxctgBss3RmJncPcYl1-J2lAxHxKY3nrbinpkhQnQznqHzLvv2mr8cuPvIKB-3kab4hRJwjO8boqJYKpxeY-b2uK_oqIf_or3zhJzzURdvMSFVoayzQ2cWzITrb_7ThPg_7zrg7qORBSFcrN_VWsjTi-VpeQWeQRuZClSw49gsSdJ1XMWNhFM_SdjOY3WC5PJZ7GsNa6gaQRvGXMWybNyZSdjOYT-wIBXR3hMyCJ0p5mxlJ4L75PqF82J_3QwK9tWr8CbZVBc04wOysN-QCRRKtWtlxZLQZizuiB5DklaHtOfoQD2BXBJuxG4iIgTpQgDMul846VsUozbi_AV9FAAU-plg7D6_sabw4kSg6tGymk8DIyoOGWk4jSlaSmnJuxzFtd4kR9wMsif0z02pOhJou_rNK8CsLpksr5V0OSfXgBWjWMIIFYuNV_wjJ6QapjKFYIIxOYMruiSeMMIIFYuN4MRADXMSUqSf..',
    'isg': 'BBISwWFVrgDzM9J9UBzeEr7FY9j0Ixa9gsNujdxrakUf77fpzrOezJVJWktT345V',
}

headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'bx-v': '2.5.11',
    'content-type': 'application/json',
    'origin': 'https://aidc-jobs.alibaba.com',
    'priority': 'u=1, i',
    'referer': 'https://aidc-jobs.alibaba.com/off-campus/position-list',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
    # 'cookie': 'cna=w0ozITlofgECAXJWZhIj6sCW; xlly_s=1; prefered-lang=zh; XSRF-TOKEN=ffe7874f-ae9e-45b2-8e56-086d0e81b587; SESSION=QkY3N0RCRjlFRDZDMTUyMzM2QkI3QjExMDBERkFCQzY=; tfstk=gn3-5kXCAKvljrXPwvxctgBss3RmJncPcYl1-J2lAxHxKY3nrbinpkhQnQznqHzLvv2mr8cuPvIKB-3kab4hRJwjO8boqJYKpxeY-b2uK_oqIf_or3zhJzzURdvMSFVoayzQ2cWzITrb_7ThPg_7zrg7qORBSFcrN_VWsjTi-VpeQWeQRuZClSw49gsSdJ1XMWNhFM_SdjOY3WC5PJZ7GsNa6gaQRvGXMWybNyZSdjOYT-wIBXR3hMyCJ0p5mxlJ4L75PqF82J_3QwK9tWr8CbZVBc04wOysN-QCRRKtWtlxZLQZizuiB5DklaHtOfoQD2BXBJuxG4iIgTpQgDMul846VsUozbi_AV9FAAU-plg7D6_sabw4kSg6tGymk8DIyoOGWk4jSlaSmnJuxzFtd4kR9wMsif0z02pOhJou_rNK8CsLpksr5V0OSfXgBWjWMIIFYuNV_wjJ6QapjKFYIIxOYMruiSeMMIIFYuN4MRADXMSUqSf..; isg=BBISwWFVrgDzM9J9UBzeEr7FY9j0Ixa9gsNujdxrakUf77fpzrOezJVJWktT345V',
}

params = {
    '_csrf': 'ffe7874f-ae9e-45b2-8e56-086d0e81b587',
}

company = "阿里国际"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 30          # 你原始也是 10

    while True:
        json_data = {
            'channel': 'group_official_site',
            'language': 'zh',
            'batchId': '',
            'categories': '',
            'deptCodes': [],
            'key': keyword,
            'pageIndex': page_index,
            'pageSize': page_size,
            'regions': '',
            'subCategories': '',
        }

        resp = requests.post(
            'https://aidc-jobs.alibaba.com/position/search',
            params=params,
            cookies=cookies,
            headers=headers,
            json=json_data,
        )
        resp.raise_for_status()
        data = resp.json()
        jobs = data.get("content", {}).get('datas', [])
        
        if not jobs:
            break
        print(f"关键词 '{keyword}' - 第 {page_index} 页，获取到 {len(jobs)} 条")
        all_jobs.extend(jobs)
        page_index += 1
        time.sleep(0.5)

    return all_jobs

# print(jobs['id','name','categories','publishTime','workLocations','requirement','description'])

# 全量抓取并写 CSV
big_data_jobs   = fetch_all_ant("数据")
big_model_jobs   = fetch_all_ant("大模型")

# 把结果和关键词一起打包
jobs_with_tag = [("数据", j) for j in big_data_jobs] + \
                [("大模型", j) for j in big_model_jobs]

# 把字段整理成列表字典
rows = [
    {
        "分类": tag,
        "职位id": j["id"],
        "职位名称": j["name"],
        "工作地点": j["workLocations"],
        "职位描述": j.get("description", ""),
        "职位要求": j.get("requirement", ""),
        "职位类别ID": j["level"],
        "职位类别名称": j["categories"],
        "发布时间": datetime.datetime.fromtimestamp(j.get("publishTime", "") / 1000).strftime('%Y-%m-%d')
    }
    for tag,j in jobs_with_tag
]

# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
