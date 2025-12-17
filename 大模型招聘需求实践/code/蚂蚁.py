import requests, pandas as pd, time, datetime

# --------------------------------------------------
# 1. 和浏览器完全一致的头/ cookie / ctoken
# --------------------------------------------------

cookies = {
    'receive-cookie-deprecation': '1',
    '_CHIPS-ALIPAYJSESSIONID': 'dCGPZQLdPqiO1NRxoUawq8TbEUsuQQFcternbase',
    'ALIPAYJSESSIONID': 'dCGPZQLdPqiO1NRxoUawq8TbEUsuQQFcternbase',
    'ctoken': 'bigfish_ctoken_1a00b0jjlh',
    'SESSION': 'OTI2MTVGQjdGNzQ5NUY4ODk3QUVEMDExOEFGRDNGNUM=',
    'spanner': 'tumLxvjfhxZFNAA/urXj01Ti/gLWiZ/T4EJoL7C0n0A=',
    'tfstk': 'gLznEs1NbkoC0TLBx_0CAhHf36sOR2gSH8L-e4HPbAk6y6wROucrI8mJyyn8EUPSByHL92ykq-DgJ6NKyUkrRPEpyvBIr7yxrtBAkZFQOmgPHtesXx5-ujJyLLHE_5u-ZpzxlNFQO4O14cmCJWOlL6pqL4yr7clmZ48y8X5iQjGjzeoy8hJZCburzvly7Nljiburz8PNsbMZT4ozU55iNAkrB8tSnJwY7TYXg6zUmpzstDD4TvVLiP5njHNrIUY7L0nn3ZHMzUziT7b6ykLc2xrjpmHbQNLnr5luImEhne2uaSZrjzvF9-PzdP0IogCImlrTDDqGLUDTJbmqzcAyz50iC-4tuZJi1lPL4ywHZUlQJrnohcfyP03akmr4KQ1b_2lzFmaRheMUaSabDqbHhXqaiVjyuh-2IMLS__UwV3iE1fDYGg1b5rehOafGsn7ZYfGsh1fMV3iE1fDAs1xV7Dls1x1..',
}

headers = {
    'accept': 'application/json',
    'accept-language': 'zh-CN,zh;q=0.9',
    'content-type': 'application/json;charset=UTF-8',
    'front-user-id': '6d0857df-765c-4541-91e6-a519beb87de452',
    'origin': 'https://talent.antgroup.com',
    'priority': 'u=1, i',
    'referer': 'https://talent.antgroup.com/',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
    # 'cookie': 'receive-cookie-deprecation=1; _CHIPS-ALIPAYJSESSIONID=dCGPZQLdPqiO1NRxoUawq8TbEUsuQQFcternbase; ALIPAYJSESSIONID=dCGPZQLdPqiO1NRxoUawq8TbEUsuQQFcternbase; ctoken=bigfish_ctoken_1a00b0jjlh; SESSION=OTI2MTVGQjdGNzQ5NUY4ODk3QUVEMDExOEFGRDNGNUM=; spanner=tumLxvjfhxZFNAA/urXj01Ti/gLWiZ/T4EJoL7C0n0A=; tfstk=gLznEs1NbkoC0TLBx_0CAhHf36sOR2gSH8L-e4HPbAk6y6wROucrI8mJyyn8EUPSByHL92ykq-DgJ6NKyUkrRPEpyvBIr7yxrtBAkZFQOmgPHtesXx5-ujJyLLHE_5u-ZpzxlNFQO4O14cmCJWOlL6pqL4yr7clmZ48y8X5iQjGjzeoy8hJZCburzvly7Nljiburz8PNsbMZT4ozU55iNAkrB8tSnJwY7TYXg6zUmpzstDD4TvVLiP5njHNrIUY7L0nn3ZHMzUziT7b6ykLc2xrjpmHbQNLnr5luImEhne2uaSZrjzvF9-PzdP0IogCImlrTDDqGLUDTJbmqzcAyz50iC-4tuZJi1lPL4ywHZUlQJrnohcfyP03akmr4KQ1b_2lzFmaRheMUaSabDqbHhXqaiVjyuh-2IMLS__UwV3iE1fDYGg1b5rehOafGsn7ZYfGsh1fMV3iE1fDAs1xV7Dls1x1..',
}

params = {
    'ctoken': 'bigfish_ctoken_1a00b0jjlh',
}

company = "蚂蚁集团"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 30          # 你原始也是 10

    while True:
        json_data = {
            "key": keyword,
            "regions": "",
            "categories": "",
            "subCategories": "",
            "bgCode": "",
            "socialQrCode": "",
            "pageIndex": page_index,
            "pageSize": page_size,
            "channel": "group_official_site",
            "language": "zh",
        }

        resp = requests.post(
            "https://hrcareersweb.antgroup.com/api/social/position/search",
            params=params,
            headers=headers,
            cookies=cookies,
            json=json_data,
            timeout=10
        )
        resp.raise_for_status()
        data = resp.json()
        jobs = data.get("content", {})
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
        "工作地点": j["workLocations"],
        "职位描述": j.get("description", ""),
        "职位要求": j.get("requirement", ""),
        "职位类别ID": '',
        "职位类别名称": j["categories"],
        "发布时间": j.get("publishTime", "")[0:10]
    }
    for tag,j in jobs_with_tag
]

# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
