import requests, pandas as pd, time, datetime

# --------------------------------------------------
# 1. 和浏览器完全一致的头/ cookie / ctoken
# --------------------------------------------------
import requests

cookies = {
    'acw_tc': '1a0c384b17562949591032000ee9e735f1e3d53535390db117b5d7468e608b',
    'RESUMEJSESSIONID': 'dff034c9-254a-48a4-86a6-1df3801cd953',
    'SERVERID': 'a232c76a369c895f1ebffb6ac1d6f78e|1756294974|1756294959',
}

headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'content-type': 'application/x-www-form-urlencoded',
    'origin': 'https://wecruit.hotjob.cn',
    'priority': 'u=1, i',
    'referer': 'https://wecruit.hotjob.cn/SU60769cec0dcad4510451cb0e/pb/social.html?postName=%E6%95%B0%E6%8D%AE',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
    # 'cookie': 'acw_tc=1a0c384b17562949591032000ee9e735f1e3d53535390db117b5d7468e608b; RESUMEJSESSIONID=dff034c9-254a-48a4-86a6-1df3801cd953; SERVERID=a232c76a369c895f1ebffb6ac1d6f78e|1756294974|1756294959',
}

params = {
    'iSaJAx': 'isAjax',
    'request_locale': 'zh_CN',
    't': '1756294997411',
}

company = "科大讯飞"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 30          # 你原始也是 10
    jobs = []
    while True:   
        data = {
            'isFrompb': 'true',
            'recruitType': '2',
            'pageSize': page_size,
            'currentPage': page_index,
            'postName': keyword,
        }

        response = requests.post(
            'https://wecruit.hotjob.cn/wecruit/positionInfo/listPosition/SU60769cec0dcad4510451cb0e',
            params=params,
            cookies=cookies,
            headers=headers,
            data=data,
        )

        data = response.json()
        postid_json = data.get("data", {}).get("pageForm", []).get("pageData", [])
        postid = [j["postId"] for j in postid_json]
        for id in postid:
            data = {
                'recruitType': '2',
                'postId': id
            }
            resp = requests.post(
                'https://wecruit.hotjob.cn/wecruit/positionInfo/listPositionDetail/SU60ba0f510dcad46f8599decf',
                params=params,
                cookies=cookies,
                headers=headers,
                data=data,
            )
            print(f"正在获取职位ID={id}的详情...")  
            data = resp.json()
            job_detail = data.get("data", {})
            jobs.append(job_detail)
        
        if not postid_json:
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
            "职位id": j["postId"],
            "职位名称": j["postName"],
            "工作地点": j["workPlaceStr"],
            "职位描述": j.get("workContent", ""),
            "职位要求": j.get("serviceCondition", ""),
            "职位类别ID": '',
            "职位类别名称": j["department"],
            "发布时间": j["publishDate"][:10]
        }
        for tag,j in jobs_with_tag
    ]
# 直接写 Excel
df = pd.DataFrame(rows)
df.to_excel(f"招聘JD\{company}.xlsx", index=False)

print(f"\n共 {len(jobs_with_tag)} 条职位，已保存到 {company}.xlsx")
