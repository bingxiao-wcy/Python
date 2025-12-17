import requests, pandas as pd, time, datetime
import os

cookies = {
    'NOWCODERCLINETID': '9EF6426FC53447A2E15E195AA2AD160E',
    'NOWCODERUID': 'd8b9922971ad496280bdce955c442d24',
    'Hm_lvt_a808a1326b6c06c437de769d1b85b870': '1756602683',
    'HMACCOUNT': '1266A79A200DBD9A',
    'gr_user_id': 'e3af9d5b-73f5-4652-badd-3bad8f73d1f9',
    'c196c3667d214851b11233f5c17f99d5_gr_session_id': 'c0517fc0-3c5e-45c8-bac4-fae086b21b32',
    't': '38DE3985346C651CC96D26A7825C3F7B',
    'c196c3667d214851b11233f5c17f99d5_gr_last_sent_sid_with_cs1': 'c0517fc0-3c5e-45c8-bac4-fae086b21b32',
    'c196c3667d214851b11233f5c17f99d5_gr_last_sent_cs1': '482206722',
    'fromPut': 'h5_discuss_cmt',
    'c196c3667d214851b11233f5c17f99d5_gr_session_id_c0517fc0-3c5e-45c8-bac4-fae086b21b32': 'true',
    'Hm_lpvt_a808a1326b6c06c437de769d1b85b870': '1756606063',
    'c196c3667d214851b11233f5c17f99d5_gr_cs1': '482206722',
}

headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'content-type': 'application/json',
    'origin': 'https://www.nowcoder.com',
    'priority': 'u=1, i',
    'referer': 'https://www.nowcoder.com/',
    'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36',
    'x-requested-with': 'XMLHttpRequest',
    # 'cookie': 'NOWCODERCLINETID=9EF6426FC53447A2E15E195AA2AD160E; NOWCODERUID=d8b9922971ad496280bdce955c442d24; Hm_lvt_a808a1326b6c06c437de769d1b85b870=1756602683; HMACCOUNT=1266A79A200DBD9A; gr_user_id=e3af9d5b-73f5-4652-badd-3bad8f73d1f9; c196c3667d214851b11233f5c17f99d5_gr_session_id=c0517fc0-3c5e-45c8-bac4-fae086b21b32; t=38DE3985346C651CC96D26A7825C3F7B; c196c3667d214851b11233f5c17f99d5_gr_last_sent_sid_with_cs1=c0517fc0-3c5e-45c8-bac4-fae086b21b32; c196c3667d214851b11233f5c17f99d5_gr_last_sent_cs1=482206722; fromPut=h5_discuss_cmt; c196c3667d214851b11233f5c17f99d5_gr_session_id_c0517fc0-3c5e-45c8-bac4-fae086b21b32=true; Hm_lpvt_a808a1326b6c06c437de769d1b85b870=1756606063; c196c3667d214851b11233f5c17f99d5_gr_cs1=482206722',
}

def fetch_all_ant(keyword: str, conf: dict) -> list[dict]:
    all_jobs = {}
    page_index = 1
    print(f"{keyword}: 公司列表={conf['companyList']}, 职位ID={conf['jobid']}")
    while page_index <= 100:  # 最多 100 页
        # 数据
        params = {
            '_': '1756611397391',
        }
        json_data = {
            'companyList':conf['companyList'],
            'jobId': conf['jobid'],
            'level': 2,
            'order': 10,
            'page': page_index,
            'isNewJob': True,
        }

        resp = requests.post(
            'https://gw-c.nowcoder.com/api/sparta/job-experience/experience/job/list',
            params=params,
            cookies=cookies,
            headers=headers,
            json=json_data,
        )

        resp.raise_for_status()
        jobs = resp.json()
        # print(jobs)
        if not jobs:
            break
        print(f"关键词 '{keyword}' - 第 {page_index} 页，获取到 {len(jobs)} 条")
        all_jobs[page_index] = jobs
        page_index += 1
        time.sleep(0.5)
    return all_jobs

def extract_row(js: dict) -> dict:
    result = {}
    # 先判断是否存在momentData
    if "momentData" in js:
        result = {
            "contentId": js.get("contentId", ""),
            "contentType": js.get("contentType", ""),
            "userId": js["userBrief"]["userId"],
            "nickname": js["userBrief"]["nickname"],
            "workTime": js["userBrief"]["workTime"],
            "educationInfo": js["userBrief"]["educationInfo"],
            "secondMajorName": js["userBrief"]["secondMajorName"],
            "title": js["momentData"]["title"],
            "ip4Location": js["momentData"]["ip4Location"],
            "id": js["momentData"]["id"],
            "uuid": js["momentData"]["uuid"],
            "content": js["momentData"]["content"].replace("\n", "\\n"),
            "urlid": js["momentData"]["uuid"]
        }
    elif "contentData" in js:
        result = {
            "contentId": js.get("contentId", ""),
            "contentType": js.get("contentType", ""),
            "userId": js["userBrief"]["userId"],
            "nickname": js["userBrief"]["nickname"],
            "workTime": js["userBrief"]["workTime"],
            "educationInfo": js["userBrief"]["educationInfo"],
            "secondMajorName": js["userBrief"]["secondMajorName"],
            "title": js["contentData"]["title"],
            "ip4Location": '',
            "id": js["contentData"]["id"],
            "uuid": js["contentData"]["uuid"],
            "content": js["contentData"]["content"].replace("\n", "\\n"),
            "urlid": js["contentData"]["id"]
        }
    else:
        result = {
            "contentId": js.get("contentId", ""),
            "contentType": js.get("contentType", ""),
            "userId": js["userBrief"]["userId"],
            "nickname": js["userBrief"]["nickname"],
            "workTime": js["userBrief"]["workTime"],
            "educationInfo": js["userBrief"]["educationInfo"],
            "secondMajorName": js["userBrief"]["secondMajorName"],
            "title": '',
            "ip4Location": '',
            "id": '',
            "uuid": '',
            "content": '',
            "urlid": ''
        }
    return result

def fetch_answer(df):
    all_jobs = {}
    count = 1
    for index,row in df.iterrows() :
        params = {
            'entityId': row['contentId'],
            'entityType': row['contentType'],
            '_': '1756622181048',
        }

        resp = requests.get(
            'https://gw-c.nowcoder.com/api/sparta/ai-experience/pc/queryExperienceQuestionList',
            params=params,
            cookies=cookies,
            headers=headers,
        )

        resp.raise_for_status()
        jobs = resp.json()
        data = (jobs or {}).get('data') or {}
        data = data.get('experienceQuestionList') or []
        clean_data = []
        for item in data:
            clean_data.append({'title': item.get('title', ''),'answer': item.get('answer', '')})
        all_jobs[row['contentId']] = clean_data
        time.sleep(1)
        print(f"已获取 {row['contentId']} 共 {len(clean_data)} 条")
        count += 1
        # if count > 5:
        #     break
    return all_jobs

def dict_of_listdict_to_df(src: dict) -> pd.DataFrame:
    # 1. 把键-列表对拆成 (key, row) 的生成器
    records = ((k, row) for k, lst in src.items() for row in lst)
    # 2. 构造 DataFrame
    df = pd.DataFrame.from_records(records, columns=['key', 'row'])
    # 3. 把 row 列里的字典拆成多列，并合并
    return df.join(pd.json_normalize(df.pop('row')))

# 大厂代码
bigcompany = [134, 138, 139, 147, 149, 151, 179, 239, 652, 665, 732, 931]
# 关键词代码,11204:大数据, 11240:大模型
jobId = [11204,11240]

d = {
     '数据大厂面经':{'companyList':bigcompany,'jobid':11204},
     '数据面经':{'companyList':[],'jobid':11204},
    '大模型大厂面经':{'companyList':bigcompany,'jobid':11240},
      '大模型面经':{'companyList':[],'jobid':11240}
}


result2 = {}
result3 = {}
for k, v in d.items():
    data = fetch_all_ant(k,v)
    result = {}
    for d in data.values():
        for i in range(len(d['data']['records'])):
            row = extract_row(d['data']['records'][i])
            result[row['contentId']] = row
        df = pd.DataFrame(result.values())
    r = {row['contentId']:row['contentType'] for index, row in df.iterrows() if 'contentId' in row and 'contentType' in row}
    result2.update(r) 
    result3.update(result)

df1 = pd.DataFrame(result3.values())
df1.to_excel('面经\面经.xlsx')
df2 = pd.DataFrame(result2.items(), columns=['contentId', 'contentType'])
data = fetch_answer(df2)
df3 = dict_of_listdict_to_df(data)
df3.to_excel('面经\牛客网数据面经&答案.xlsx', index=False)
print(f"总共获取到 {len(df3)} 条")