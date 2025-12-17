import requests, pandas as pd, time, datetime

# --------------------------------------------------
# 1. 和浏览器完全一致的头/ cookie / ctoken
# --------------------------------------------------
import requests

import requests

cookies = {
    'locale': 'zh-CN',
    's_v_web_id': 'verify_metwq6t1_OBrFqLRu_K8iR_4UpW_AS5u_9wz54tKHY80F',
    'passport_web_did': '7547667303912407068',
    'passport_trace_id': '7547667303933100060',
    'QXV0aHpDb250ZXh0': '26e82aca540947be89a91a85ccaa18b1',
    '_gcl_au': '1.1.1093960967.1757328238',
    'landing_url': 'https://www.feishu.cn/accounts/page/ug_register?template_id=7168397144091885571&tracking_code=701TL0000097HDTYA2&utm_from=Bingsem_keyword_pc_office_pinpai_pinpai_pinpai2&device=PC&e_keywordid=1&bd_vid=7640742032784342938&msclkid=ba3359f6c8fd12fc36c931487b97a573',
    '__tea__ug__uid': '7547667296471746099',
    'fid': '63e123a8-d632-4417-8349-3e2068c8c69a',
    's_v_web_id': 'verify_mfazvbu9_67N39Hqn_LOF2_4sQf_AmDy_6weRxS1KfkyY',
    '_uuid_hera_ab_path_1': '7547667525396807684',
    'i18n_locale': 'zh-CN',
    '_uetvid': 'd80c49d08ca011f08f5d459c2b23ea17',
    'Hm_lvt_a79616d9322d81f12a92402ac6ae32ea': '1757328523',
    'swp_csrf_token': '09855cb6-b138-4c5f-8e2f-c5c8db8fadbb',
    't_beda37': 'ff56310c83ac26c4cbeacd17e6bb3e69bf5d233b9241c7b9ba614b987d3592fa',
    '_uuid_hera_ab_hire_path_1': '7552917212156411906',
    'first_landing_url': 'https%3A%2F%2Fhire.feishu.cn%2F',
    '_gid': 'GA1.2.1060513354.1758550577',
    '_ga': 'GA1.2.1895883327.1757328238',
    '_ga_VPYRHN104D': 'GS2.1.s1758550577$o2$g1$t1758550615$j22$l0$h0',
    'device-id': '7543229759452055090',
    'channel': 'saas-career',
    'platform': 'pc',
    'atsx-csrf-token': 'Rci7wM2KBJXlmePF1K3Bwz9lrKEppk1JJefOeIdxARg%3D',
}

headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN',
    'cache-control': 'no-cache',
    'content-type': 'application/json',
    'env': 'undefined',
    'origin': 'https://zhipu-ai.jobs.feishu.cn',
    'portal-channel': 'saas-career',
    'portal-platform': 'pc',
    'pragma': 'no-cache',
    'priority': 'u=1, i',
    'referer': 'https://zhipu-ai.jobs.feishu.cn/index/?keywords=&category=&location=&project=&type=&job_hot_flag=&current=2&limit=10&functionCategory=&tag=',
    'sec-ch-ua': '"Chromium";v="140", "Not=A?Brand";v="24", "Google Chrome";v="140"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Mobile Safari/537.36',
    'website-path': 'index',
    'x-csrf-token': 'Rci7wM2KBJXlmePF1K3Bwz9lrKEppk1JJefOeIdxARg=',
    # 'cookie': 'locale=zh-CN; s_v_web_id=verify_metwq6t1_OBrFqLRu_K8iR_4UpW_AS5u_9wz54tKHY80F; passport_web_did=7547667303912407068; passport_trace_id=7547667303933100060; QXV0aHpDb250ZXh0=26e82aca540947be89a91a85ccaa18b1; _gcl_au=1.1.1093960967.1757328238; landing_url=https://www.feishu.cn/accounts/page/ug_register?template_id=7168397144091885571&tracking_code=701TL0000097HDTYA2&utm_from=Bingsem_keyword_pc_office_pinpai_pinpai_pinpai2&device=PC&e_keywordid=1&bd_vid=7640742032784342938&msclkid=ba3359f6c8fd12fc36c931487b97a573; __tea__ug__uid=7547667296471746099; fid=63e123a8-d632-4417-8349-3e2068c8c69a; s_v_web_id=verify_mfazvbu9_67N39Hqn_LOF2_4sQf_AmDy_6weRxS1KfkyY; _uuid_hera_ab_path_1=7547667525396807684; i18n_locale=zh-CN; _uetvid=d80c49d08ca011f08f5d459c2b23ea17; Hm_lvt_a79616d9322d81f12a92402ac6ae32ea=1757328523; swp_csrf_token=09855cb6-b138-4c5f-8e2f-c5c8db8fadbb; t_beda37=ff56310c83ac26c4cbeacd17e6bb3e69bf5d233b9241c7b9ba614b987d3592fa; _uuid_hera_ab_hire_path_1=7552917212156411906; first_landing_url=https%3A%2F%2Fhire.feishu.cn%2F; _gid=GA1.2.1060513354.1758550577; _ga=GA1.2.1895883327.1757328238; _ga_VPYRHN104D=GS2.1.s1758550577$o2$g1$t1758550615$j22$l0$h0; device-id=7543229759452055090; channel=saas-career; platform=pc; atsx-csrf-token=Rci7wM2KBJXlmePF1K3Bwz9lrKEppk1JJefOeIdxARg%3D',
}

params = {
    'keyword': '',
    'limit': '10',
    'offset': '0',
    'job_category_id_list': '',
    'tag_id_list': '',
    'location_code_list': '',
    'subject_id_list': '',
    'recruitment_id_list': '',
    'portal_type': '6',
    'job_function_id_list': '',
    'storefront_id_list': '',
    'portal_entrance': '1',
    '_signature': 'eoobbwAAAACZ4DUNsgM69XqKG3AABIm',
}




company = "智谱AI"

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index  = 0          # 从第 0 页开始
    limit = 30         # 每页条数

    while True:
        json_data = {
            'keyword': keyword,
            'limit': limit,
            'offset': page_index * limit,   # 关键点：动态偏移
            'job_category_id_list': [],
            'tag_id_list': [],
            'location_code_list': [],
            'subject_id_list': [],
            'recruitment_id_list': [],
            'portal_type': 6,
            'job_function_id_list': [],
            'storefront_id_list': [],
            'portal_entrance': 1,
        }

        resp = requests.post(
            'https://zhipu-ai.jobs.feishu.cn/api/v1/search/job/posts',
            params=params,
            cookies=cookies,
            headers=headers,
            json=json_data,
        )
        resp.raise_for_status()
        data = resp.json()
        jobs = data.get("data", {}).get("job_post_list", [])
        
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
        "职位id": j["job_category"]["id"],
        "职位名称": j["title"],
        "工作地点": '',
        "职位描述": j.get("description", ""),
        "职位要求": j.get("requirement", ""),
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
