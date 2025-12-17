# pinduoduo_jobs_spider.py
# 拼多多招聘数据爬虫 – 绕过 anti_content
# 依赖: pip install DrissionPage
# 运行: python pinduoduo_jobs_spider.py

import json, os, time, glob
from DrissionPage import ChromiumPage, ChromiumOptions
import pandas as pd
from tqdm import tqdm

company = "拼多多"
SAVE_EXCEL = f"招聘JD\{company}.xlsx"
SAVE_JSON = f"招聘JD\pdd_detail_json"
os.makedirs(SAVE_JSON)
def fetch_jobs_code():
    result = []
    co = ChromiumOptions()
    co.headless(False)          # 无头模式可改为 True
    page = ChromiumPage(co)
    api_path = '/api/recruit/position/list'
    page.listen.start(api_path)

    page.get('https://careers.pinduoduo.com/jobs#/')
    res = page.listen.wait()
    if res:
        result.append(res.response.body)

    for p in range(2, 79):      # 按需调整页数
        page.listen.start(api_path)
        btn = page.ele(f'x://a[text()="{p}"]')
        if btn:
            btn.click()
            resp = page.listen.wait()
            if resp:
                result.append(res.response.body)
        time.sleep(2)
    page.quit()
    return result

def fetch_detail(code: str, page: ChromiumPage):
    result = []
    api = "/api/recruit/position/detail"
    page.listen.start(api)
    page.get(f"https://careers.pddglobalhr.com/jobs/detail?code={code}")
    resp = page.listen.wait()
    if resp and resp.response.body:
        data = resp.response.body
        # 落盘 JSON
        json_path = os.path.join(SAVE_JSON, f"{code}.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return data
    return None

def get_exist_code():
    exist_code = []
    for f in glob.glob(SAVE_JSON):
        data = json.load(open(f, encoding="utf-8"))
        if data.get("errorCode", []) == 1000000:
            exist_code.append(data.get('result',{}).get('code'))
            os.remove(f)
    return exist_code

def save_to_excel():
    row = []
    for f in glob.glob(SAVE_JSON):
        data = json.load(open(f, encoding="utf-8"))
        if data.get("errorCode", []) == 1000000:
            item = data.get('result',{})
            row.append(
                {
                    "分类": '大模型',
                    "职位id": item.get('code'),
                    "职位名称": item["name"],
                    "工作地点": item.get('workLocation'),
                    "职位描述": item["jobDuty"],
                    "职位要求": item["serveRequirement"],
                    "职位类别ID": '',
                    "职位类别名称": item["job"],
                    "发布时间": item["updateTime"]
                }
            )
        else:
            os.remove(f)
    d = pd.DataFrame(row)
    d.to_excel(f'招聘JD\{company}.xlsx',index=False)


def main():
    code_result = fetch_jobs_code()
    codes = []
    for data in code_result:
        if data.get("result", []).get('list',[]):
            for item in data.get("result", []).get('list',[]):
                if ('数据' in item.get('name') or '大模型' in item.get('name')): #and item.get('updateTime') >= '2025-06-01':
                    codes.append(item.get('code'))
    print("共提取", len(codes), "个职位 code")

    co = ChromiumOptions()
    co.headless(True)               # 调试用 False
    page = ChromiumPage(co)
    all_data = []

    for code in tqdm(codes, desc="detail"):
        exist_code = get_exist_code()
        if code not in exist_code:
            try:
                d = fetch_detail(code, page)
                if d:
                    all_data.append(d)
            except Exception as e:
                print(f"[warn] {code} 异常：{e}")
            time.sleep(1.5)
        else:
            continue
    page.quit()
    save_to_excel()

main()