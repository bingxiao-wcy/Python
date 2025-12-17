import requests, pandas as pd, time, datetime
import os, json
cookies = {
    'a1': '198e5fdb0afv1475801jh3uc8qa7qazacd8chpzzs00000874538',
    'webId': 'a3f02890030a940dee6bb2bf0c6b65df',
    'gid': 'yjYd2ifD0qSqyjYd2ifD8fUu0ihy4W2Y8yVxYuMq7SYA0S88WA0u0x888YW42qY8Kj8DjJdD',
    'abRequestId': 'a3f02890030a940dee6bb2bf0c6b65df',
    'webBuild': '4.83.1',
    'acw_tc': '0a4a2ce817615651217654668e6eab1d38c10e09e99ba70dafaa743a4d14d4',
    'xsecappid': 'xhs-pc-web',
    'loadts': '1761565377409',
    'web_session': '040069b5a8046d5697279112ca3a4b372a8147',
    'websectiga': '3fff3a6f9f07284b62c0f2ebf91a3b10193175c06e4f71492b60e056edcdebb2',
    'sec_poison_id': 'ee978c90-3940-4883-94a9-93ba094e5950',
    'unread': '{%22ub%22:%2268fec0140000000005010905%22%2C%22ue%22:%2268dbbdd1000000000703279f%22%2C%22uc%22:25}',
}

headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9',
    'cache-control': 'no-cache',
    'content-type': 'application/json;charset=UTF-8',
    'origin': 'https://www.xiaohongshu.com',
    'pragma': 'no-cache',
    'priority': 'u=1, i',
    'referer': 'https://www.xiaohongshu.com/',
    'sec-ch-ua': '"Chromium";v="140", "Not=A?Brand";v="24", "Google Chrome";v="140"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-site',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Mobile Safari/537.36',
    'x-b3-traceid': '2b5da068061b0881',
    'x-s': 'XYS_2UQhPsHCH0c1Pjh9HjIj2erjwjQhyoPTqBPT49pjHjIj2eHjwjQgynEDJ74AHjIj2ePjwjQTJdPIP/ZlgMrU4SmH4B8xnBS/zSze/BlQwnzePe+9+7SByfRraeQBanYInDlT/o4yJrkx2fE0JfMTzF+tJBTI80Q0GAYaLdDA/A8NNMzyyFSj+pYHJrE3GpYOzbSh4dGApASCLgmFw/bTpFM3LeLM/LkO8F+TpMHUwBkxLf+h4nMsLdzkwnSaznEi/AQoP/zQaBc6JrQ8y9YiPb494UR8/Lli4FkD/9lP4Dh7+BD3PdmMHjIj2ecjwjQ6GfkSG7cjKc==',
    'x-s-common': '2UQAPsHC+aIjqArjwjHjNsQhPsHCH0rjNsQhPaHCH0c1Pjh9HjIj2eHjwjQgynEDJ74AHjIj2ePjwjQhyoPTqBPT49pjHjIj2ecjwjHFN0WAN0rjNsQh+aHCH0rEwBLM8fzjPBbf40rF+ALhPebxye+MGAYlG/4lGgkYG9chG9YI2dkAPeZIPeZh+AcMPAWjNsQh+jHCHjHVHdW7H0ijHjIj2eWjwjQQPAYUaBzdq9k6qB4Q4fpA8b878FSet9RQzLlTcSiM8/+n4MYP8F8LagY/P9Ql4FpUzfpS2BcI8nT1GFbC/L88JdbFyrSiafp/8DMra7pFLDDAa7+8J7QgabmFz7Qjp0mcwp4fanD68p40+fp8qgzELLbILrDA+9p3JpH9LLI3+LSk+d+DJfpSL98lnLYl49IUqgcMc0mrcDShtUTozbG6qM8FyFSh8o+h4g4U+obFyLSi4nbQz/+SPFlnPrDApSzQcA4SPopFJeQmzBMA/o8Szb+NqM+c4ApQzg8Ayp8FaDRl4AYs4g4fLomD8pzBpFRQ2ezLanSM+Skc49Qc4gzGag8VGLlj87PAqgzhagYSqAbn4FYQy7pTanTQ2npx87+8NM4L89L78p+l4BL6ze4AzB+IygmS8Bp8qDzFaLP98Lzn4AQQzLEAL7bFJBEVL7pwyS8Fag868nTl4e+0n04ApfuF8FSbL7SQyrLhtASrpLS92dDFa/YOanS0+Mkc4FbQ4fSa+Bu6qFzP8oP9Lo4naLP78p+D+7+DPbHFaLp9qA+QzFMFpd4panSDqA+AN7+hnDESyp8FGf+p8np8pd4iag8bqoi6cnpf4g4aqeSmq98c4FQQ2BlFagYyL9RM4FRdpd4Iq7HFyBppN9L9/o8Szbm7zDS987PlqfRAPLzyyLSk+7+xGfRAP94UzDSbPBLALoz9anSjLDRl4FROqgziagYSq7Yc4A4QyrbSpSmFyrSiN7+8qgz/z7b72nMc4FzQ4DS3a/+Q4ezYzMPFnaRSygpFyDSkJgQQzLRALM8F2DQ6zDF6wg8Sy0Sy4DSkzLEo4gzCqdpFJrS94fLALozp/7mN8p8gcgPAqBY7anY6qAPE/7PA4gzAGMm7GLSead+gLoqManSd8nTSqLlQcFTSyfc6q98c4epQ2e4A2op7zezTqo4QyM4Eag8SqA8BP7PlaLRSPb46qM+M4bkQy9I9agYkaaHVHdWEH0ilP/chPAqFPAWANsQhP/Zjw0ZVHdWlPaHCHfE6qfMYJsQR',
    'x-t': '1761565462222',
    'x-xray-traceid': 'cd12bde759a67424008a2292394ceb51',
    # 'cookie': 'a1=198e5fdb0afv1475801jh3uc8qa7qazacd8chpzzs00000874538; webId=a3f02890030a940dee6bb2bf0c6b65df; gid=yjYd2ifD0qSqyjYd2ifD8fUu0ihy4W2Y8yVxYuMq7SYA0S88WA0u0x888YW42qY8Kj8DjJdD; abRequestId=a3f02890030a940dee6bb2bf0c6b65df; webBuild=4.83.1; acw_tc=0a4a2ce817615651217654668e6eab1d38c10e09e99ba70dafaa743a4d14d4; xsecappid=xhs-pc-web; loadts=1761565377409; web_session=040069b5a8046d5697279112ca3a4b372a8147; websectiga=3fff3a6f9f07284b62c0f2ebf91a3b10193175c06e4f71492b60e056edcdebb2; sec_poison_id=ee978c90-3940-4883-94a9-93ba094e5950; unread={%22ub%22:%2268fec0140000000005010905%22%2C%22ue%22:%2268dbbdd1000000000703279f%22%2C%22uc%22:25}',
}

# --------------------------------------------------
# 2. 和你一样用 POST + json_data
# --------------------------------------------------
def fetch_all_ant(keyword: str) -> list[dict]:
    all_jobs = []
    page_index = 1
    page_size  = 20          # 你原始也是 10

    while True:
        json_data = {
            'keyword': keyword,
            'page': page_index,
            'page_size': page_size,
            'search_id': '2fidyi9tlyjchtc6s4o1i',
            'sort': 'general',
            'note_type': 0,
            'geo': '',
            'image_formats': [
                'jpg',
                'webp',
                'avif',
            ],
        }
        response = requests.post(
            'https://edith.xiaohongshu.com/api/sns/web/v1/search/notes',
            cookies=cookies,
            headers=headers,
            json=json_data,
        )
        result = response.json()
        for item in result.get('data', {}).get('items', []):
            if item:
                note_id = item.get('id')
                xsec_token = item.get('xsec_token')
                url = (
                    f"https://edith.xiaohongshu.com/api/sns/web/v2/comment/page?"
                    f"note_id={note_id}&cursor=&top_comment_id="
                    f"&image_formats=jpg,webp,avif&xsec_token={xsec_token}"
                )
                print(f"正在抓取笔记详情：{note_id}")
                response_detail = requests.get(
                    url,
                    cookies=cookies,
                    headers=headers,
                )
                json_dir = r'spider\note\detailinfo'          # 原始目录
                json_path = os.path.join(json_dir, f"{item.get('id')}.json")

                # 1. 目录不存在就递归创建
                os.makedirs(json_dir, exist_ok=True)

                # 2. 写文件
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(response_detail.json(), f, ensure_ascii=False, indent=2)
                time.sleep(1)
            else:
                print("空item")
        print(result)
        page_index += 1
        time.sleep(1)
fetch_all_ant("大模型askmeanything")
