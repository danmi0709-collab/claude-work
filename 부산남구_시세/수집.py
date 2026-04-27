"""부산 남구 매매 실거래가: 2021-01 ~ 2023-12 (36개월)"""
import sys, io, os, json, time, csv, urllib.parse, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, 'config.json'), 'r', encoding='utf-8') as f:
    cfg = json.load(f)

URL_TRADE = 'https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev'


def fetch_all_month(base_url, ymd):
    out = []; page = 1
    while True:
        params = {
            'serviceKey': cfg['service_key'], 'LAWD_CD': cfg['lawd_cd'], 'DEAL_YMD': ymd,
            'numOfRows': '1000', 'pageNo': str(page), '_type': 'json',
        }
        url = base_url + '?' + urllib.parse.urlencode(params, safe=':/')
        with urllib.request.urlopen(url, timeout=60) as r:
            data = json.loads(r.read().decode('utf-8'))
        body = data.get('response', {}).get('body', {})
        items = body.get('items', {})
        if not items: break
        lst = items.get('item', [])
        if isinstance(lst, dict): lst = [lst]
        out.extend(lst)
        total = int(body.get('totalCount', 0))
        if page * 1000 >= total: break
        page += 1; time.sleep(0.3)
    return out


def months_between(s, e):
    sy, sm = int(s[:4]), int(s[4:])
    ey, em = int(e[:4]), int(e[4:])
    out = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        out.append(f'{y:04d}{m:02d}')
        m += 1
        if m > 12: m = 1; y += 1
    return out


ym_list = months_between('202101', '202312')
all_items = []
for i, ym in enumerate(ym_list, 1):
    try:
        items = fetch_all_month(URL_TRADE, ym)
        print(f'  {ym}: {len(items)}건 ({i}/{len(ym_list)})')
        all_items.extend(items)
        time.sleep(0.3)
    except Exception as e:
        print(f'  {ym}: 에러 - {e}')

FIELDS = ['dealYear','dealMonth','dealDay','umdNm','aptNm','buildYear',
          'excluUseAr','dealAmount','floor','jibun','aptSeq']
out_path = os.path.join(BASE, '매매_원본.csv')
with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=FIELDS, extrasaction='ignore')
    w.writeheader()
    for it in all_items: w.writerow(it)
print(f'저장: {out_path} ({len(all_items)}건)')
