"""천안 동남구 전세 실거래가 추가 수집 (2020-01 ~ 2025-04, 기존 12개월 앞 구간)"""
import sys, io, os, json, time, csv, urllib.parse, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, 'config.json'), 'r', encoding='utf-8') as f:
    cfg = json.load(f)

URL = 'https://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent'

def fetch_all(ymd):
    out = []; page = 1
    while True:
        p = {'serviceKey': cfg['service_key'], 'LAWD_CD': cfg['lawd_cd'], 'DEAL_YMD': ymd,
             'numOfRows': '1000', 'pageNo': str(page), '_type': 'json'}
        url = URL + '?' + urllib.parse.urlencode(p, safe=':/')
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
    sy, sm = int(s[:4]), int(s[4:]); ey, em = int(e[:4]), int(e[4:])
    out = []; y, m = sy, sm
    while (y, m) <= (ey, em):
        out.append(f'{y:04d}{m:02d}'); m += 1
        if m > 12: m = 1; y += 1
    return out

ym_list = months_between('202001', '202504')
all_items = []
for i, ym in enumerate(ym_list, 1):
    try:
        it = fetch_all(ym)
        print(f'  {ym}: {len(it)}건 ({i}/{len(ym_list)})')
        all_items.extend(it); time.sleep(0.3)
    except Exception as e:
        print(f'  {ym}: 에러 {e}')

# 기존 전세_원본.csv에 prepend
existing = []
csv_path = os.path.join(BASE, '전세_원본.csv')
if os.path.exists(csv_path):
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        existing = list(csv.DictReader(f))

FIELDS = ['dealYear','dealMonth','dealDay','umdNm','aptNm','buildYear',
          'excluUseAr','deposit','monthlyRent','contractType','floor','jibun','aptSeq']

merged = all_items + existing
with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=FIELDS, extrasaction='ignore')
    w.writeheader()
    for it in merged: w.writerow(it)
print(f'저장: {csv_path} (추가 {len(all_items)} + 기존 {len(existing)} = {len(merged)}건)')
