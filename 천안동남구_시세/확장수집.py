"""2018-01 ~ 2019-12 (24개월) + 2026-01 ~ 2026-04 (4개월) 추가 수집"""
import sys, io, os, json, time, csv, urllib.parse, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, 'config.json'), 'r', encoding='utf-8') as f:
    cfg = json.load(f)

URL_T = 'https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev'
URL_R = 'https://apis.data.go.kr/1613000/RTMSDataSvcAptRent/getRTMSDataSvcAptRent'


def fetch_all(url, ymd):
    out = []; page = 1
    while True:
        p = {'serviceKey': cfg['service_key'], 'LAWD_CD': cfg['lawd_cd'], 'DEAL_YMD': ymd,
             'numOfRows': '1000', 'pageNo': str(page), '_type': 'json'}
        u = url + '?' + urllib.parse.urlencode(p, safe=':/')
        with urllib.request.urlopen(u, timeout=60) as r:
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


def collect(url, ym_list, label):
    items = []
    for i, ym in enumerate(ym_list, 1):
        try:
            it = fetch_all(url, ym)
            print(f'  [{label}] {ym}: {len(it)}건 ({i}/{len(ym_list)})')
            items.extend(it); time.sleep(0.3)
        except Exception as e:
            print(f'  [{label}] {ym}: 에러 {e}')
    return items


def append_csv(path, items, fields):
    existing = []
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8-sig') as f:
            existing = list(csv.DictReader(f))
    merged = existing + items
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction='ignore')
        w.writeheader()
        for it in merged: w.writerow(it)
    print(f'  저장 완료: {path} (기존 {len(existing)} + 추가 {len(items)} = {len(merged)}건)')


TF = ['dealYear','dealMonth','dealDay','umdNm','aptNm','buildYear',
      'excluUseAr','dealAmount','floor','jibun','aptSeq']
RF = ['dealYear','dealMonth','dealDay','umdNm','aptNm','buildYear',
      'excluUseAr','deposit','monthlyRent','contractType','floor','jibun','aptSeq']

ym_list = months_between('201801', '201912') + months_between('202601', '202604')
print('=== 매매 ===')
trades = collect(URL_T, ym_list, '매매')
append_csv(os.path.join(BASE, '매매_원본.csv'), trades, TF)

print('\n=== 전세 ===')
rents = collect(URL_R, ym_list, '전세')
append_csv(os.path.join(BASE, '전세_원본.csv'), rents, RF)
