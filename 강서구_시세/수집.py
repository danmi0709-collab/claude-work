"""서울 강서구 매매 실거래가 수집 (2018-01 ~ 2026-04)"""
import sys, io, os, json, time, csv, urllib.parse, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, 'config.json'), 'r', encoding='utf-8') as f:
    cfg = json.load(f)

URL_T = 'https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev'

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

TF = ['dealYear','dealMonth','dealDay','umdNm','aptNm','buildYear',
      'excluUseAr','dealAmount','floor','jibun','aptSeq']

ym_list = months_between('201801', '202604')
all_items = []
for i, ym in enumerate(ym_list, 1):
    try:
        it = fetch_all(URL_T, ym)
        print(f'[매매] {ym}: {len(it)}건 ({i}/{len(ym_list)})')
        all_items.extend(it); time.sleep(0.3)
    except Exception as e:
        print(f'[매매] {ym}: 에러 {e}')

out_path = os.path.join(BASE, '매매_원본.csv')
with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=TF, extrasaction='ignore')
    w.writeheader()
    for it in all_items: w.writerow(it)
print(f'\n저장: {out_path} ({len(all_items)}건)')
