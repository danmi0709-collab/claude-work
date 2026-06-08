# -*- coding: utf-8 -*-
import pandas as pd
import re

df = pd.read_excel('C:/Users/한나/Downloads/03. 표제부_2026-05-05 08_56_38.xlsx')
apt_all = df[df['주용도코드명'] == '공동주택']

# 미매칭 단지 → 원본 데이터에서 유사한 이름 검색
missing = [
    '잠실엘스', '잠실주공5단지', '파크리오', '잠실르엘', '잠실올림픽공원아이파크',
    '풍납현대', '아시아선수촌', '우성4차', '트리지움', '레이크팰리스',
    '올림픽선수기자촌', '대림가락', '헬리오시티', '가락삼익맨숀',
    '가락대림', '미륭', '힐스테이트e편한세상문정', '문정시영',
    '위례아이파크', '가락우창', '삼환가락', '가락극동', '가락한신',
    '올림픽훼밀리타운', '미성맨션', '코오롱', '가락상아', '신성노바빌',
    '현대리버빌', '상아2차', '위례중앙푸르지오', '잠실더샵루벤', '한신잠실코아'
]

with open('송파구/find_missing_result.txt', 'w', encoding='utf-8') as f:
    for keyword in missing:
        kw_clean = re.sub(r'\(.*?\)', '', keyword).strip()
        # 건물명에 keyword 포함된 것 찾기
        found = apt_all[apt_all['건물명'].str.contains(kw_clean, na=False, case=False)]
        if len(found) == 0:
            # 일부 단어로 재검색
            words = kw_clean[:4]
            found = apt_all[apt_all['건물명'].str.contains(words, na=False, case=False)]

        if len(found) > 0:
            names = found['건물명'].unique()[:5]
            vals = []
            for nm in names:
                v = found[found['건물명']==nm]['용적률(%)'].mean()
                vals.append(f"{nm}({v:.0f}%)")
            f.write(f"[{keyword}] → {', '.join(vals)}\n")
        else:
            f.write(f"[{keyword}] → 없음\n")
