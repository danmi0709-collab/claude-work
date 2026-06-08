import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

with open('송파구/apt_data.json', encoding='utf-8') as f:
    data = json.load(f)

# 송파구 이미지의 지리 경계 (실측 기반 추정)
# 이미지: 4845 x 5111 픽셀
IMG_W = 4845
IMG_H = 5111

# 지도 범위 (위도/경도)
LAT_MAX = 37.5470   # 이미지 위쪽 (북)
LAT_MIN = 37.4570   # 이미지 아래쪽 (남)
LNG_MIN = 127.0750  # 이미지 왼쪽 (서)
LNG_MAX = 127.1970  # 이미지 오른쪽 (동)

def to_pixel(lat, lng):
    x = (lng - LNG_MIN) / (LNG_MAX - LNG_MIN) * IMG_W
    y = (LAT_MAX - lat) / (LAT_MAX - LAT_MIN) * IMG_H
    return x, y

# 마커 JS 데이터
markers_js = []
for apt in data:
    x, y = to_pixel(apt['lat'], apt['lng'])
    label = f"{apt['name']}<br>{apt['yr2']}' {apt['units']:,}^"
    popup = f"{apt['name']}<br>{apt['yr_full']} · {apt['units']:,}세대"
    markers_js.append(
        f"{{x:{x:.1f},y:{y:.1f},label:{json.dumps(label, ensure_ascii=False)},"
        f"popup:{json.dumps(popup, ensure_ascii=False)},"
        f"color:{json.dumps(apt['color'])},bg:{json.dumps(apt['bg'])}}}"
    )

markers_str = ',\n'.join(markers_js)

legend = '''
<div id="legend">
  <b>준공 연식</b><br>
  <span style="color:#c0392b">■</span> ~1989 구축<br>
  <span style="color:#d35400">■</span> 1990~1999<br>
  <span style="color:#1e8449">■</span> 2000~2009<br>
  <span style="color:#1a5276">■</span> 2010~2019<br>
  <span style="color:#6c3483">■</span> 2020년~
</div>
'''

html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>송파구 아파트 임장지도</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#222; overflow:hidden; }}
#wrap {{ position:relative; display:inline-block; cursor:grab; }}
#wrap:active {{ cursor:grabbing; }}
#map-img {{ display:block; }}
.apt-label {{
  position:absolute;
  font-family:'Malgun Gothic','나눔고딕',sans-serif;
  font-size:13px;
  font-weight:bold;
  line-height:1.5;
  padding:3px 7px;
  border-radius:5px;
  border-width:2px;
  border-style:solid;
  white-space:nowrap;
  box-shadow:1px 2px 5px rgba(0,0,0,0.3);
  cursor:pointer;
  transform:translate(-50%,-50%);
  pointer-events:auto;
}}
#tooltip {{
  position:fixed;
  background:rgba(0,0,0,0.82);
  color:#fff;
  padding:7px 12px;
  border-radius:7px;
  font-family:'Malgun Gothic',sans-serif;
  font-size:13px;
  line-height:1.6;
  pointer-events:none;
  display:none;
  z-index:999;
}}
#legend {{
  position:fixed;
  bottom:20px;
  left:20px;
  background:rgba(255,255,255,0.93);
  padding:10px 14px;
  border-radius:8px;
  border:2px solid #aaa;
  font-family:'Malgun Gothic',sans-serif;
  font-size:13px;
  line-height:1.8;
  box-shadow:2px 2px 6px rgba(0,0,0,0.2);
  z-index:998;
}}
#controls {{
  position:fixed;
  top:16px;
  right:16px;
  z-index:998;
  display:flex;
  gap:8px;
}}
.ctrl-btn {{
  background:white;
  border:2px solid #888;
  border-radius:6px;
  font-size:20px;
  width:38px; height:38px;
  cursor:pointer;
  font-weight:bold;
  box-shadow:1px 1px 4px rgba(0,0,0,0.2);
}}
</style>
</head>
<body>
<div id="controls">
  <button class="ctrl-btn" onclick="zoom(1.2)">+</button>
  <button class="ctrl-btn" onclick="zoom(0.83)">−</button>
  <button class="ctrl-btn" onclick="resetView()" title="초기화">⌂</button>
</div>
{legend}
<div id="tooltip"></div>
<div id="wrap">
  <img id="map-img" src="송파구이미지2.png" draggable="false">
</div>

<script>
const RAW_W = {IMG_W}, RAW_H = {IMG_H};
const markers = [
{markers_str}
];

const wrap = document.getElementById('wrap');
const img = document.getElementById('map-img');
const tooltip = document.getElementById('tooltip');

let scale = 1, offX = 0, offY = 0;
let dragging = false, startX, startY, startOX, startOY;

function applyTransform() {{
  wrap.style.transform = `translate(${{offX}}px,${{offY}}px) scale(${{scale}})`;
  wrap.style.transformOrigin = '0 0';
}}

function initScale() {{
  const ws = window.innerWidth / RAW_W;
  const hs = window.innerHeight / RAW_H;
  scale = Math.min(ws, hs);
  offX = (window.innerWidth - RAW_W * scale) / 2;
  offY = (window.innerHeight - RAW_H * scale) / 2;
  applyTransform();
}}

function zoom(factor) {{
  const cx = window.innerWidth / 2, cy = window.innerHeight / 2;
  const nx = (cx - offX) / scale, ny = (cy - offY) / scale;
  scale *= factor;
  offX = cx - nx * scale;
  offY = cy - ny * scale;
  applyTransform();
}}

function resetView() {{ initScale(); }}

// drag
wrap.addEventListener('mousedown', e => {{
  dragging = true; startX = e.clientX; startY = e.clientY;
  startOX = offX; startOY = offY;
}});
window.addEventListener('mousemove', e => {{
  if (!dragging) return;
  offX = startOX + e.clientX - startX;
  offY = startOY + e.clientY - startY;
  applyTransform();
}});
window.addEventListener('mouseup', () => dragging = false);

// 휠 줌
window.addEventListener('wheel', e => {{
  e.preventDefault();
  const f = e.deltaY < 0 ? 1.12 : 0.89;
  const nx = (e.clientX - offX) / scale, ny = (e.clientY - offY) / scale;
  scale *= f;
  offX = e.clientX - nx * scale;
  offY = e.clientY - ny * scale;
  applyTransform();
}}, {{passive:false}});

// 마커 생성
markers.forEach(m => {{
  const el = document.createElement('div');
  el.className = 'apt-label';
  el.innerHTML = m.label;
  el.style.left = m.x + 'px';
  el.style.top = m.y + 'px';
  el.style.color = m.color;
  el.style.background = m.bg;
  el.style.borderColor = m.color;

  el.addEventListener('mouseenter', e => {{
    tooltip.innerHTML = m.popup;
    tooltip.style.display = 'block';
  }});
  el.addEventListener('mousemove', e => {{
    tooltip.style.left = (e.clientX + 14) + 'px';
    tooltip.style.top = (e.clientY - 10) + 'px';
  }});
  el.addEventListener('mouseleave', () => tooltip.style.display = 'none');
  wrap.appendChild(el);
}});

initScale();
</script>
</body>
</html>"""

with open('송파구/송파구_아파트_임장지도.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("완료!")
