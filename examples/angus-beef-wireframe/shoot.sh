#!/bin/bash
# 截每屏 375x667 @2.4x → 900x1600；以及整页长图 @2x
cd "$(dirname "$0")"
mkdir -p screens
for f in _shots/s*.html; do
  id=$(basename "$f" .html)
  google-chrome --headless=new --no-sandbox --disable-gpu --hide-scrollbars --force-device-scale-factor=2.4 \
    --window-size=375,667 --screenshot="$PWD/_shots/$id.raw.png" "file://$PWD/$f" 2>/dev/null
done
H=$(python3 -c "import json;m=json.load(open('manifest.json'));print(120+40*2+667*len(m))")
google-chrome --headless=new --no-sandbox --disable-gpu --hide-scrollbars --force-device-scale-factor=2 \
  --window-size=375,$H --screenshot="$PWD/angus-wireframe-fullpage.png" "file://$PWD/_shot_full.html" 2>/dev/null
python3 - <<'P'
import json
from PIL import Image
m=json.load(open('manifest.json'))
for i,x in enumerate(m,1):
    im=Image.open(f"_shots/{x['id']}.raw.png").convert("RGB")
    print(x['id'], im.size, end='; ')
    im=im.resize((900,1600), Image.LANCZOS)
    im.save(f"screens/{i:02d}_{x['label'].split(' · ')[1]}.png")
print(); print('full', Image.open('angus-wireframe-fullpage.png').size)
P
