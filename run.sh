#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"
python3 tools/check_sheet_names.py job_config.json
python3 build_xlsx.py
python3 test/test_xlsx.py
node test/test_twilio_notify.js
python3 tools/make_screens.py
python3 tools/render_readme.py
# 最初のスクリーンショットのすぐ下に架空データの断りを1行置く（render_readme.pyが毎回READMEを作り直すため）
python3 - <<'PY'
import re
p='README.md';s=open(p).read();L='All data in this sample is fictional.'
s=s.replace('\n'+L+'\n','\n')
m=re.search(r'!\[[^\]]*\]\([^)]*\)\n',s)
s=s[:m.end()]+'\n'+L+'\n\n'+s[m.end():] if m else s
open(p,'w').write(s)
PY
