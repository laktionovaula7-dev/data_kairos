import re
from pathlib import Path
s=Path('work/turnover-monitor-preview.html').read_text(encoding='utf-8')
for m in re.finditer(r'<div class="kp-metric-label">(.*?)</div><div class="kp-metric-value">(.*?)</div>',s): print((m.group(1)+' | '+m.group(2)).encode('unicode_escape').decode())
