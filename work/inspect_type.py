from pathlib import Path
s=Path('work/turnover-monitor-preview.html').read_text(encoding='utf-8')
i=s.find('id="kp-type"'); print(s[i-80:i+600].encode('unicode_escape').decode())
