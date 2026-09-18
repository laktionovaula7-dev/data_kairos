from pathlib import Path
s=Path('work/turnover-monitor-preview.html').read_text(encoding='utf-8');i=s.find('id="kp-category"');print(s[i:i+1000].encode('unicode_escape').decode())
