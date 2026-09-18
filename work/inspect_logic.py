from pathlib import Path
s=Path('work/turnover-monitor-preview.html').read_text(encoding='utf-8')
for q in ['addEventListener','const typ=','kp-type']:
 i=s.rfind(q); print('\n---',q,i); print(s[i-900:i+1400].encode('unicode_escape').decode())
