from pathlib import Path
p=Path('work/turnover-monitor-preview.html'); s=p.read_text(encoding='utf-8')
needle="function render(){const q=search.value.trim().toLocaleLowerCase('ru');const t=type.value;let a=data.filter(x=>{"
replacement="function render(){const q=search.value.trim().toLocaleLowerCase('ru');const t=type.value;let a=data.filter(x=>{const expiryYear=String(x.expires||'').slice(-4);if((x.kind==='expired'||x.kind==='expiring')&&!['2026','2027'].includes(expiryYear))return false;"
if needle not in s: raise SystemExit('render marker not found')
s=s.replace(needle,replacement)
p.write_text(s,encoding='utf-8')
