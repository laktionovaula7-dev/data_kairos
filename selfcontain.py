# -*- coding: utf-8 -*-
"""Делает собранный HTML самодостаточным: картинки, шрифты и three.js встраиваются внутрь файла.
Файл открывается двойным кликом из любой папки и без интернета (как требует PRODUCT.md).

Что встраивается:
  * <img src="assets/...">                      -> data:-URI
  * фото менеджеров assets/overview_v2/mgr/*    -> window.__MGR[имя] (путь собирается в JS динамически)
  * <script src="assets/three.min.js">          -> inline
  * <link ... fonts.googleapis.com ...>         -> @font-face с шрифтами из assets/fonts (woff2, base64)
"""
import base64, re
from pathlib import Path

_MIME = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.svg': 'image/svg+xml',
         '.webp': 'image/webp', '.gif': 'image/gif', '.woff2': 'font/woff2'}

# подмножества шрифтов, которые нужны интерфейсу (вьетнамский не используем)
_SUBSETS = ('latin', 'latin-ext', 'cyrillic', 'cyrillic-ext')
_FAMILIES = {'Montserrat': '400 700', 'Unbounded': '600 800'}


def _data_uri(path: Path) -> str:
    mime = _MIME.get(path.suffix.lower(), 'application/octet-stream')
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


def _font_css(root: Path) -> str:
    """@font-face из assets/fonts/fonts.css (его отдаёт Google Fonts) + локальные woff2."""
    fdir = root / 'assets' / 'fonts'
    css_path = fdir / 'fonts.css'
    if not css_path.exists():
        return ''
    css = css_path.read_text(encoding='utf-8')
    out, seen = [], set()
    for sub, body in re.findall(r'/\*\s*([\w-]+)\s*\*/\s*@font-face\s*\{([^}]*)\}', css):
        if sub not in _SUBSETS:
            continue
        fam = re.search(r"font-family:\s*'([^']+)'", body).group(1)
        if fam not in _FAMILIES:
            continue
        url = re.search(r'url\(([^)]+)\)', body).group(1)
        rng = re.search(r'unicode-range:\s*([^;]+);', body).group(1)
        local = fdir / f"{fam}-{sub}-{url.rsplit('/', 1)[-1]}"
        if not local.exists() or (fam, sub) in seen:
            continue
        seen.add((fam, sub))
        out.append("@font-face{font-family:'%s';font-style:normal;font-weight:%s;font-display:swap;"
                   "src:url(%s) format('woff2');unicode-range:%s}" % (fam, _FAMILIES[fam], _data_uri(local), rng))
    return '<style>' + ''.join(out) + '</style>' if out else ''


def make_self_contained(html: str, root: Path) -> str:
    root = Path(root)

    # 1. шрифты: вместо ссылки на Google Fonts — встроенные @font-face
    fcss = _font_css(root)
    if fcss:
        html = re.sub(r'<link[^>]*fonts\.googleapis\.com[^>]*>', lambda m: fcss, html, count=1)
        # скрипт «Обзора» дозагружает те же шрифты по сети — теперь они уже внутри, лишний запрос убираем
        html = re.sub(r"https://fonts\.googleapis\.com/[^'\"]*", 'data:text/css,', html)

    # 2. статические <img src="assets/..."> -> data:-URI
    def _img(m):
        p = root / 'assets' / m.group(1)
        return f'src="{_data_uri(p)}"' if p.exists() else m.group(0)
    html = re.sub(r'src="assets/([A-Za-z0-9_./-]+\.(?:png|jpe?g|svg|webp|gif))"', _img, html)

    # 3. фото менеджеров: путь собирается в JS -> таблица window.__MGR
    mgr_dir = root / 'assets' / 'overview_v2' / 'mgr'
    if mgr_dir.exists():
        table = {p.stem: _data_uri(p) for p in sorted(mgr_dir.glob('*.png'))}
        pat = r'src="assets/overview_v2/mgr/\'\+([A-Za-z_][\w.]*)\+\'\.png"'
        html = re.sub(pat, lambda m: 'src="\'+(window.__MGR[' + m.group(1) + ']||\'\')+\'"', html)
        import json
        tag = '<script>window.__MGR=' + json.dumps(table) + ';</script>'
        html = html.replace('<body>', '<body>' + tag, 1) if '<body>' in html else tag + html

    # 4. three.js inline
    three = root / 'assets' / 'three.min.js'
    if three.exists():
        js = three.read_text(encoding='utf-8').replace('</script', '<\\/script')
        html = html.replace('<script src="assets/three.min.js"></script>', '<script>' + js + '</script>', 1)
    return html
