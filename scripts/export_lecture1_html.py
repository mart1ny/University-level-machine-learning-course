"""Export lecture 1 as a single offline HTML file."""
from pathlib import Path
import base64
import re

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'materials/Пара1'
OUTPUT = ROOT / 'output/html/Первая_лекция.html'


def inline_css(path):
    css = path.read_text()
    css = re.sub(r'@import\s+url\([^;]+;', '', css)

    css = re.sub(r',?\s*url\(([^)]+)\)\s*format\([^)]+\)',
                 lambda m: m[0] if (path.parent / m[1].strip('\"\'')).exists() else '', css)

    def embed(match):
        name = match[1].strip('\"\'')
        if name.startswith('data:'):
            return match[0]
        asset = (path.parent / name).resolve()
        mime = {'.woff2': 'font/woff2', '.woff': 'font/woff', '.ttf': 'font/ttf',
                '.svg': 'image/svg+xml'}.get(asset.suffix, 'application/octet-stream')
        encoded = base64.b64encode(asset.read_bytes()).decode('ascii')
        return f'url("data:{mime};base64,{encoded}")'

    return re.sub(r'url\(([^)]+)\)', embed, css)


page = (SOURCE / 'index.html').read_text()
page = re.sub(r'<link rel="stylesheet" href="([^"]+)">',
              lambda m: '<style>\n' + inline_css(SOURCE / m[1].split('?')[0]) + '\n</style>', page)
page = re.sub(r'<script src="([^"]+)"></script>',
              lambda m: '<script>\n' + re.sub(r'</script', r'<\\/script',
                (SOURCE / m[1]).read_text(), flags=re.I) + '\n</script>', page)
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
OUTPUT.write_text(page)
print(OUTPUT)
print(f'{OUTPUT.stat().st_size / 1024 / 1024:.1f} MB')
