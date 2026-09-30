# Ghép tool_src.html + thư viện + mã nguồn thành 1 file HTML offline duy nhất.
# Chạy:  py build.py
import os, re, sys
sys.stdout.reconfigure(encoding='utf-8')
here = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(here, 'tool_src.html'), encoding='utf-8').read()


def safe_js(code):
    """Tránh chuỗi làm hỏng bộ phân tích <script> của HTML; đổi ký tự non-ASCII sang \\uXXXX."""
    for bad in ('<!--', '<script', '</script'):
        code = re.sub(re.escape(bad), lambda m: r'\x3c' + m.group(0)[1:], code, flags=re.I)
    out = []
    for ch in code:
        o = ord(ch)
        if o < 128:
            out.append(ch)
        elif o < 0x10000:
            out.append('\\u%04x' % o)
        else:
            o -= 0x10000
            out.append('\\u%04x\\u%04x' % (0xD800 + (o >> 10), 0xDC00 + (o & 0x3FF)))
    return ''.join(out)


def lib(m):
    return '<script>' + safe_js(open(os.path.join(here, 'libs', m.group(1)), encoding='utf-8').read()) + '</script>'


def app(_):
    code = '\n'.join(open(os.path.join(here, f), encoding='utf-8').read() for f in ['core.js', 'charts.js', 'excel.js', 'app.js'])
    return '<script>\n' + code + '\n</script>'


import base64
logo = 'data:image/png;base64,' + base64.b64encode(open(os.path.join(here, 'assets', 'logo-nts.png'), 'rb').read()).decode()
src = src.replace('__LOGO__', logo)
out = re.sub(r'<script>/\*LIB:([^*]+)\*/</script>', lib, src)
out = re.sub(r'<script>/\*APP:[^*]+\*/</script>', app, out)
dst = os.path.join(os.path.dirname(here), 'Tool tổng hợp KPI.html')
open(dst, 'w', encoding='utf-8').write(out)
# bản index.html để chạy trực tiếp qua GitHub Pages
open(os.path.join(os.path.dirname(here), 'index.html'), 'w', encoding='utf-8').write(out)
print('OK ->', dst, round(len(out.encode('utf-8')) / 1024), 'KB')
