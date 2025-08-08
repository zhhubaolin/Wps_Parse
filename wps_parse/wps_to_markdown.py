"""
把 .wps 文件转化成 .md
"""

import olefile
import re
import unicodedata
from pathlib import Path


# ---------- 工具函数 ----------
def read_plain_text(wps_path: Path) -> str:
    """从 OLE 容器取出 WordDocument 流并做简单清理"""
    if not olefile.isOleFile(wps_path):
        raise RuntimeError(f'{wps_path} 不是 OLE')

    with olefile.OleFileIO(wps_path) as ole:
        if not ole.exists('WordDocument'):
            raise RuntimeError('找不到 WordDocument 流')
        data = ole.openstream('WordDocument').read()

    text = data.decode('utf-16le', errors='ignore')
    # 清理控制字符 / 私用区
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F'
                  r'\uFFFE\uFFFF\u0000-\u0008\u000B\u000C\u000E-\u001F'
                  r'\uD800-\uDFFF]', '', text)
    return re.sub(r'\r\n?', '\n', text)


def is_readable(text: str, threshold: float = 0.8) -> bool:
    """
    判断一段文本是否“可读”。
    可打印字符（字母/数字/汉字/标点/空格）比例 >= threshold 视为可读。
    """
    if not text:
        return False
    printable = 0
    for ch in text:
        # 空格也算可读
        if ch.isspace() or unicodedata.category(ch)[0] in ('L', 'N', 'P', 'S'):
            printable += 1
    return printable / len(text) >= threshold


# ---------- 主逻辑 ----------
def wps_to_md(src: Path, dst: Path, readability_threshold: float = 1.0):
    text = read_plain_text(src)

    dst.parent.mkdir(parents=True, exist_ok=True)
    with dst.open('w', encoding='utf-8') as f:
        # 用文件名作为一级标题
        f.write(f"# {src.stem}\n\n")

        for para in filter(None, text.split('\n')):
            cleaned = para.strip()
            if cleaned and is_readable(cleaned, readability_threshold):
                f.write(f"{cleaned}\n\n")

    print(f'搞定！转化完成，已生成 {dst}')


if __name__ == '__main__':
    src = Path('input.wps')
    dst = Path('output.md')
    wps_to_md(src, dst)
