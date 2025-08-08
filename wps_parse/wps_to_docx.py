"""
把 .wps 文件转化成 .docx
"""
import olefile
from docx import Document
import re
import unicodedata
from pathlib import Path


def read_plain_text(wps_path: Path) -> str:
    if not olefile.isOleFile(wps_path):
        raise RuntimeError(f'{wps_path} 不是 OLE')

    with olefile.OleFileIO(wps_path) as ole:
        if not ole.exists('WordDocument'):
            raise RuntimeError('找不到 WordDocument 流')
        data = ole.openstream('WordDocument').read()

    # 解码
    text = data.decode('utf-16le', errors='ignore')

    # 清洗数据：NULL、控制符、surrogate、私用区全部干掉
    # 只保留可打印字符 + 正常换行
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F'
                  r'\uFFFE\uFFFF\u0000-\u0008\u000B\u000C\u000E-\u001F'
                  r'\uD800-\uDFFF]', '', text)

    # 统一换行
    return re.sub(r'\r\n?', '\n', text)


def is_readable(text: str, threshold: float = 0.8) -> bool:
    """
    判断一段文本是否"可读"。
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


def wps_to_docx(src: Path, dst: Path, readability_threshold: float = 1.0):
    text = read_plain_text(src)
    doc = Document()
    for para in filter(None, text.split('\n')):
        cleaned = para.strip()
        if cleaned and is_readable(cleaned, readability_threshold):  # 使用is_readable过滤乱码
            doc.add_paragraph(cleaned)
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst)
    print(f'搞定！转化完成，已生成 {dst}')


if __name__ == '__main__':
    src = Path('input.wps')
    dst = Path('output.docx')
    wps_to_docx(src, dst)
