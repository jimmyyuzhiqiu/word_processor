
# word_processor.py
import os
import re
import html
import win32com.client as win32

# 假列表前缀： 1. / 2) / （3） / 1、 以及 - • * 等
NUM_PREFIX = re.compile(r"^\s*(?:\d+\s*[.)、]|[\(\（]\s*\d+\s*[\)\）])\s+")
BUL_PREFIX = re.compile(r"^\s*[-–—•●·*]\s+")

# 多种空白（含全角空格/nbsp）
RE_MULTI_SPACE = re.compile(r"[ \u00A0\u2002\u2003\u2009\u3000]{2,}")


def normalize_text(s: str, tab_to_space: bool = True, compress_spaces: bool = True) -> str:
    """清理：HTML实体、Tab->空格、全角空格归一、连续空格压缩、去首尾空白"""
    if s is None:
        return ""
    s = html.unescape(s)

    if tab_to_space:
        s = s.replace("\t", " ")

    # 全角空格/nbsp -> 普通空格
    s = s.replace("\u3000", " ").replace("\u00A0", " ")

    if compress_spaces:
        s = RE_MULTI_SPACE.sub(" ", s)

    return s.strip()


def detect_fake_list(text: str):
    """返回 ('number'/'bullet'/None, stripped_text)"""
    m = NUM_PREFIX.match(text)
    if m:
        return "number", text[m.end():].strip()

    m = BUL_PREFIX.match(text)
    if m:
        return "bullet", text[m.end():].strip()

    return None, text


def iter_paragraphs_safe(range_obj):
    """
    COM 集合遍历更稳：用 Count + Item(i)
    Word 的 Paragraphs 索引从 1 开始
    """
    paras = range_obj.Paragraphs
    count = paras.Count
    for i in range(1, count + 1):
        yield paras.Item(i)


def compress_blank_lines_in_range(range_obj, keep_max_blank_lines: int):
    """把一个 Range 里的连续空行压到 keep_max_blank_lines（倒序删更安全）"""
    if keep_max_blank_lines < 0:
        return

    paras = list(iter_paragraphs_safe(range_obj))
    blank_run = 0

    for p in reversed(paras):
        txt = p.Range.Text or ""
        # Word 段落末尾通常带 "\r"
        content = txt[:-1].strip() if txt.endswith("\r") else txt.strip()
        is_blank = (content == "")

        if is_blank:
            blank_run += 1
            if blank_run > keep_max_blank_lines:
                p.Range.Delete()
        else:
            blank_run = 0


def apply_list_format(p, list_type: str, prev_list_template=None):
    """
    真列表：
    - 第一项 ApplyNumberDefault / ApplyBulletDefault
    - 后续项 ApplyListTemplate(prev_template, ContinuePreviousList=True) 连起来
    """
    lf = p.Range.ListFormat

    if prev_list_template is None:
        if list_type == "number":
            lf.ApplyNumberDefault()
        else:
            lf.ApplyBulletDefault()
        return lf.ListTemplate

    # Continue previous list
    lf.ApplyListTemplate(prev_list_template, True)
    return lf.ListTemplate


def process_range(
    range_obj,
    *,
    keep_max_blank_lines: int = 1,
    tab_to_space: bool = True,
    compress_spaces: bool = True
):
    """清理一个 Range：空格/tab + 假列表转真列表 + 压缩空行"""
    prev_type = None
    prev_template = None

    for p in iter_paragraphs_safe(range_obj):
        pr = p.Range
        raw = pr.Text or ""
        if not raw:
            continue

        has_para_mark = raw.endswith("\r")
        content = raw[:-1] if has_para_mark else raw
        content = normalize_text(content, tab_to_space=tab_to_space, compress_spaces=compress_spaces)

        if content == "":
            # 空段落：清空内容（保留段落符）
            r2 = pr.Duplicate
            if has_para_mark:
                r2.End = r2.End - 1
            r2.Text = ""
            prev_type = None
            prev_template = None
            continue

        list_type, stripped = detect_fake_list(content)

        # 写回：只替换内容，不动段落符
        r2 = pr.Duplicate
        if has_para_mark:
            r2.End = r2.End - 1
        r2.Text = stripped if list_type else content

        # 应用真列表
        if list_type in ("number", "bullet"):
            if list_type != prev_type:
                prev_template = apply_list_format(p, list_type, None)
            else:
                prev_template = apply_list_format(p, list_type, prev_template)
            prev_type = list_type
        else:
            prev_type = None
            prev_template = None

    compress_blank_lines_in_range(range_obj, keep_max_blank_lines)


def process_document(
    input_path: str,
    output_path: str,
    *,
    keep_max_blank_lines: int = 1,
    tab_to_space: bool = True,
    compress_spaces: bool = True,
    process_headers_footers: bool = True
):
    """
    处理单个文件（.doc/.docx 都可由 Word 打开）并导出到 output_path
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    word = None
    doc = None

    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(input_path)

        # 正文
        process_range(
            doc.Content,
            keep_max_blank_lines=keep_max_blank_lines,
            tab_to_space=tab_to_space,
            compress_spaces=compress_spaces
        )

        # 页眉/页脚（可选）
        if process_headers_footers:
            for si in range(1, doc.Sections.Count + 1):
                sec = doc.Sections(si)
                # 1 = Primary header/footer
                try:
                    process_range(
                        sec.Headers(1).Range,
                        keep_max_blank_lines=keep_max_blank_lines,
                        tab_to_space=tab_to_space,
                        compress_spaces=compress_spaces
                    )
                except Exception:
                    pass

                try:
                    process_range(
                        sec.Footers(1).Range,
                        keep_max_blank_lines=keep_max_blank_lines,
                        tab_to_space=tab_to_space,
                        compress_spaces=compress_spaces
                    )
                except Exception:
                    pass

        # 保存（按输出后缀）
        ext = os.path.splitext(output_path)[1].lower()
        if ext == ".docx":
            doc.SaveAs(output_path, FileFormat=12)  # wdFormatXMLDocument
        elif ext == ".doc":
            doc.SaveAs(output_path, FileFormat=0)   # wdFormatDocument
        else:
            doc.SaveAs(output_path + ".docx", FileFormat=12)

    finally:
        # 确保关闭 doc / 退出 Word（不留后台进程）
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass