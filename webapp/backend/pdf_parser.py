"""PDF 文本提取（纯标准库，零第三方依赖）。

面向超声报告 PDF（由工作站/报告系统导出的文本型 PDF）：
· 解析常规对象与 PDF 1.5 对象流（/Type /ObjStm）
· 解压 FlateDecode 内容流
· 解析 ToUnicode CMap，还原中文 CID 字体文本
· 按 /Pages 树顺序逐页输出纯文本

不支持：加密 PDF（需密码）、扫描件图片 OCR、纯图片页（无文本层）。
"""
from __future__ import annotations

import re
import zlib

# ============================ 基础对象解析 ============================

_OBJ_HEAD = re.compile(rb"(?<![0-9])(\d+)\s+(\d+)\s+obj\b")
_STREAM_KW = re.compile(rb"stream\r?\n")
_REF = re.compile(rb"(\d+)\s+\d+\s+R")


def _split_stream(body: bytes) -> tuple[bytes, bytes | None]:
    """把对象体拆成 (字典文本, 流字节)。"""
    m = _STREAM_KW.search(body)
    if not m:
        return body, None
    head, rest = body[:m.start()], body[m.end():]
    lm = re.search(rb"/Length\s+(\d+)", head)
    if lm:
        n = int(lm.group(1))
        if 0 <= n <= len(rest):
            return head, rest[:n]
    end = rest.find(b"endstream")
    return head, rest[:end if end != -1 else len(rest)]


def _decode_stream(stream: bytes, head: bytes) -> bytes:
    """按 Filter 解码流（支持 FlateDecode，含 PNG 预测器）。"""
    if not stream:
        return b""
    if b"FlateDecode" not in head and b"/Fl " not in head:
        return stream
    raw = b""
    try:
        raw = zlib.decompress(stream)
    except zlib.error:
        try:
            d = zlib.decompressobj()
            raw = d.decompress(stream)
        except zlib.error:
            return b""
    pred = re.search(rb"/Predictor\s+(\d+)", head)
    cols = re.search(rb"/Columns\s+(\d+)", head)
    if pred and int(pred.group(1)) >= 10 and cols:
        raw = _undo_png_predictor(raw, int(cols.group(1)))
    return raw


def _undo_png_predictor(data: bytes, columns: int) -> bytes:
    """还原 PNG 行预测器（Predictor >= 10），多用于 xref 流与图像流。"""
    out, prev = [], bytearray(columns)
    step = columns + 1
    for i in range(0, len(data) - step + 1, step):
        ftype = data[i]
        row = bytearray(data[i + 1:i + step])
        for j in range(columns):
            left = row[j - 1] if j else 0
            up = prev[j]
            ul = prev[j - 1] if j else 0
            if ftype == 1:
                row[j] = (row[j] + left) & 0xFF
            elif ftype == 2:
                row[j] = (row[j] + up) & 0xFF
            elif ftype == 3:
                row[j] = (row[j] + ((left + up) >> 1)) & 0xFF
            elif ftype == 4:
                p = left + up - ul
                pa, pb, pc = abs(p - left), abs(p - up), abs(p - ul)
                pr = left if pa <= pb and pa <= pc else (up if pb <= pc else ul)
                row[j] = (row[j] + pr) & 0xFF
        out.append(bytes(row))
        prev = row
    return b"".join(out)


def _collect_objects(data: bytes) -> dict[int, bytes]:
    """扫描所有 “N 0 obj … endobj”，并展开对象流（ObjStm）。"""
    objs: dict[int, bytes] = {}
    for m in _OBJ_HEAD.finditer(data):
        num = int(m.group(1))
        end = data.find(b"endobj", m.end())
        objs[num] = data[m.end(): end if end != -1 else len(data)]

    for num, body in list(objs.items()):
        head, stream = _split_stream(body)
        if b"/ObjStm" not in head:
            continue
        nm = re.search(rb"/N\s+(\d+)", head)
        fm = re.search(rb"/First\s+(\d+)", head)
        if not (nm and fm):
            continue
        dec = _decode_stream(stream or b"", head)
        if not dec:
            continue
        count, first = int(nm.group(1)), int(fm.group(1))
        nums = dec[:first].split()
        for i in range(count):
            try:
                onum, off = int(nums[2 * i]), int(nums[2 * i + 1])
                nxt = int(nums[2 * i + 3]) if i + 1 < count else len(dec) - first
            except (IndexError, ValueError):
                break
            objs.setdefault(onum, dec[first + off: first + nxt])
    return objs


# ============================ 结构查找 ============================

def _resolve(objs: dict[int, bytes], ref: int) -> bytes:
    return objs.get(ref, b"")


def _inline_dict(text: bytes, key: bytes) -> bytes | None:
    """取 /key 后的内联字典 <<...>> 文本（括号配平）。"""
    m = re.search(rb"/" + key + rb"\s*<<", text)
    if not m:
        return None
    start = m.end() - 2
    depth, i = 0, start
    while i < len(text) - 1:
        if text[i:i + 2] == b"<<":
            depth += 1
            i += 2
            continue
        if text[i:i + 2] == b">>":
            depth -= 1
            i += 2
            if depth == 0:
                return text[start:i]
            continue
        i += 1
    return text[start:]


def _dict_value(objs: dict[int, bytes], text: bytes, key: bytes) -> bytes | None:
    """取 /key 指向的对象体：内联字典直接返回，引用则解析。"""
    m = re.search(rb"/" + key + rb"\s+(\d+)\s+\d+\s+R", text)
    if m:
        return _resolve(objs, int(m.group(1)))
    return _inline_dict(text, key)


def _page_numbers(objs: dict[int, bytes]) -> list[int]:
    """按 /Catalog → /Pages 树顺序返回页面对象号；失败时按对象号排序兜底。"""
    pages: list[int] = []

    def walk(num: int, depth: int = 0):
        if depth > 32 or num in pages:
            return
        body = _resolve(objs, num)
        head, _ = _split_stream(body)
        if b"/Type" in head and b"/Page" in head and b"/Pages" not in head:
            pages.append(num)
            return
        kids = re.search(rb"/Kids\s*\[(.*?)\]", head, re.S)
        if kids:
            for k in _REF.findall(kids.group(1)):
                walk(int(k), depth + 1)

    catalog = next((n for n, b in objs.items() if b"/Type /Catalog" in b), None)
    if catalog is not None:
        root = re.search(rb"/Pages\s+(\d+)\s+\d+\s+R", _resolve(objs, catalog))
        if root:
            walk(int(root.group(1)))
    if not pages:
        pages = sorted(n for n, b in objs.items() if re.search(rb"/Type\s*/Page\b", b))
    return pages


# ============================ 字体与 CMap ============================

def _hex_to_text(h: str) -> str:
    try:
        b = bytes.fromhex(h)
    except ValueError:
        return ""
    if len(b) % 2:
        b += b"\x00"
    try:
        return b.decode("utf-16-be")
    except UnicodeDecodeError:
        return ""


def _parse_cmap(text: bytes) -> dict[int, str]:
    """解析 ToUnicode CMap 的 bfchar / bfrange 映射。"""
    src = text.decode("latin-1", "ignore")
    cmap: dict[int, str] = {}
    for block in re.findall(r"beginbfchar(.*?)endbfchar", src, re.S):
        for s, d in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            cmap[int(s, 16)] = _hex_to_text(d)
    for block in re.findall(r"beginbfrange(.*?)endbfrange", src, re.S):
        for lo, hi, dst in re.findall(
                r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]+>|\[[^\]]*\])", block):
            start, stop = int(lo, 16), int(hi, 16)
            if stop - start > 0xFFFF:
                continue
            if dst.startswith("["):
                items = re.findall(r"<([0-9A-Fa-f]+)>", dst)
                for i, h in enumerate(items):
                    if start + i <= stop:
                        cmap[start + i] = _hex_to_text(h)
            else:
                base = dst.strip("<>")
                base_bytes = bytes.fromhex(base) if base else b""
                for i in range(stop - start + 1):
                    raw = bytearray(base_bytes)
                    if raw:
                        # 末字节按序号递增（Unicode 连续码位）
                        val = _hex_to_text(base)
                        if val:
                            cmap[start + i] = chr(ord(val[-1]) + i) if len(val) == 1 else val
                        else:
                            _ = raw
                            cmap[start + i] = _hex_to_text(base)
    return cmap


class _Font:
    __slots__ = ("cmap", "two_byte", "encoding", "subtype")

    def __init__(self, cmap: dict[int, str], two_byte: bool, encoding: str, subtype: str):
        self.cmap = cmap
        self.two_byte = two_byte
        self.encoding = encoding
        self.subtype = subtype

    def decode(self, raw: bytes) -> str:
        if self.cmap:
            if self.two_byte:
                codes = [int.from_bytes(raw[i:i + 2], "big") for i in range(0, len(raw) - 1, 2)]
            else:
                codes = list(raw)
            return "".join(self.cmap.get(c, "") for c in codes)
        # 无 ToUnicode：按编码名退化解码
        enc = self.encoding.upper()
        if "UCS2" in enc or "UTF16" in enc:
            try:
                return raw.decode("utf-16-be")
            except UnicodeDecodeError:
                return ""
        if "GBK" in enc or "GB-EUC" in enc or "GBT" in enc:
            try:
                return raw.decode("gbk")
            except UnicodeDecodeError:
                return ""
        try:
            return raw.decode("cp1252")
        except UnicodeDecodeError:
            return ""


def _font_info(objs: dict[int, bytes], body: bytes) -> _Font:
    head, _ = _split_stream(body)
    subtype = b"Type0" if re.search(rb"/Subtype\s*/Type0", head) else b""
    enc_m = re.search(rb"/Encoding\s*/([A-Za-z0-9_-]+)", head)
    encoding = enc_m.group(1).decode("latin-1") if enc_m else ""
    cmap: dict[int, str] = {}
    tu = re.search(rb"/ToUnicode\s+(\d+)\s+\d+\s+R", head)
    if tu:
        tbody = _resolve(objs, int(tu.group(1)))
        thead, tstream = _split_stream(tbody)
        cmap = _parse_cmap(_decode_stream(tstream or b"", thead))
    wide = bool(subtype) or any(k > 0xFF for k in cmap) or "IDENTITY" in encoding.upper()
    return _Font(cmap, wide, encoding, subtype.decode())


def _page_fonts(objs: dict[int, bytes], page_body: bytes) -> dict[bytes, _Font]:
    res = _dict_value(objs, page_body, b"Resources")
    if not res:
        return {}
    fontdict = _dict_value(objs, res, b"Font")
    if not fontdict:
        return {}
    fonts: dict[bytes, _Font] = {}
    for name, num in re.findall(rb"/([A-Za-z0-9_.+-]+)\s+(\d+)\s+\d+\s+R", fontdict):
        fonts[name] = _font_info(objs, _resolve(objs, int(num)))
    return fonts


# ============================ 内容流解析 ============================

_DELIM = b"\x00\t\n\f\r ()<>[]{}/%"


def _unescape(raw: bytes) -> bytes:
    out, i = bytearray(), 0
    while i < len(raw):
        c = raw[i]
        if c != 0x5C:
            out.append(c)
            i += 1
            continue
        i += 1
        if i >= len(raw):
            break
        n = raw[i]
        if n in b"nrtbf":
            out.append({0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12}[n])
            i += 1
        elif n in b"()\\":
            out.append(n)
            i += 1
        elif 0x30 <= n <= 0x37:
            j, val = i, 0
            while j < len(raw) and j < i + 3 and 0x30 <= raw[j] <= 0x37:
                val = val * 8 + raw[j] - 0x30
                j += 1
            out.append(val & 0xFF)
            i = j
        elif n in b"\r\n":
            i += 1
        else:
            out.append(n)
            i += 1
    return bytes(out)


def _tokens(s: bytes):
    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c in b"\x00\t\n\f\r ":
            i += 1
            continue
        if c == 0x25:  # %
            j = s.find(b"\n", i)
            i = n if j < 0 else j + 1
            continue
        if c == 0x28:  # (
            depth, j, buf = 1, i + 1, bytearray()
            while j < n:
                ch = s[j]
                if ch == 0x5C:
                    buf += s[j:j + 2]
                    j += 2
                    continue
                if ch == 0x28:
                    depth += 1
                elif ch == 0x29:
                    depth -= 1
                    if depth == 0:
                        break
                buf.append(ch)
                j += 1
            yield ("str", _unescape(bytes(buf)))
            i = j + 1
            continue
        if c == 0x3C:
            if i + 1 < n and s[i + 1] == 0x3C:
                yield ("dict", None)
                i += 2
                continue
            j = s.find(b">", i)
            hx = re.sub(rb"[^0-9A-Fa-f]", b"", s[i + 1:j] if j != -1 else b"")
            if len(hx) % 2:
                hx += b"0"
            yield ("str", bytes.fromhex(hx.decode()) if hx else b"")
            i = j + 1 if j != -1 else n
            continue
        if c == 0x3E and i + 1 < n and s[i + 1] == 0x3E:
            yield ("dictend", None)
            i += 2
            continue
        if c == 0x5B:
            yield ("[", None)
            i += 1
            continue
        if c == 0x5D:
            yield ("]", None)
            i += 1
            continue
        if c == 0x2F:  # /Name
            j = i + 1
            while j < n and s[j] not in _DELIM:
                j += 1
            yield ("name", s[i + 1:j])
            i = j
            continue
        j = i
        while j < n and s[j] not in _DELIM:
            j += 1
        word = s[i:j]
        if re.fullmatch(rb"[+-]?(\d+\.?\d*|\.\d+)", word or b""):
            yield ("num", float(word))
        else:
            yield ("op", word)
        i = j


def _content_text(content: bytes, fonts: dict[bytes, _Font]) -> str:
    """解释内容流中的文本操作符，按行还原文本。"""
    buf: list[str] = []
    stack: list[tuple] = []
    arrays: list[list[tuple]] = []
    font: _Font | None = None
    size = 12.0
    x = y = 0.0
    last_y: float | None = None

    def newline():
        if buf and buf[-1] != "\n":
            buf.append("\n")

    def emit(raw: bytes):
        nonlocal last_y
        if not raw:
            return
        if last_y is not None and abs(y - last_y) > 1.0:
            newline()
        last_y = y
        buf.append(font.decode(raw) if font else _fallback_decode(raw))

    for kind, val in _tokens(content):
        if kind in ("num", "name", "str"):
            item = (kind, val)
            (arrays[-1] if arrays else stack).append(item)
            continue
        if kind == "dict":
            stack.append(("dict", None))
            continue
        if kind == "dictend":
            continue
        if kind == "[":
            arrays.append([])
            continue
        if kind == "]":
            arr = arrays.pop() if arrays else []
            item = ("array", arr)
            (arrays[-1] if arrays else stack).append(item)
            continue
        op = val
        if op == b"BT":
            stack.clear()
            continue
        if op == b"ET":
            stack.clear()
            continue
        if op == b"Tf" and len(stack) >= 2:
            font = fonts.get(stack[-2][1])
            size = float(stack[-1][1])
        elif op in (b"Td", b"TD") and len(stack) >= 2:
            tx, ty = float(stack[-2][1]), float(stack[-1][1])
            x, y = x + tx, y + ty
        elif op == b"Tm" and len(stack) >= 6:
            x, y = float(stack[-2][1]), float(stack[-1][1])
        elif op == b"T*":
            y -= size * 1.2
        elif op == b"Tj" and stack and stack[-1][0] == "str":
            emit(stack[-1][1])
        elif op == b"'" and stack and stack[-1][0] == "str":
            y -= size * 1.2
            emit(stack[-1][1])
        elif op == b'"' and stack and stack[-1][0] == "str":
            y -= size * 1.2
            emit(stack[-1][1])
        elif op == b"TJ" and stack and stack[-1][0] == "array":
            for k, v in stack[-1][1]:
                if k == "str":
                    emit(v)
                elif k == "num" and v <= -120:
                    buf.append(" ")
        stack.clear()

    return "".join(buf)


def _fallback_decode(raw: bytes) -> str:
    for enc in ("utf-8", "gbk", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return ""


# ============================ 对外入口 ============================

_LABEL_RE = re.compile(
    r"^\s*(?:报告号|检查号|检查项目|检查部位|项目名称|姓名|性别|年龄|科室|床号|"
    r"超声所见|超声提示|超声结论|结论|印象|诊断|医师|日期|时间)\s*[：:]"
)
# 行尾为这些标点说明句子已结束，否则视为换行折行，与下一行合并
_END_PUNCT = "。；;！？!?：:"


def _merge_wrapped(text: str) -> str:
    """还原 PDF 折行：非句末标点且下一行不是字段标签时，与下一行直接合并。"""
    out: list[str] = []
    for line in text.splitlines():
        cur = line.strip()
        if not cur:
            out.append("")
            continue
        prev = out[-1] if out else ""
        if prev and prev and not _LABEL_RE.match(prev) and not _LABEL_RE.match(cur) \
                and prev[-1] not in _END_PUNCT:
            out[-1] = prev + cur
        else:
            out.append(cur)
    return "\n".join(out).strip()


def extract_pages(data: bytes) -> tuple[list[str], list[str]]:
    """返回 (每页纯文本列表, 警告列表)；非 PDF 或加密时抛出 ValueError。"""
    if not data.lstrip()[:5].startswith(b"%PDF"):
        raise ValueError("不是有效的 PDF 文件（缺少 %PDF 文件头）")
    if re.search(rb"/Encrypt\s", data):
        raise ValueError("PDF 已加密（需要密码），暂不支持解析，请导出未加密版本")

    warnings: list[str] = []
    objs = _collect_objects(data)
    if not objs:
        raise ValueError("PDF 结构异常：未找到任何对象")

    pages: list[str] = []
    for num in _page_numbers(objs):
        body = _resolve(objs, num)
        head, _ = _split_stream(body)
        contents = []
        cm = re.search(rb"/Contents\s+(\d+)\s+\d+\s+R", head)
        if cm:
            contents = [int(cm.group(1))]
        else:
            arr = re.search(rb"/Contents\s*\[(.*?)\]", head, re.S)
            if arr:
                contents = [int(x) for x in _REF.findall(arr.group(1))]
        if not contents:
            warnings.append("存在无内容流的页面，已跳过")
            continue
        fonts = _page_fonts(objs, head)
        merged = b"\n".join(
            _decode_stream(*(lambda b: (_split_stream(b)[1] or b"", _split_stream(b)[0]))(_resolve(objs, c)))
            for c in contents
        )
        text = _content_text(merged, fonts).strip()
        if text:
            text = re.sub(r"[ \t]{2,}", " ", text)
            text = _merge_wrapped(text)
            pages.append(text)
        else:
            warnings.append("检测到无文本层页面（可能为扫描图片），已跳过")

    return pages, warnings
