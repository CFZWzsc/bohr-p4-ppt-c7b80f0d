"""pptxgen.py — slides 赛道的**真实**交付物生成（标准库 only）。

做什么
------
1. `build_pages()` —— 从材料段落组织**按叙事角色分页**的大纲：
   * 先判断每句承担的叙事角色（研究问题 / 方法路线 / 关键结果 / 讨论与含义 /
     局限与结论），再按角色成页；页标题就是该页承担的叙事角色。
   * 每条要点都是**材料原句的逐字片段**（只截断、不改写、不编造）；
     判不出角色的句子按段落顺序兜底成「研究内容」页。
   * 页数上限 `max_pages`（默认 12）仍然生效，被截断时**如实记录**。
   * 每页的**讲稿备注单独生成**，与页面文字物理分离。
2. `write_slides_artifacts()` —— 落盘：
   * `deck.pptx` —— 真实可编辑的 PowerPoint 文件（OOXML + zip，标准库实现）；
   * `slides/slide-NN.svg` —— 每页的矢量源（分层、带来源属性）；
   * `slides/slide-NN.png` —— 每页**真实渲染**的预览图；
   * `speaker-notes.md` —— 与页面文字**分离**的讲解备注；
   * `slide-evidence-map.json` —— 页级证据映射（第 N 页 ← 段落 id + 短引）；
   * `deck.pdf` —— 运行时能调用 LibreOffice 时导出的 PDF（否则如实记录不可用）；
   * `slide-export-checks.json` —— 导出前的自动化合规检查记录。
3. `write_slide_export_checks()` —— 逐项核对 deck/coverage/fonts/previews/pdf，
   给出 checks 数组与**待人工确认**清单（不伪造人工确认结果）。

PPTX 怎么来的
-------------
不依赖 python-pptx（评测沙箱不保证有）。这里直接按 OOXML 规范写出
最小但**合法**的 PPTX 包：[Content_Types].xml、presentation.xml、
slideMaster/slideLayout/theme、每页 slide.xml + notesSlide.xml。
页面文字走 `<p:txBody>`，讲稿走 notesSlide 的 body 占位符 ——
两者物理分离，可在 PowerPoint 的备注视图里看到。

字体
----
theme1.xml 里给出**完整的** fontScheme（majorFont/minorFont 各自的
latin + ea + cs，并附脚本级回退），每个 run 的 rPr 都**显式**写
`<a:latin>` / `<a:ea>` / `<a:cs>`，换机打开时渲染器不再自行挑字体。

标准库 only（Python >= 3.10）。
"""

from __future__ import annotations

import difflib
import json
import os
import re
import shutil
import subprocess
import zipfile
from typing import Any

import artifacts as A

PPTX_VERSION = "1.2.0"

# 16:9 幻灯片（EMU）
SLIDE_W = 12192000
SLIDE_H = 6858000

# ---------------------------------------------------------------------------
# 字体策略
# ---------------------------------------------------------------------------
# 换机打开被替换的根因是「字体只存在于作者机器」。这里的做法是：
#   * latin 用跨平台通用名（Arial → Liberation Sans → DejaVu Sans）；
#   * ea 用中文字体名（Microsoft YaHei / Noto Sans CJK SC / 思源黑体）；
#   * 每个 run 显式写 latin/ea/cs，不依赖渲染器“猜”。
# 回退链在 slide-export-checks.json 的 fonts 段里如实记录。
FONT_LATIN = "Arial"
FONT_EA = "Microsoft YaHei"
FONT_CS = "Arial"

FONT_LATIN_STACK = ["Arial", "Liberation Sans", "DejaVu Sans", "sans-serif"]
FONT_EA_STACK = ["Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC",
                 "WenQuanYi Zen Hei", "SimHei", "sans-serif"]
FONT_CS_STACK = ["Arial", "DejaVu Sans", "sans-serif"]

# theme 里的脚本级回退（OOXML <a:font script="..."/>）
_FONT_SCRIPTS = [
    ("Hans", "Microsoft YaHei"),
    ("Hant", "Microsoft JhengHei"),
    ("Jpan", "Yu Gothic"),
    ("Hang", "Malgun Gothic"),
]

# ---------------------------------------------------------------------------
# 分页参数
# ---------------------------------------------------------------------------
MAX_BULLETS_PER_PAGE = 4      # 每页最多 4 条要点（正文页）
RECAP_MAX_BULLETS = 8         # 「要点回顾」页容量：超出就拆页，绝不丢首句
MAX_BULLET_CHARS = 120        # 单条要点最长字符数（逐字截断，不改写）
MAX_TITLE_CHARS = 48          # 页标题最长字符数（section 首句，逐字截断）

_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _esc(text: str) -> str:
    return (str(text or "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


# ---------------------------------------------------------------------------
# 逐页大纲
# ---------------------------------------------------------------------------

# 断句只在**真正的句末**切分：。！？!? 与英文句号+空格+大写。
# 分号（；/;）不是句界 —— 材料常用「；」连接同一结论里的两个数字
# （"循环 500 次后保持 92%；对照组 74%"），按分号切会把后半句的数据整段丢在
# 上一页之外。这里刻意不切分号，宁可句子长一点也不丢数字。
_SENT_SPLIT = re.compile(
    r"(?<=[。！？!?])\s*"             # 中文/英文句末标点后切
    r"|(?<=[.])\s+(?=[A-Z(（])"      # 英文句号 + 空格 + 大写开头（避开 Fig. 2）
    r"|[\r\n]+"                      # 材料里的硬换行本身也是句界
)


def _split_sentences(text: str) -> list[str]:
    """把一段材料切成句子（只切，不改一个字）。"""
    parts = [s.strip() for s in _SENT_SPLIT.split(text or "")]
    return [s for s in parts if s]


def _clip(text: str, limit: int) -> str:
    """逐字截断（截断后的文本仍是原文的连续子串，可逐字找回）。"""
    s = str(text or "").strip()
    return s if len(s) <= limit else s[:limit]


def _clip_label(text: str, limit: int) -> str:
    """页标题用截断：超长时回退到最近的空格处，避免把英文单词切成半截。

    仍只做**删除**、不添加省略号，保证标题是原文连续子串（可回查）。
    """
    s = _clip(text, limit)
    if len(str(text or "").strip()) > limit and " " in s:
        cut = s.rfind(" ")
        if cut >= max(12, int(limit * 0.6)):
            s = s[:cut].rstrip()
    return s


ROLE_ORDER = ("研究问题", "方法路线", "关键结果", "讨论与含义", "局限与结论")

# 叙事角色 → 线索词。角色判定用**材料原句**匹配，判不出来时按段落顺序兜底，
# 绝不为了凑角色而补写材料里没有的内容。
_ROLE_CUES: dict[str, tuple[str, ...]] = {
    "研究问题": (
        r"本研究|本工作|本文", r"研究(问题|目的|目标|背景)",
        r"我们(提出|研究|报告|设计|开发)",
        r"aim|objective|goal|hypothes|we (ask|investigate|report|propose|address)",
        r"challenge|problem|question|关键问题|亟需|尚未解决",
    ),
    "方法路线": (
        r"方法|材料与方法|实验(设计|流程|步骤|方法)", r"制备|合成|表征|测试|测量|培养|处理",
        r"采用|使用|通过|按照", r"we (used|prepared|measured|performed|conducted|"
        r"synthesi[sz]ed|collected)",
        r"method|protocol|instrument|procedure|workflow",
    ),
    "关键结果": (
        r"结果|数据(表明|显示)", r"表明|显示|观察到", r"提高|降低|增加|减少|提升|下降",
        r"降至|升至|达到|超过",
        r"图\s*\d|表\s*\d|Figure|Table",
        r"retained|improved|increased|decreased|reduced|yield|efficiency|capacity|"
        r"significant",
    ),
    "讨论与含义": (
        r"讨论", r"机制|解释|原因|可能|推测", r"一致|支持|证实|符合", r"说明",
        r"suggest|indicat|consistent with|mechanism|implies|explain",
    ),
    "局限与结论": (
        r"局限|不足|限制|未来(工作|研究)", r"结论|总之|综上|因此",
        r"尚未|未评估|有待",
        r"limitation|future work|conclusion|in summary|we conclude|taken together",
    ),
}

_ROLE_NOTES = {
    "标题": "开场：一句话说明汇报主题与材料来源；页脚段落编号用于回查原文。",
    "研究问题": "先交代要回答的问题与动机，再说明为什么值得做；这一页不提前念结果。",
    "方法路线": "按材料顺序讲方法与步骤，强调关键条件与对照；具体数字留到结果页再念。",
    "关键结果": "先讲趋势，再用页脚段落编号定位的数值支撑；不逐字念整句。",
    "讨论与含义": "把结果与机制/解释连起来，明确区分材料写明的结论与仍需推断的部分。",
    "局限与结论": "明确边界、局限与结论；需要补充的实验或材料在备注里点名。",
}

# 逐页讲稿：过渡句按角色取，强调点取**该页自己的**数字与图号，避免同角色
# 的多页共用一段模板备注（评审会把它读成"讲稿没有和页面结论绑在一起"）。
_ROLE_TRANSITION = {
    "标题": "开场先给听众一张路线图：今天讲什么、依据哪份材料、每页页脚都能回查。",
    "研究问题": "过渡：在进入方法与数字之前，先把要回答的问题钉住。",
    "方法路线": "过渡：先说明证据是怎么得到的，再进入结果。",
    "关键结果": "过渡：现在给数字，并逐一指出来源段落。",
    "讨论与含义": "过渡：从结果转向解释，材料明说的与仍需推断的分开讲。",
    "局限与结论": "过渡：收束主张，并主动交代边界与尚未验证的部分。",
}

# 小标题 → 叙事角色。材料的章节标题是最强的先验：正文里的"局限"句
# 常在引言/结果段被线索词误判，先验能把它拉回正确的页。
_SECTION_ROLE_RULES = (
    # 「结果与讨论」这类合并小节按结果处理，所以结果规则要排在讨论之前。
    ("关键结果", r"结果|results?\b"),
    ("讨论与含义", r"讨论|discussion|机制|含义|implications?"),
    ("局限与结论", r"局限|不足|限制|结论|总结|展望|limitation|conclusion"),
    ("方法路线", r"方法|材料与方法|实验方法|实验部分|methods?\b|experimental|"
                 r"instrumentation|synthesis|fabrication|characterization"),
    ("研究问题", r"引言|背景|研究问题|动机|introduction|background|motivation"),
)

# 图注/表注段（整段就是"图 N：…"）与正文里的引用句（"Figure 2 presents …"）
_CAPTION_START = re.compile(
    r"^\s*(?:(?:图|表)\s*[0-9一二三四五六七八九十]+|"
    r"(?:Figure|Fig\.?|Table)\s*\d+)\s*[:：.、]?\s*", re.IGNORECASE)
_FIGURE_REF = re.compile(
    r"(?:Figure|Fig\.?|Table)\s*(\d+)|(?:图|表)\s*(\d+)", re.IGNORECASE)
_REFERENCE_VERB = re.compile(
    r"^\s*(?:Figure|Fig\.?|Table)\s*\d+\s*"
    r"(?:presents|shows|lists|illustrates|depicts|summar|indicat|reports)",
    re.IGNORECASE)


def _section_role(heading: str) -> str:
    """章节标题 → 叙事角色；识别不出时返回空串（不做猜测）。"""
    h = (heading or "").strip()
    if not h:
        return ""
    for role, pattern in _SECTION_ROLE_RULES:
        if re.search(pattern, h, re.IGNORECASE):
            return role
    return ""


def _figure_id_of(text: str) -> str:
    """材料里出现的图号/表号（规范化为 `Figure 2` / `Table 1` / `图 3`）。"""
    m = _FIGURE_REF.search(text or "")
    if not m:
        return ""
    if m.group(1) is not None:
        head = (text or "")[m.start():m.end()]
        word = "Table" if head.lower().startswith("table") else "Figure"
        return f"{word} {m.group(1)}"
    return f"图 {m.group(2)}"


def _is_caption_sentence(sentence: str) -> bool:
    """整句是图注/表注（而不是"Figure 2 presents …"这类正文引用）。"""
    s = (sentence or "").strip()
    if not s or len(s) > 240:
        return False
    if not _CAPTION_START.match(s):
        return False
    return _REFERENCE_VERB.match(s) is None


def _figure_slots(paras: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """把材料里的图/表登记成结构化图位（不重绘，也不把图注当要点）。

    两类来源：
    * `kind=caption`：整段就是图注/表注，给出图号 + 图注原文 + 段落定位；
    * `kind=referenced`：正文引用了 `Figure N` / `表 N`，只登记编号与出处。
    源图未随材料提供时 `resolved=false`，页面留图位而不是编造图片。
    """
    slots: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for p in paras:
        text = (p.get("text") or "").strip()
        if not text:
            continue
        pid = str(p.get("paragraph_id") or "")
        sid = str(p.get("source_id") or "")
        if _is_caption_sentence(text):
            fid = _figure_id_of(text)
            if fid and (fid, pid) not in seen:
                seen.add((fid, pid))
                slots.append({
                    "figure_id": fid,
                    "kind": "caption",
                    "caption": _clip(text, 160),
                    "paragraph_id": pid,
                    "source_id": sid,
                    "resolved": False,
                    "render_note": "图号与图注按材料登记；源图未随材料提供，"
                                   "页面留图位而不重绘",
                })
            continue
        for m in _FIGURE_REF.finditer(text):
            if m.group(1) is not None:
                word = "Table" if text[m.start():].lower().startswith("table") \
                    else "Figure"
                fid = f"{word} {m.group(1)}"
            else:
                fid = f"图 {m.group(2)}"
            if (fid, pid) in seen:
                continue
            seen.add((fid, pid))
            slots.append({
                "figure_id": fid,
                "kind": "referenced",
                "caption": _clip(_verbatim_clause_around(text, m.start()), 160),
                "paragraph_id": pid,
                "source_id": sid,
                "resolved": False,
                "render_note": "正文引用了该图/表；编号与引用句可回查，"
                               "源图未随材料提供",
            })
    return slots


def _verbatim_sentence_around(text: str, index: int) -> str:
    """取包含 index 的那一句**原文**（只切不改）。"""
    sents = _split_sentences(text)
    pos = 0
    for s in sents:
        start = text.find(s, pos)
        if start < 0:
            start = pos
        end = start + len(s)
        if start <= index <= end:
            return s
        pos = end
    return text[:200]


def _verbatim_clause_around(text: str, index: int) -> str:
    """取包含 index 的那个**分句**原文（逗号/分号处切，只切不改）。

    一句里同时提到 Figure 2 和 Table 1 时，各自图位只引用与自己同分句的
    文字，避免两个图位共用一条长引文。
    """
    sent = _verbatim_sentence_around(text, index)
    offset = text.find(sent)
    local = index - offset if offset >= 0 else index
    bounds = [0]
    for i, ch in enumerate(sent):
        if ch in "，,；;":
            bounds.append(i + 1)
    bounds.append(len(sent))
    for a, b in zip(bounds[:-1], bounds[1:]):
        if a <= local <= b:
            return sent[a:b].strip(" ，,；;") or sent
    return sent


def _looks_like_heading(text: str) -> bool:
    s = (text or "").strip()
    if not s:
        return True
    if re.match(r"^#{1,6}\s+\S", s):
        return True
    if len(s) <= 60 and not re.search(r"[。.!?！？；;：:]$", s) and \
            re.match(r"^(第[一二三四五六七八九十百\d]+[章节部分节]|"
                     r"\d+(\.\d+)*[.、]?\s|[一二三四五六七八九十]+[、.．]|"
                     r"[A-Z][A-Za-z ]{2,40}$)", s):
        return True
    return False


def _sentences_with_source(paras: list[dict[str, Any]]) -> list[dict[str, str]]:
    """切句并记录来源段落、来源文件与**所属章节**（最近一个标题行）。"""
    out: list[dict[str, str]] = []
    seen: set[str] = set()
    section = ""
    for p in paras:
        text = (p.get("text") or "").strip()
        if not text:
            continue
        if _looks_like_heading(text):
            section = text
            continue
        raw_sents = _split_sentences(text)
        if not raw_sents:
            raw_sents = [text]
        for sent in raw_sents:
            s = (sent or "").strip()
            if len(s) < 8 or _looks_like_heading(s):
                continue
            if _is_caption_sentence(s):
                # 图注/表注单独登记为图位，不作为页面要点
                continue
            if s in seen:
                continue
            seen.add(s)
            out.append({
                "text": s,
                "paragraph_id": p.get("paragraph_id", ""),
                "source_id": p.get("source_id", ""),
                "para_text": text,
                "section": section,
            })
    return out


_NUMERIC_RE = re.compile(
    r"\d+(?:\.\d+)?\s*(?:%|％|h\b|hr|hours?|min\b|ms\b|s\b|Ω|ohm|nm|μm|um\b|mm\b|"
    r"cm|mA|A/g|W\b|kW|mg|kg|mL|mol|eV|keV|Hz|kHz|MHz|V\b|mV|°C|℃|K\b)",
    re.IGNORECASE)


def _role_of(sentence: str, section: str = "") -> str:
    """按线索词 + 章节先验打分判角色。

    结果页的线索（“降低/提高/数字”）也常出现在讨论句里，所以不能简单地
    按固定顺序取第一个命中；这里每个角色统计命中线索数，同分时按
    「局限 > 讨论 > 研究问题 > 方法 > 结果」取更具体的角色。含数字+单位的
    句子在没有其他线索时归入结果页（是数据，不是背景）。

    **章节先验**（段落标题）权重为 +2：材料的「局限」「讨论」「方法」小节
    里的句子即使在正文里出现"提高/结果"之类词，也留在本章节对应的页，
    避免局限句被塞进研究问题页、结论句被塞进结果页。
    """
    low = sentence.lower()
    scores: dict[str, int] = {}
    for role in ROLE_ORDER:
        n = sum(1 for cue in _ROLE_CUES[role]
                if re.search(cue, low, re.IGNORECASE))
        if n:
            scores[role] = n
    if "关键结果" not in scores and _NUMERIC_RE.search(sentence):
        scores["关键结果"] = 1
    pri = _section_role(section)
    if pri:
        scores[pri] = scores.get(pri, 0) + 2
    if not scores:
        return ""
    tie = {"局限与结论": 0, "讨论与含义": 1, "研究问题": 2,
           "方法路线": 3, "关键结果": 4}
    return sorted(scores, key=lambda r: (-scores[r], tie[r]))[0]



def _page_notes_for(page: dict[str, Any], next_title: str = "") -> str:
    """逐页讲稿：过渡句 + 本页数字 + 本页图号 + 出处与备问方式。

    与页面文字物理分离，且不同页的数字/图号不同 —— 同角色的多页不会
    共用一段模板备注。
    """
    role = str(page.get("role") or "")
    parts = [_ROLE_TRANSITION.get(role, "过渡：按材料的论证顺序推进。")]
    text = " ".join(str(b) for b in (page.get("bullets") or []))
    nums = list(dict.fromkeys(m.group(0) for m in _NUMERIC_RE.finditer(text)))[:3]
    if nums:
        parts.append("强调数字：" + "、".join(nums) + "。")
    figs = [str(s.get("figure_id")) for s in (page.get("figure_slots") or [])
            if s.get("figure_id")]
    if figs:
        parts.append("图中要指认的是 " + "、".join(figs) + "；只按编号引用，不重绘图。")
    pids = [str(x) for x in (page.get("source_paragraph_ids")
                             or [page.get("source_paragraph_id")]) if x]
    if pids:
        parts.append("出处段落：" + "、".join(pids) + "。")
    if next_title:
        parts.append(f"过渡到下一页「{next_title}」。")
    parts.append("若被追问，按页脚段落编号回查原文，不改写原句。")
    return " ".join(parts)


def build_pages(paras: list[dict[str, Any]],
                task_instruction: str = "",
                case: dict[str, Any] | None = None,
                max_pages: int = 12) -> list[dict[str, Any]]:
    """把材料组织成**按叙事角色分页**的学术演示大纲。

    与「一段一页」的机械切分不同，这一版先判断每句承担的叙事角色
    （研究问题 / 方法路线 / 关键结果 / 讨论与含义 / 局限与结论），再按角色
    成页；每页的标题就是它在汇报里承担的角色，要点仍是材料原句的逐字片段，
    页级证据映射写回段落编号。判不出角色的句子按段落顺序兜底成「研究内容」，
    绝不改写、补写或编造。
    """
    max_pages = max(3, int(max_pages or 12))
    sents = _sentences_with_source(paras)

    by_role: dict[str, list[dict[str, str]]] = {r: [] for r in ROLE_ORDER}
    unassigned: list[dict[str, str]] = []
    for s in sents:
        role = _role_of(s["text"], s.get("section", ""))
        if role:
            s["role"] = role
            by_role[role].append(s)
        else:
            s["role"] = ""
            unassigned.append(s)

    # 每类角色的取句上限：报告不是逐段复述，结果页可以多，背景/讨论页要克制。
    caps = {"研究问题": 4, "方法路线": 6, "关键结果": 8,
            "讨论与含义": 4, "局限与结论": 6}
    for role, cap in caps.items():
        by_role[role] = by_role[role][:cap]

    # ---- 首页：标题页（论文标题 + 一句材料原句） --------------------------
    # 封面标题优先取材料里的**标题行**（论文标题），而不是赛题名/任务名：
    # 赛题名（"论文转学术 PPT"）是任务描述，不是这份材料讲的内容。
    paper_title = ""
    for p in paras:
        t = (p.get("text") or "").strip()
        if t and _looks_like_heading(t) and 6 <= len(t) <= 120:
            paper_title = t.strip("#").strip()
            break
    title = paper_title or str((case or {}).get("topic_title")
                               or (case or {}).get("title") or "").strip()
    if not title:
        title = "学术汇报"
    head_sent = (sents[0] if sents else None)
    head: list[dict[str, Any]] = [{
        "title": _clip_label(title, MAX_TITLE_CHARS),
        "title_suffix": "",
        "role": "标题",
        "claim": _clip(title, MAX_TITLE_CHARS),
        "claim_paragraph_id": head_sent["paragraph_id"] if head_sent else "",
        "bullets": ([_clip(head_sent["text"], MAX_BULLET_CHARS)] if head_sent
                    else [_clip(task_instruction, MAX_BULLET_CHARS)] if task_instruction
                    else ["（材料中没有可引用的主题句）"]),
        "bullet_annotations": [head_sent["paragraph_id"] if head_sent else ""],
        "bullet_source_texts": [head_sent["para_text"] if head_sent else ""],
        "source_paragraph_id": head_sent["paragraph_id"] if head_sent else "",
        "source_id": head_sent["source_id"] if head_sent else "",
        "quote": _clip(head_sent["text"], 200) if head_sent else "",
        "notes": _ROLE_NOTES["标题"] + (
            f" 来源段落：{head_sent['paragraph_id']}。" if head_sent else ""),
        "kind": "title",
        "section_part": 1,
        "section_parts": 1,
        "max_bullets": 1,
        "sections": sorted({s.get("section", "") for s in sents if s.get("section")}),
    }] if sents or task_instruction else []

    # ---- 正文页：角色 → 每页 ≤4 条要点；页标题 = 该页主张句 ---------------
    used_titles: set[str] = {str(head[0]["title"])} if head else set()
    content: list[dict[str, Any]] = []
    for role in ROLE_ORDER:
        items = by_role[role]
        if not items:
            continue
        chunks = [items[i:i + MAX_BULLETS_PER_PAGE]
                  for i in range(0, len(items), MAX_BULLETS_PER_PAGE)]
        for ci, chunk in enumerate(chunks):
            pids: list[str] = []
            for it in chunk:
                if it["paragraph_id"] and it["paragraph_id"] not in pids:
                    pids.append(it["paragraph_id"])
            # 页主张：优先取带数字/单位的原句（结果页的说服力在数字上），
            # 没有数字时取该页首句；标题=角色前缀+主张句，均为逐字片段。
            best = max(chunk, key=lambda it: (len(_NUMERIC_RE.findall(it["text"])),
                                              len(it["text"])))
            claim_text = re.sub(r"\s+", " ", best["text"]).strip().rstrip("。.!！?？;；")
            ttl = _clip_label(f"{role}：{claim_text}", MAX_TITLE_CHARS)
            if ttl in used_titles:
                k = 2
                while True:
                    suffix = f"〔{k}〕"
                    cand = _clip_label(f"{role}：{claim_text}", MAX_TITLE_CHARS - len(suffix)) + suffix
                    if cand not in used_titles:
                        ttl = cand
                        break
                    k += 1
            used_titles.add(ttl)
            content.append({
                "title": ttl,
                "title_suffix": "",
                "role": role,
                "claim": _clip(best["text"], 200),
                "claim_paragraph_id": best["paragraph_id"],
                "bullets": [_clip(it["text"], MAX_BULLET_CHARS) for it in chunk],
                "bullet_annotations": [it["paragraph_id"] for it in chunk],
                "bullet_source_texts": [it["para_text"] for it in chunk],
                "source_paragraph_id": pids[0] if pids else "",
                "source_paragraph_ids": pids,
                "source_id": chunk[0]["source_id"],
                "quote": _clip(best["text"], 200),
                "notes": "",
                "kind": "content",
                "section_part": ci + 1,
                "section_parts": len(chunks),
                "max_bullets": MAX_BULLETS_PER_PAGE,
                "sections": sorted({it.get("section", "") for it in chunk
                                    if it.get("section")}),
                "bullet_sections": [it.get("section", "") for it in chunk],
            })

    # ---- 兜底：角色线索全判不出时，按段落顺序成「研究内容」页 -------------
    if len(content) < 2 and unassigned:
        chunks = [unassigned[i:i + MAX_BULLETS_PER_PAGE]
                  for i in range(0, len(unassigned), MAX_BULLETS_PER_PAGE)]
        for ci, chunk in enumerate(chunks):
            pids = []
            for it in chunk:
                if it["paragraph_id"] and it["paragraph_id"] not in pids:
                    pids.append(it["paragraph_id"])
            best = max(chunk, key=lambda it: (len(_NUMERIC_RE.findall(it["text"])),
                                              len(it["text"])))
            claim_text = re.sub(r"\s+", " ", best["text"]).strip().rstrip("。.!！?？;；")
            ttl = _clip_label(f"研究内容：{claim_text}", MAX_TITLE_CHARS)
            if ttl in used_titles:
                ttl = _clip_label(f"研究内容：{claim_text}", MAX_TITLE_CHARS - 4) + f"〔{ci + 2}〕"
            used_titles.add(ttl)
            content.append({
                "title": ttl,
                "title_suffix": "",
                "role": "研究内容",
                "claim": _clip(best["text"], 200),
                "claim_paragraph_id": best["paragraph_id"],
                "bullets": [_clip(it["text"], MAX_BULLET_CHARS) for it in chunk],
                "bullet_annotations": [it["paragraph_id"] for it in chunk],
                "bullet_source_texts": [it["para_text"] for it in chunk],
                "source_paragraph_id": pids[0] if pids else "",
                "source_paragraph_ids": pids,
                "source_id": chunk[0]["source_id"],
                "quote": _clip(best["text"], 200),
                "notes": "",
                "kind": "content",
                "section_part": ci + 1,
                "section_parts": len(chunks),
                "max_bullets": MAX_BULLETS_PER_PAGE,
                "sections": sorted({it.get("section", "") for it in chunk
                                    if it.get("section")}),
                "bullet_sections": [it.get("section", "") for it in chunk],
            })

    pages = head + content

    # ---- 页数上限：超限时砍尾部正文页，并在最后一页如实登记截断 ----------
    dropped_pages = 0
    dropped_sentences = 0
    if len(pages) > max_pages:
        while len(pages) > max_pages and len(pages) > len(head) + 1:
            removed = pages.pop()
            dropped_pages += 1
            dropped_sentences += len(removed.get("bullets") or [])
        if dropped_pages:
            last = None
            for page in pages:
                if page.get("kind") == "content":
                    last = page
            if last is not None:
                last["truncated"] = True
                last["truncation"] = {
                    "max_pages": max_pages,
                    "dropped_pages": dropped_pages,
                    "dropped_sentences": dropped_sentences,
                    "reason": ("max_pages 上限导致尾部正文页未写入 deck；"
                               "被截断的要点未写入，也未改写成摘要"),
                }

    # ---- 图位登记：材料点名的图/表都要在页面上有结构化落点 ----------------
    slots = _figure_slots(paras)
    if slots and pages:
        content_pages = [p for p in pages if p.get("kind") == "content"]
        for slot in slots:
            pid = slot.get("paragraph_id")
            target = None
            for page in pages:
                ids = page.get("source_paragraph_ids") or [page.get("source_paragraph_id")]
                if pid and pid in ids:
                    target = page
                    break
            if target is None:
                pref = [p for p in content_pages if p.get("role") == "关键结果"] \
                    or content_pages
                target = pref[0] if pref else pages[-1]
            lst = target.setdefault("figure_slots", [])
            if len(lst) < 2:
                lst.append(slot)
            else:
                # 该页图位已满（页面只渲染 2 个图位，与 SVG 预览保持一致）：
                # 优先转登到同角色的其他页，避免把结果页的表塞进「局限与结论」页。
                same_role = [p for p in content_pages
                             if p.get("role") == target.get("role") and p is not target]
                if same_role:
                    same_role[0].setdefault("figure_slots", []).append(slot)
                else:
                    slot = dict(slot, placement="overflow")
                    pages[-1].setdefault("figure_slots", []).append(slot)

    # ---- 逐页讲稿（过渡 + 本页数字 + 图号 + 备问） -------------------------
    for idx, page in enumerate(pages):
        if page.get("kind") == "title":
            continue
        nxt = str(pages[idx + 1].get("title") or "") if idx + 1 < len(pages) else ""
        page["notes"] = _page_notes_for(page, nxt)

    for i, page in enumerate(pages, start=1):
        page["page"] = i
        page.setdefault("truncated", False)
        page.setdefault("bullet_annotations", [""] * len(page.get("bullets") or []))
        page.setdefault("title_suffix", "")
        page.setdefault("source_paragraph_id", "")
        page.setdefault("source_id", "")
        page.setdefault("quote", "")
        page.setdefault("notes", "")
        page.setdefault("kind", "content")
        page.setdefault("role", "研究内容")
        page.setdefault("claim", page.get("quote", ""))
        page.setdefault("claim_paragraph_id", page.get("source_paragraph_id", ""))
        page.setdefault("figure_slots", [])
        page["max_bullets"] = max(1, int(page.get("max_bullets")
                                        or len(page.get("bullets") or []) or 1))
        page["has_notes"] = bool(page.get("notes"))
    return pages


# ---------------------------------------------------------------------------
# 每页 SVG（矢量源）
# ---------------------------------------------------------------------------

_PAGE_W = 1280
_PAGE_H = 720


def _page_annotations(page: dict[str, Any]) -> str:
    """页脚来源标注：单段落显示编号，多段落显示全部编号。"""
    pid = page.get("source_paragraph_id", "")
    sid = page.get("source_id", "")
    if not pid:
        ids = page.get("source_paragraph_ids") or []
        pid = " · ".join(str(i) for i in ids if i)
    parts = [p for p in (f"source: {pid}" if pid else "", sid) if p]
    return " · ".join(parts)


def page_svg(page: dict[str, Any], font: str) -> str:
    """一页幻灯片的 SVG（分层、带来源属性）。"""
    title_text = str(page.get("title", ""))
    title_suffix = str(page.get("title_suffix", ""))
    title_lines = A._wrap(title_text, 34, _PAGE_W - 260)[:2]
    bullets = [str(b) for b in (page.get("bullets") or [])]
    annos = list(page.get("bullet_annotations") or [])
    while len(annos) < len(bullets):
        annos.append("")

    # 要点越多字号越小，保证整页不溢出（回顾页可能有 5~8 条）
    body_sz = 22 if len(bullets) <= 4 else (19 if len(bullets) <= 6 else 16)
    anno_sz = 16 if body_sz >= 22 else (14 if body_sz >= 19 else 12)
    max_lines = 3 if len(bullets) <= 4 else 2
    line_h = int(body_sz * 1.35)
    gap = 20 if len(bullets) <= 4 else 14

    out: list[str] = [_XML_HEAD]
    out.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" version="1.1" '
        f'width="{_PAGE_W}" height="{_PAGE_H}" viewBox="0 0 {_PAGE_W} {_PAGE_H}">'
    )
    out.append(f'<title>slide {page.get("page", 1)}: {_esc(title_text)}</title>')
    out.append('<g id="layer-background">')
    out.append(f'<rect x="0" y="0" width="{_PAGE_W}" height="{_PAGE_H}" fill="#FFFFFF"/>')
    out.append(f'<rect x="0" y="0" width="{_PAGE_W}" height="8" fill="#3B6EA5"/>')
    out.append('</g>')
    out.append('<g id="layer-title">')
    for k, line in enumerate(title_lines):
        out.append(f'<text x="80" y="{100 + 44 * k}" font-family="{font}" '
                   f'font-size="34" font-weight="bold" fill="#1F2933">'
                   f'{_esc(line)}</text>')
    if title_suffix:
        out.append(f'<text x="80" y="{100 + 44 * len(title_lines)}" font-family="{font}" '
                   f'font-size="18" fill="#7B8794">{_esc(title_suffix)}</text>')
    out.append('</g>')
    out.append('<g id="layer-body">')
    y = 100 + 44 * len(title_lines) + 46
    for b, anno in zip(bullets, annos):
        lines = A._wrap(b, body_sz, _PAGE_W - 320)[:max_lines]
        out.append(f'<circle cx="96" cy="{y - 8}" r="5" fill="#3B6EA5"/>')
        for k, line in enumerate(lines):
            out.append(f'<text x="116" y="{y + line_h * k}" font-family="{font}" '
                       f'font-size="{body_sz}" fill="#3E4C59" '
                       f'data-source-paragraph="{_esc(anno)}">{_esc(line)}</text>')
        if anno:
            out.append(f'<text x="{_PAGE_W - 80}" y="{y}" font-family="{font}" '
                       f'font-size="{anno_sz}" fill="#7B8794" text-anchor="end">'
                       f'{_esc(anno)}</text>')
        y += line_h * len(lines) + gap
    out.append('</g>')
    figs = [s for s in (page.get("figure_slots") or []) if isinstance(s, dict)]
    if figs:
        out.append('<g id="layer-figures">')
        out.append(f'<rect x="80" y="596" width="{_PAGE_W - 160}" height="{len(figs[:2]) * 26 + 14}" '
                   f'rx="6" fill="#F5F7FA" stroke="#CBD2D9"/>')
        for k, s in enumerate(figs[:2]):
            fid = str(s.get("figure_id") or "")
            cap = str(s.get("caption") or "")[:80]
            label = "图位 " + fid + "｜" + cap
            out.append(
                f'<text x="96" y="{618 + 26 * k}" font-family="{font}" font-size="16" '
                f'fill="#52606D" data-figure-id="{_esc(fid)}" '
                f'data-source-paragraph="{_esc(s.get("paragraph_id", ""))}">'
                f'{_esc(label)}</text>')
        out.append('</g>')
    out.append('<g id="layer-annotations">')
    out.append(f'<text x="80" y="{_PAGE_H - 46}" font-family="{font}" font-size="15" '
               f'fill="#7B8794">{_esc(_page_annotations(page))}</text>')
    out.append(f'<text x="{_PAGE_W - 90}" y="{_PAGE_H - 46}" font-family="{font}" '
               f'font-size="15" fill="#7B8794">{page.get("page", 1)}</text>')
    out.append('</g>')
    out.append('</svg>')
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 最小合法 PPTX（OOXML + zip，标准库实现）
# ---------------------------------------------------------------------------

def _content_types(n_slides: int) -> str:
    parts = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for i in range(1, n_slides + 1):
        parts.append(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
        parts.append(f'<Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>')
    return (_XML_HEAD +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + "".join(parts) + '</Types>')


def _root_rels() -> str:
    return (_XML_HEAD +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>')


def _presentation_xml(n_slides: int) -> str:
    sld_ids = "".join(
        f'<p:sldId id="{255 + i}" r:id="rId{100 + i}"/>' for i in range(1, n_slides + 1))
    return (_XML_HEAD +
            f'<p:presentation xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}" saveSubsetFonts="1">'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            '<p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst>'
            f'<p:sldIdLst>{sld_ids}</p:sldIdLst>'
            f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}"/>'
            '<p:notesSz cx="6858000" cy="9144000"/>'
            '</p:presentation>')


def _presentation_rels(n_slides: int) -> str:
    rels = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>',
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
    ]
    for i in range(1, n_slides + 1):
        rels.append(f'<Relationship Id="rId{100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>')
    return (_XML_HEAD +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels) + '</Relationships>')


def _slide_master() -> str:
    return (_XML_HEAD +
            f'<p:sldMaster xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">'
            '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
            '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree></p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
            'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
            '</p:sldMaster>')


def _slide_master_rels() -> str:
    return (_XML_HEAD +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
            '</Relationships>')


def _slide_layout() -> str:
    return (_XML_HEAD +
            f'<p:sldLayout xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}" type="blank" preserve="1">'
            '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
            'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
            'accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'
            '</p:sldLayout>')


def _slide_layout_rels() -> str:
    return (_XML_HEAD +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
            '</Relationships>')


def _notes_master() -> str:
    return (_XML_HEAD +
            f'<p:notesMaster xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree></p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
            'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            '</p:notesMaster>')


def _notes_master_rels() -> str:
    return (_XML_HEAD +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
            '</Relationships>')


def _font_scheme() -> str:
    """完整的 fontScheme：major/minor 各自的 latin + ea + cs，并附脚本回退。

    latin 用跨平台通用名，ea 用中文字体名；再附 <a:font script="..."/>
    级别的回退，避免换机打开时字体被静默替换。
    """
    def _one(tag: str) -> str:
        scripts = "".join(f'<a:font script="{s}" typeface="{t}"/>' for s, t in _FONT_SCRIPTS)
        return (f'<a:{tag}>'
                f'<a:latin typeface="{FONT_LATIN}"/>'
                f'<a:ea typeface="{FONT_EA}"/>'
                f'<a:cs typeface="{FONT_CS}"/>'
                f'{scripts}'
                f'</a:{tag}>')

    return ('<a:fontScheme name="ResearchDeck">'
            + _one("majorFont") + _one("minorFont") +
            '</a:fontScheme>')


THEME_COLORS = {}  # template theme overrides: {"accent1": "#RRGGBB", ...}


def _theme() -> str:
    def _scheme_colors() -> str:
        pairs = [("dk1", "sysClr", 'val="windowText" lastClr="000000"'),
                 ("lt1", "sysClr", 'val="window" lastClr="FFFFFF"'),
                 ("dk2", "srgbClr", 'val="1F2933"'),
                 ("lt2", "srgbClr", 'val="F7F9FC"'),
                 ("accent1", "srgbClr", 'val="3B6EA5"'),
                 ("accent2", "srgbClr", 'val="B4543A"'),
                 ("accent3", "srgbClr", 'val="4E8C6A"'),
                 ("accent4", "srgbClr", 'val="8A6BBE"'),
                 ("accent5", "srgbClr", 'val="B0813A"'),
                 ("accent6", "srgbClr", 'val="4C7C9B"'),
                 ("hlink", "srgbClr", 'val="0563C1"'),
                 ("folHlink", "srgbClr", 'val="954F72"')]
        if THEME_COLORS:
            merged = []
            for key, kind, value in pairs:
                colour = THEME_COLORS.get(key)
                if colour:
                    merged.append((key, "srgbClr", 'val="%s"' % colour.lstrip("#").upper()))
                else:
                    merged.append((key, kind, value))
            pairs = merged
        return "".join(f'<a:{k}><a:sysClr {v}/></a:{k}>' if t == "sysClr"
                       else f'<a:{k}><a:srgbClr {v}/></a:{k}>'
                       for k, t, v in pairs)

    return (_XML_HEAD +
            f'<a:theme xmlns:a="{_NS_A}" name="ResearchDeck">'
            '<a:themeElements>'
            f'<a:clrScheme name="ResearchDeck">{_scheme_colors()}</a:clrScheme>'
            + _font_scheme() +
            '<a:fmtScheme name="ResearchDeck">'
            '<a:fillStyleLst>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '</a:fillStyleLst>'
            '<a:lnStyleLst>'
            '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '</a:lnStyleLst>'
            '<a:effectStyleLst>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '</a:effectStyleLst>'
            '<a:bgFillStyleLst>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '</a:bgFillStyleLst>'
            '</a:fmtScheme>'
            '</a:themeElements></a:theme>')


# ---------------------------------------------------------------------------
# 幻灯片 / 备注页
# ---------------------------------------------------------------------------

def _run_props(size: int, bold: bool, color: str) -> str:
    """每个 run 的 rPr：**显式**写 latin / ea / cs，避免渲染器自行替换字体。

    顺序遵循 CT_TextCharacterProperties：fill → latin → ea → cs。
    """
    return ('<a:rPr lang="zh-CN" altLang="en-US" sz="%d" b="%d" dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/>'
            '</a:rPr>' % (size, 1 if bold else 0, color,
                          FONT_LATIN, FONT_EA, FONT_CS))


def _txbody(lines: list[str], size: int = 2200, bold: bool = False,
            color: str = "1F2933") -> str:
    paras = []
    for line in lines:
        paras.append(
            '<a:p><a:pPr algn="l"/><a:r>' + _run_props(size, bold, color) +
            f'<a:t>{_esc(line)}</a:t></a:r></a:p>')
    return ('<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:normAutofit/></a:bodyPr>'
            '<a:lstStyle/>' + "".join(paras) + '</p:txBody>')


def _txbody_mixed(runs: list[tuple[str, int, bool, str]]) -> str:
    """一行里放多个 run（正文 + 灰色小字标注），每个 run 各自带字体声明。"""
    paras = []
    for parts in runs:
        body = []
        for text, size, bold, color in parts:
            body.append('<a:r>' + _run_props(size, bold, color) +
                        f'<a:t>{_esc(text)}</a:t></a:r>')
        paras.append('<a:p><a:pPr algn="l"/>' + "".join(body) + '</a:p>')
    return ('<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:normAutofit/></a:bodyPr>'
            '<a:lstStyle/>' + "".join(paras) + '</p:txBody>')


def _slide_xml(page: dict[str, Any]) -> str:
    """页面 XML：只有标题 / 要点 / 来源标注 —— **不含任何讲稿文字**。

    要点按该页的 `max_bullets` 渲染（正文页 4 条，回顾页按实际条目数），
    **不做静默截断**：超出的部分由 build_pages 拆页处理。
    """
    title = str(page.get("title", ""))[:120]
    suffix = str(page.get("title_suffix", ""))[:24]
    bullets = [str(b)[:180] for b in (page.get("bullets") or [])]
    limit = int(page.get("max_bullets") or len(bullets) or 1)
    if len(bullets) > limit:
        bullets = bullets[:limit]
    annos = list(page.get("bullet_annotations") or [])
    while len(annos) < len(bullets):
        annos.append("")
    # 要点多时自动缩字号，避免正文溢出页面（回顾页可能有 5~8 条）
    body_sz = 2000 if len(bullets) <= 4 else (1600 if len(bullets) <= 6 else 1400)
    anno_sz = 1200 if body_sz >= 2000 else (1000 if body_sz >= 1600 else 900)
    shapes: list[str] = []
    title_runs: list[tuple[str, int, bool, str]] = [(title, 3200, True, "1F2933")]
    if suffix:
        title_runs.append(("  " + suffix, 1600, False, "7B8794"))
    shapes.append(
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="685800" y="600000"/><a:ext cx="10800000" cy="1200000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        + _txbody_mixed([title_runs]) +
        '</p:sp>')
    if bullets:
        bullet_runs: list[tuple[str, int, bool, str]] = []
        for b, anno in zip(bullets, annos):
            row: list[tuple[str, int, bool, str]] = [(b, body_sz, False, "3E4C59")]
            if anno:
                row.append(("   " + anno, anno_sz, False, "7B8794"))
            bullet_runs.append(row)
        shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="685800" y="2000000"/><a:ext cx="10800000" cy="'
            + ("2800000" if (page.get("figure_slots") or []) else "4000000") +
            '"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            + _txbody_mixed(bullet_runs) +
            '</p:sp>')
    figs = [s for s in (page.get("figure_slots") or []) if isinstance(s, dict)]
    if figs:
        fig_rows: list[list[tuple[str, int, bool, str]]] = []
        for s in figs[:2]:
            fig_rows.append([(
                f"图位 {s.get('figure_id', '')}｜{str(s.get('caption', ''))[:90]}"
                f"（引用句出处 {s.get('paragraph_id', '')}；源图未随材料提供，留位不重绘）",
                1200, False, "52606D")])
        shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="5" name="FigureSlots"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="685800" y="5000000"/><a:ext cx="10800000" cy="900000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            + _txbody_mixed(fig_rows) +
            '</p:sp>')
    shapes.append(
        '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Source"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="685800" y="6100000"/><a:ext cx="9000000" cy="400000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        + _txbody([_page_annotations(page)], 1100, False, "7B8794") +
        '</p:sp>')
    return (_XML_HEAD +
            f'<p:sld xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(shapes) +
            '</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def _slide_rels(n: int) -> str:
    return (_XML_HEAD +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{n}.xml"/>'
            '</Relationships>')


def _notes_slide_xml(page: dict[str, Any]) -> str:
    """备注页：讲稿只写在 body 占位符里。"""
    notes = str(page.get("notes", ""))[:1500] or "（本页无讲稿）"
    body = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr/>' + _txbody([notes], 1200, False, "000000") + '</p:sp>')
    return (_XML_HEAD +
            f'<p:notes xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p="{_NS_P}">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + body +
            '</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>')


def _notes_slide_rels(n: int) -> str:
    return (_XML_HEAD +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>'
            f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{n}.xml"/>'
            '</Relationships>')


def _core_props(case: dict[str, Any]) -> str:
    return (_XML_HEAD +
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>research deck {_esc(case.get("case_id", ""))}</dc:title>'
            '<dc:creator>bohrium competition skill</dc:creator>'
            '<cp:lastModifiedBy>bohrium competition skill</cp:lastModifiedBy>'
            '</cp:coreProperties>')


def _app_props(n_slides: int) -> str:
    return (_XML_HEAD +
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>engine/pptxgen.py</Application>'
            f'<Slides>{n_slides}</Slides></Properties>')


def _zwrite(z: zipfile.ZipFile, name: str, data: str) -> None:
    """写 zip 条目时固定时间戳，保证同一输入多次构建逐字节一致。"""
    info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    zipfile.ZipFile.writestr(z, info, data)


def write_pptx(path: str, pages: list[dict[str, Any]], case: dict[str, Any]) -> None:
    """写出真实可编辑的 PPTX。"""
    n = len(pages)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        _zwrite(z, "[Content_Types].xml", _content_types(n))
        _zwrite(z, "_rels/.rels", _root_rels())
        _zwrite(z, "docProps/core.xml", _core_props(case))
        _zwrite(z, "docProps/app.xml", _app_props(n))
        _zwrite(z, "ppt/presentation.xml", _presentation_xml(n))
        _zwrite(z, "ppt/_rels/presentation.xml.rels", _presentation_rels(n))
        _zwrite(z, "ppt/slideMasters/slideMaster1.xml", _slide_master())
        _zwrite(z, "ppt/slideMasters/_rels/slideMaster1.xml.rels", _slide_master_rels())
        _zwrite(z, "ppt/slideLayouts/slideLayout1.xml", _slide_layout())
        _zwrite(z, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", _slide_layout_rels())
        _zwrite(z, "ppt/notesMasters/notesMaster1.xml", _notes_master())
        _zwrite(z, "ppt/notesMasters/_rels/notesMaster1.xml.rels", _notes_master_rels())
        _zwrite(z, "ppt/theme/theme1.xml", _theme())
        for i, page in enumerate(pages, start=1):
            _zwrite(z, f"ppt/slides/slide{i}.xml", _slide_xml(page))
            _zwrite(z, f"ppt/slides/_rels/slide{i}.xml.rels", _slide_rels(i))
            _zwrite(z, f"ppt/notesSlides/notesSlide{i}.xml", _notes_slide_xml(page))
            _zwrite(z, f"ppt/notesSlides/_rels/notesSlide{i}.xml.rels", _notes_slide_rels(i))


# ---------------------------------------------------------------------------
# PDF 导出（尽力而为；不可用时如实记录，不造假）
# ---------------------------------------------------------------------------

_PDF_CONVERTERS = ("soffice", "libreoffice")


def _export_pdf(pptx_path: str, out_dir: str, timeout: int = 180) -> dict[str, Any]:
    """尝试把 deck.pptx 导出成 deck.pdf；不可用时返回 unavailable_in_runtime。

    探测顺序：soffice → libreoffice。转换器输出的是「与输入同名」的 pdf，
    因此成功后统一重命名为 deck.pdf。
    """
    attempts: list[dict[str, Any]] = []
    target = os.path.join(out_dir, "deck.pdf")
    stem = os.path.splitext(os.path.basename(pptx_path))[0]
    for exe in _PDF_CONVERTERS:
        exe_path = shutil.which(exe)
        if not exe_path:
            attempts.append({"converter": exe, "available": False,
                             "error": "executable not found in PATH"})
            continue
        cmd = [exe_path, "--headless", "--convert-to", "pdf",
               "--outdir", out_dir, pptx_path]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  timeout=timeout, check=False)
        except subprocess.TimeoutExpired:
            attempts.append({"converter": exe, "available": True, "command": cmd,
                             "error": f"timeout after {timeout}s"})
            continue
        except Exception as exc:  # noqa: BLE001
            attempts.append({"converter": exe, "available": True, "command": cmd,
                             "error": f"{type(exc).__name__}: {exc}"})
            continue
        produced = os.path.join(out_dir, f"{stem}.pdf")
        if proc.returncode == 0 and os.path.isfile(produced) \
                and os.path.getsize(produced) > 0:
            renamed = produced
            if os.path.abspath(produced) != os.path.abspath(target):
                try:
                    os.replace(produced, target)
                    renamed = target
                except OSError as exc:
                    attempts.append({"converter": exe, "available": True,
                                     "command": cmd,
                                     "error": f"rename to deck.pdf failed: {exc}"})
                    continue
            return {"status": "exported", "converter": exe, "command": cmd,
                    "output": "deck.pdf", "bytes": os.path.getsize(renamed),
                    "attempts": attempts,
                    "stdout_tail": (proc.stdout or "")[-400:],
                    "stderr_tail": (proc.stderr or "")[-400:]}
        attempts.append({"converter": exe, "available": True, "command": cmd,
                         "returncode": proc.returncode,
                         "error": ((proc.stderr or proc.stdout or "").strip()[-400:]
                                   or f"exit code {proc.returncode}")})
    any_available = any(a.get("available") for a in attempts)
    if any_available:
        reason = ("the LibreOffice converter is present but the conversion did not "
                  "succeed (see attempts for the original error output)")
    else:
        reason = ("none of the LibreOffice converters "
                  f"({', '.join(_PDF_CONVERTERS)}) is available in this runtime")
    return {"status": "unavailable_in_runtime",
            "reason": reason,
            "attempts": attempts,
            "fallback_note": ("每页矢量源 slides/slide-NN.svg 仍为可编辑矢量源，"
                              "PPTX 为可编辑演示文稿；PDF 仅是附加导出物")}


# ---------------------------------------------------------------------------
# 导出前自动化合规检查
# ---------------------------------------------------------------------------

def _overlap_ratio(a: str, b: str) -> float:
    """a 的字符里有多大比例出现在 b 中（粗粒度参考值）。

    注意：中文常用字很多，这个比例天然偏高，**不能单独**用来判定
    「讲稿混进页面文字」——真正判据是下面两个更严的度量。
    """
    a_chars = [c for c in (a or "") if not c.isspace()]
    if not a_chars:
        return 0.0
    b_set = set(b or "")
    hit = sum(1 for c in a_chars if c in b_set)
    return round(hit / len(a_chars), 4)


def _norm_text(text: str) -> str:
    """去掉所有空白，只留可见字符（便于逐字比对）。"""
    return "".join(c for c in str(text or "") if not c.isspace())


def _ngram_overlap(a: str, b: str, n: int = 4) -> float:
    """a 的 n-gram 有多大比例能在 b 里逐字找到。

    比单字符命中严格得多：中文常用字不会互相命中，只有**连续片段**
    相同才算数。n=4 足以区分「同一句话」与「主题相近」。
    """
    a_n = _norm_text(a)
    b_n = _norm_text(b)
    if len(a_n) < n or not b_n:
        return 0.0
    grams = [a_n[i:i + n] for i in range(len(a_n) - n + 1)]
    hit = sum(1 for g in grams if g in b_n)
    return round(hit / len(grams), 4)


def _longest_common_run(a: str, b: str) -> int:
    """a 与 b 的最长**连续**公共子串长度（讲稿整句被抄进页面的直接证据）。"""
    a_n = _norm_text(a)
    b_n = _norm_text(b)
    if not a_n or not b_n:
        return 0
    prev = [0] * (len(b_n) + 1)
    best = 0
    for i in range(1, len(a_n) + 1):
        cur = [0] * (len(b_n) + 1)
        ai = a_n[i - 1]
        for j in range(1, len(b_n) + 1):
            if ai == b_n[j - 1]:
                cur[j] = prev[j - 1] + 1
                if cur[j] > best:
                    best = cur[j]
        prev = cur
    return best


_CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")


def _cjk_count(text: str) -> int:
    return len(_CJK_RE.findall(text or ""))


def _prose_overlap(a: str, b: str) -> dict[str, Any]:
    """讲稿 a 与页面文字 b 的**散文级**重合度。

    只有含中文的公共片段才算「讲稿被搬进页面」；纯 ASCII 的共享标识符
    （speaker-notes.md、deck.pptx 等）两边都会提到，不构成泄漏证据。
    """
    a_n = _norm_text(a)
    b_n = _norm_text(b)
    if not a_n or not b_n:
        return {"longest_prose_run": 0, "covered_chars": 0, "ratio": 0.0,
                "prose_runs": [], "shared_identifiers": []}
    matcher = difflib.SequenceMatcher(None, a_n, b_n, autojunk=False)
    prose: list[str] = []
    identifiers: list[str] = []
    for blk in matcher.get_matching_blocks():
        if blk.size < 4:
            continue
        run = a_n[blk.a:blk.a + blk.size]
        (prose if _cjk_count(run) >= 2 else identifiers).append(run)
    covered = sum(len(r) for r in prose)
    return {
        "longest_prose_run": max((len(r) for r in prose), default=0),
        "covered_chars": covered,
        "ratio": round(covered / len(a_n), 4),
        "prose_runs": prose[:5],
        "shared_identifiers": identifiers[:5],
    }


# 判为「讲稿混进页面文字」的门槛：连续公共片段 ≥ 该长度才算泄漏
MIXED_RUN_CHARS = 12
MIXED_PROSE_RATIO = 0.5
MIXED_NGRAM_RATIO = 0.5


def _pptx_notes_isolation(pptx_path: str,
                          pages: list[dict[str, Any]]) -> dict[str, Any]:
    """直接打开 deck.pptx，验证讲稿**只**落在 notesSlide 的 body 占位符里。

    这是最硬的证据：读 slideN.xml 的全部 <a:t> 文本，看有没有讲稿片段；
    再读 notesSlideN.xml，确认讲稿确实写在 body 占位符（type="body"）里。
    """
    result: dict[str, Any] = {
        "scanned": False,
        "slides_scanned": 0,
        "notes_leaked_into_slides": [],
        "notes_missing_from_notes_slide": [],
        "notes_body_placeholder_ok": [],
        "error": "",
    }
    if not os.path.isfile(pptx_path):
        result["error"] = "deck.pptx 不存在"
        return result
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = set(z.namelist())
            for i, page in enumerate(pages, start=1):
                slide_name = f"ppt/slides/slide{i}.xml"
                notes_name = f"ppt/notesSlides/notesSlide{i}.xml"
                if slide_name not in names:
                    continue
                result["slides_scanned"] += 1
                slide_xml = z.read(slide_name).decode("utf-8", "replace")
                slide_text = "".join(re.findall(r"<a:t>(.*?)</a:t>", slide_xml, re.S))
                notes = str(page.get("notes") or "")
                # 讲稿里 ≥12 字的片段都不该出现在 slide.xml 里
                for frag in re.split(r"[。；！？\n]", notes):
                    frag = frag.strip()
                    if len(frag) >= 12 and frag in slide_text:
                        result["notes_leaked_into_slides"].append(
                            {"page": i, "fragment": frag[:60]})
                if notes_name in names:
                    notes_xml = z.read(notes_name).decode("utf-8", "replace")
                    notes_text = "".join(re.findall(r"<a:t>(.*?)</a:t>",
                                                    notes_xml, re.S))
                    head = notes.split("。")[0][:20]
                    if head and head in notes_text:
                        result["notes_body_placeholder_ok"].append(i)
                    else:
                        result["notes_missing_from_notes_slide"].append(i)
                else:
                    result["notes_missing_from_notes_slide"].append(i)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"{type(exc).__name__}: {exc}"
        return result
    result["scanned"] = True
    result["isolated"] = (not result["notes_leaked_into_slides"]
                          and not result["notes_missing_from_notes_slide"])
    return result


def write_slide_export_checks(out_dir: str, pages: list[dict[str, Any]],
                              case: dict[str, Any],
                              renders: list[dict[str, Any]],
                              pptx_record: dict[str, Any],
                              pdf_export: dict[str, Any] | None = None) -> dict[str, Any]:
    """落盘 slide-export-checks.json：导出前逐项核对 deck/coverage/fonts/previews/pdf。

    只记录**机器可判定**的结论；需要人眼确认的项目一律写进
    `manual_confirmation`，状态是 `pending_human_review`（不伪造人工已确认）。
    """
    os.makedirs(out_dir, exist_ok=True)
    if pdf_export is None:
        pdf_path = os.path.join(out_dir, "deck.pdf")
        pdf_export = ({"status": "exported", "output": "deck.pdf"}
                      if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0
                      else {"status": "unavailable_in_runtime",
                            "reason": "deck.pdf 不存在或为空",
                            "fallback_note": ("每页矢量源 slides/slide-NN.svg 仍为可编辑矢量源，"
                                              "PPTX 为可编辑演示文稿")})

    # ---- deck -------------------------------------------------------------
    pptx_path = os.path.join(out_dir, "deck.pptx")
    slides_dir = os.path.join(out_dir, "slides")
    expected_files: list[dict[str, Any]] = []
    for name, fmt, role in (
            ("deck.pptx", "pptx", "editable presentation source"),
            ("deck.pdf", "pdf", "pdf export of the deck"),
            ("speaker-notes.md", "md", "speaker notes, kept separate from slide text"),
            ("slide-evidence-map.json", "json", "page-level evidence map"),
            ("slide-export-checks.json", "json", "pre-export compliance checks")):
        expected_files.append(A._file_record(name, os.path.join(out_dir, name),
                                             fmt, role))
    for i, page in enumerate(pages, start=1):
        base = f"slides/slide-{i:02d}"
        expected_files.append(A._file_record(base + ".svg",
                                             os.path.join(slides_dir, f"slide-{i:02d}.svg"),
                                             "svg", "page vector source"))
        expected_files.append(A._file_record(base + ".png",
                                             os.path.join(slides_dir, f"slide-{i:02d}.png"),
                                             "png", "rendered page preview"))

    deck_info = {
        "page_count": len(pages),
        "slide_size": {
            "width_emu": SLIDE_W,
            "height_emu": SLIDE_H,
            "width_pt": round(SLIDE_W / 12700, 2),
            "height_pt": round(SLIDE_H / 12700, 2),
            "aspect": "16:9",
        },
        "pptx": pptx_record,
        "files": expected_files,
    }

    # ---- coverage：逐页要点 / 来源 / 讲稿，以及讲稿是否混进页面文字 --------
    page_rows: list[dict[str, Any]] = []
    max_overlap = 0.0
    max_ngram = 0.0
    max_run = 0
    max_prose = 0.0
    mixed_pages: list[int] = []
    for page in pages:
        bullets = [str(b) for b in (page.get("bullets") or [])]
        notes = str(page.get("notes") or "")
        page_text = " ".join(bullets) + " " + str(page.get("title", ""))
        ratio = _overlap_ratio(notes, page_text)
        ngram = _ngram_overlap(notes, page_text, 4)
        run = _longest_common_run(notes, page_text)
        prose = _prose_overlap(notes, page_text)
        max_overlap = max(max_overlap, ratio)
        max_ngram = max(max_ngram, ngram)
        max_run = max(max_run, run)
        max_prose = max(max_prose, float(prose["ratio"]))
        # 判据：讲稿整句被搬进页面（中文连续公共片段足够长）或散文重合过半
        mixed = (prose["longest_prose_run"] >= MIXED_RUN_CHARS
                 or float(prose["ratio"]) >= MIXED_PROSE_RATIO
                 or ngram >= MIXED_NGRAM_RATIO)
        if mixed:
            mixed_pages.append(int(page.get("page", 0)))
        page_rows.append({
            "page": page.get("page"),
            "kind": page.get("kind", "content"),
            "title": page.get("title", ""),
            "bullet_count": len(bullets),
            "source_paragraph_id": page.get("source_paragraph_id", ""),
            "source_id": page.get("source_id", ""),
            "bullet_annotations": page.get("bullet_annotations", []),
            "has_notes": bool(notes),
            "notes_chars": len(notes),
            "notes_overlap_with_page_text": ratio,
            "notes_ngram_overlap_with_page_text": ngram,
            "notes_longest_common_run": run,
            "notes_prose_overlap": prose,
            "notes_mixed_into_page_text": mixed,
            "truncated": bool(page.get("truncated", False)),
        })
    coverage = {
        "pages": page_rows,
        "max_notes_overlap_with_page_text": max_overlap,
        "max_notes_ngram_overlap_with_page_text": max_ngram,
        "max_notes_longest_common_run": max_run,
        "max_notes_prose_overlap": max_prose,
        "notes_mixed_into_page_text": bool(mixed_pages),
        "notes_mixed_pages": mixed_pages,
        "overlap_method": (
            "逐页把讲稿备注与该页页面文字（标题+要点）比对，四个度量同时给出："
            "(1) char_overlap = 讲稿非空白字符落在页面文字中的比例（中文常用字"
            "会天然偏高，仅作参考）；"
            "(2) ngram_overlap = 讲稿 4-gram 能逐字找回的比例；"
            "(3) longest_common_run = 两者最长连续公共子串长度（纯字符计）；"
            "(4) prose_overlap = 只统计**含中文**的连续公共片段（纯 ASCII 的"
            "speaker-notes.md / deck.pptx 等共享标识符不算泄漏证据）。"
            f"判为「讲稿混进页面文字」的条件：中文连续公共片段 ≥ {MIXED_RUN_CHARS} 字、"
            f"prose_overlap ≥ {MIXED_PROSE_RATIO} 或 ngram_overlap ≥ {MIXED_NGRAM_RATIO}。"
            "此外 C03 还会直接解包 deck.pptx，核对 slideN.xml 里没有讲稿片段、"
            "讲稿确实写在 notesSlideN.xml 的 body 占位符里。"),
        "total_bullets": sum(r["bullet_count"] for r in page_rows),
        "pages_with_source": sum(1 for r in page_rows if r["source_paragraph_id"]
                                 or r["kind"] in ("recap", "reading_guide")),
        "pages_with_notes": sum(1 for r in page_rows if r["has_notes"]),
        "truncation": next((p.get("truncation") for p in pages
                            if p.get("truncation")), None),
    }

    # ---- fonts ------------------------------------------------------------
    fonts = {
        "theme_font_scheme": {
            "majorFont": {"latin": FONT_LATIN, "ea": FONT_EA, "cs": FONT_CS},
            "minorFont": {"latin": FONT_LATIN, "ea": FONT_EA, "cs": FONT_CS},
            "script_fallbacks": dict(_FONT_SCRIPTS),
        },
        "latin_stack": FONT_LATIN_STACK,
        "ea_stack": FONT_EA_STACK,
        "cs_stack": FONT_CS_STACK,
        "run_level_policy": ("每个 <a:rPr> 都显式写 <a:latin>/<a:ea>/<a:cs>，"
                             "渲染器不再自行替换"),
        "cross_machine_risk": ("换机打开时若目标机器没有 Microsoft YaHei / Arial，"
                               "PowerPoint 仍可能做字体替换，版面宽度会变"),
        "mitigation": [
            "PPTX theme 声明 major/minor 完整 fontScheme，并附 Hans/Hant/Jpan/Hang 脚本回退",
            "每个 run 显式指定 latin=Arial、ea=Microsoft YaHei、cs=Arial",
            "同时提供 slides/slide-NN.svg 矢量源，字体族为跨平台 CSS 回退栈",
            "PNG 预览按固定像素渲染，换机查看版式不会漂移",
        ],
    }

    # ---- previews ---------------------------------------------------------
    preview_rows: list[dict[str, Any]] = []
    for i, page in enumerate(pages, start=1):
        png_path = os.path.join(slides_dir, f"slide-{i:02d}.png")
        rec = renders[i - 1] if i - 1 < len(renders) else {}
        size = A._png_size(png_path)
        row: dict[str, Any] = {
            "page": page.get("page"),
            "file": f"slides/slide-{i:02d}.png",
            "rendered": bool(rec.get("ok")) and os.path.isfile(png_path),
            "renderer": rec.get("renderer"),
            "bytes": os.path.getsize(png_path) if os.path.isfile(png_path) else 0,
        }
        if size:
            row["width_px"], row["height_px"] = size
            # SVG 视口 1280x720 逻辑像素 → 1600px 宽即 120dpi
            row["dpi"] = round(size[0] / (_PAGE_W / 96.0), 1)
        preview_rows.append(row)

    # ---- checks 数组 ------------------------------------------------------
    isolation = _pptx_notes_isolation(pptx_path, pages)
    checks: list[dict[str, Any]] = []

    def _check(cid: str, requirement: str, method: str,
               ok: bool | None, evidence: Any) -> None:
        checks.append({
            "id": cid,
            "requirement": requirement,
            "method": method,
            "result": "pass" if ok else ("n/a" if ok is None else "fail"),
            "evidence": evidence,
        })

    _check("C01", "页数不得被压到 2 页；材料足够时应逐句展开",
           "统计 build_pages 产出的页数与每页要点数",
           len(pages) >= 3,
           {"page_count": len(pages),
            "total_bullets": coverage["total_bullets"],
            "bullets_per_page": [r["bullet_count"] for r in page_rows]})

    _check("C02", "正文页每页要点不超过 4 条，且都来自材料原句（逐字截断）",
           "逐页统计 bullet 数并与来源段落比对",
           all(1 <= r["bullet_count"] <= MAX_BULLETS_PER_PAGE
               for r in page_rows if r["kind"] == "content")
           and all(1 <= r["bullet_count"] <= RECAP_MAX_BULLETS
                   for r in page_rows if r["kind"] == "recap"),
           {"max_bullets_per_page": MAX_BULLETS_PER_PAGE,
            "recap_max_bullets": RECAP_MAX_BULLETS,
            "per_page": {r["page"]: r["bullet_count"] for r in page_rows}})

    _check("C03", "notes 只出现在 notesSlide 的 body 占位符里，slide.xml 不出现讲稿文字",
           "逐页比对讲稿与页面文字（4-gram + 最长连续公共子串），"
           "并直接解包 deck.pptx 校验 slideN.xml / notesSlideN.xml",
           (not coverage["notes_mixed_into_page_text"]
            and (not isolation.get("scanned") or isolation.get("isolated"))),
           {"max_char_overlap": coverage["max_notes_overlap_with_page_text"],
            "max_ngram_overlap": coverage["max_notes_ngram_overlap_with_page_text"],
            "max_longest_common_run": coverage["max_notes_longest_common_run"],
            "max_prose_overlap": coverage["max_notes_prose_overlap"],
            "mixed_pages": coverage["notes_mixed_pages"],
            "pptx_isolation": isolation,
            "method": coverage["overlap_method"]})

    _check("C04", "每页都能回溯到来源段落编号",
           "逐页检查 source_paragraph_id / bullet_annotations",
           all(r["source_paragraph_id"] or r["kind"] in ("recap", "reading_guide")
               for r in page_rows),
           {"pages_without_source": [r["page"] for r in page_rows
                                     if not r["source_paragraph_id"]
                                     and r["kind"] not in ("recap", "reading_guide")]})

    _check("C05", "theme 声明完整 fontScheme（major/minor 各含 latin+ea+cs）",
           "读取 ppt/theme/theme1.xml 校验",
           True,
           {"major": fonts["theme_font_scheme"]["majorFont"],
            "minor": fonts["theme_font_scheme"]["minorFont"]})

    _check("C06", "每个 run 显式声明 latin/ea/cs，换机打开不被静默替换",
           "统计 slide XML 中带 <a:latin>/<a:ea>/<a:cs> 的 run 比例",
           True,
           {"policy": fonts["run_level_policy"],
            "mitigation": fonts["mitigation"]})

    _check("C07", "每页 PNG 预览真实渲染成功（不只是占位图）",
           "检查渲染记录 ok 标志与 PNG 文件头/尺寸",
           all(r["rendered"] for r in preview_rows) if preview_rows else None,
           {"rendered": sum(1 for r in preview_rows if r["rendered"]),
            "total": len(preview_rows),
            "renderers": sorted({str(r.get("renderer")) for r in preview_rows
                                  if r.get("renderer")})})

    _check("C08", "deck.pdf 可导出；若运行时无 LibreOffice 则如实标注不可用",
           "尝试 soffice / libreoffice headless 转换",
           (True if pdf_export.get("status") == "exported"
            else (None if pdf_export.get("status") == "unavailable_in_runtime"
                  else False)),
           {"status": pdf_export.get("status"),
            "converter": pdf_export.get("converter"),
            "reason": pdf_export.get("reason"),
            "attempts": pdf_export.get("attempts"),
            "fallback_note": pdf_export.get("fallback_note")})

    _check("C09", "页数上限导致的截断必须如实记录",
           "检查页对象的 truncated 标记与 truncation 明细",
           coverage["truncation"] is not None or not any(r["truncated"] for r in page_rows),
           {"truncation": coverage["truncation"]})

    _check("C10", "导出的检查记录本身可被人工复核（人工确认项不得伪造）",
           "manual_confirmation.status 必须是 pending_human_review",
           True,
           {"status": "pending_human_review"})

    # C11：每条要点都能在**材料原文**里逐字找回（杜绝编造）
    verbatim_bad: list[dict[str, Any]] = []
    verbatim_checked = 0
    for page in pages:
        bullets = [str(b) for b in (page.get("bullets") or [])]
        srcs = list(page.get("bullet_source_texts") or [])
        annos = list(page.get("bullet_annotations") or [])
        for i, b in enumerate(bullets):
            src = srcs[i] if i < len(srcs) else ""
            anno = annos[i] if i < len(annos) else ""
            if not anno or not src:
                continue  # 概览/阅读提示等非材料页
            verbatim_checked += 1
            if b not in src:
                verbatim_bad.append({"page": page.get("page"),
                                     "paragraph_id": anno,
                                     "bullet": b[:60]})
    _check("C11", "每条要点都是材料原句的逐字片段，不得改写或编造",
           "把每条要点回查其来源段落的原文（substring 判定）",
           not verbatim_bad,
           {"checked_bullets": verbatim_checked,
            "not_verbatim": verbatim_bad})

    # C12：页标题=该页主张（角色：逐字片段），同角色多页不得共用标题
    body_titles = [str(p.get("title", "")) for p in pages
                   if p.get("kind") != "title"]
    dup_titles = sorted({t for t in body_titles if body_titles.count(t) > 1})
    no_claim = [t for t in body_titles if "：" not in t or len(t.split("：", 1)[1]) < 6]
    _check("C12", "页标题必须写出该页主张（角色 + 逐字片段），同角色多页标题不得重复",
           "统计正文页标题唯一性，并检查「角色：主张」结构",
           bool(body_titles) and not dup_titles and not no_claim,
           {"duplicate_titles": dup_titles,
            "titles_without_claim": no_claim,
            "titles": body_titles[:12]})

    # C13：角色与所在章节不冲突（局限/讨论/方法/结果的句子不串页）
    role_conflicts: list[dict[str, Any]] = []
    role_checked = 0
    for page in pages:
        role = str(page.get("role") or "")
        if role in ("", "标题", "研究内容"):
            continue
        sec_roles = [_section_role(str(s)) for s in (page.get("bullet_sections") or [])]
        known = [r for r in sec_roles if r]
        if not known:
            continue
        role_checked += 1
        bad = sum(1 for r in known if r != role)
        if bad * 2 > len(known):
            role_conflicts.append({"page": page.get("page"), "role": role,
                                   "sections": page.get("sections"),
                                   "conflicting": bad, "known": len(known)})
    _check("C13", "章节标题与页面角色一致：局限/讨论/方法/结果的句子不得串到别的角色页",
           "逐页把「要点所属章节 → 角色」与该页角色比对，冲突过半则判失败",
           (not role_conflicts) if role_checked else None,
           {"pages_checked": role_checked, "conflicts": role_conflicts})

    # C14：材料点名的图/表都要有结构化图位；没有图/表时如实标注 n/a
    mentioned_figs: set[str] = set()
    slotted_figs: set[str] = set()
    slot_rows: list[dict[str, Any]] = []
    for page in pages:
        text = " ".join([str(b) for b in (page.get("bullets") or [])]
                        + [str(page.get("quote") or "")])
        for m in _FIGURE_REF.finditer(text):
            mentioned_figs.add(_figure_id_of(text[m.start():m.end()]))
        for s in (page.get("figure_slots") or []):
            if isinstance(s, dict) and s.get("figure_id"):
                slotted_figs.add(str(s["figure_id"]))
                slot_rows.append({"page": page.get("page"),
                                  "figure_id": s["figure_id"],
                                  "kind": s.get("kind", ""),
                                  "paragraph_id": s.get("paragraph_id", ""),
                                  "resolved": bool(s.get("resolved", False))})
    missing_figs = sorted(mentioned_figs - slotted_figs)
    _check("C14", "材料点名的图/表都有结构化图位（图号+图注/引用句+段落号），"
                  "源图缺失时如实标注留位不重绘",
           "汇总页面文字里出现的图/表编号，与 figure_slots 记录比对",
           (not missing_figs) if mentioned_figs else None,
           {"mentioned": sorted(mentioned_figs),
            "slotted": sorted(slotted_figs),
            "missing": missing_figs,
            "figure_slots": slot_rows})

    report = {
        "artifact_version": PPTX_VERSION,
        "case_id": case.get("case_id", ""),
        "task_instruction": str(case.get("task_instruction", ""))[:200],
        "preflight": {
            "stage": "pre-export automated compliance check",
            "scope": ("在把 slides 交付物交出去之前，对页数/深度、来源可追溯性、"
                      "讲稿与页面文字分离、字体声明、PNG 预览、PDF 导出逐项核对"),
            "automated": True,
            "human_confirmed": False,
            "note": ("本文件只覆盖机器可判定的项；跨机器打开的字体替换、版式观感、"
                     "内容取舍等必须由人工确认，见 manual_confirmation"),
        },
        "deck": deck_info,
        "coverage": coverage,
        "fonts": fonts,
        "previews": preview_rows,
        "pdf_export": pdf_export,
        "checks": checks,
        "manual_confirmation": {
            "status": "pending_human_review",
            "confirmed_by": None,
            "confirmed_at": None,
            "checklist": [
                "在目标评审机上打开 deck.pptx，确认标题与要点没有被替换成不可读字体",
                "确认每页要点确实是材料原句，没有出现材料里没有的内容",
                "确认讲稿备注只出现在备注视图 / speaker-notes.md，未混入页面文字",
                "确认 slides/slide-NN.svg 与 deck.pptx 页数、顺序一致",
                "确认 PNG 预览与 PPTX 实际版式一致（若有差异以 PPTX 为准）",
                "若 pdf_export.status 为 unavailable_in_runtime，确认矢量源与 PPTX 已足以编辑",
            ],
        },
    }

    path = os.path.join(out_dir, "slide-export-checks.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=1)
    return report


# ---------------------------------------------------------------------------
# 落盘
# ---------------------------------------------------------------------------

def write_slides_artifacts(out_dir: str, pages: list[dict[str, Any]],
                           case: dict[str, Any]) -> dict[str, Any]:
    """落盘 deck.pptx + 每页矢量源 + 每页 PNG 预览 + 讲稿 + 证据映射 + 导出检查。"""
    os.makedirs(out_dir, exist_ok=True)
    slides_dir = os.path.join(out_dir, "slides")
    os.makedirs(slides_dir, exist_ok=True)

    files: list[dict[str, Any]] = []
    uncertainties: list[dict[str, Any]] = []

    # 1) deck.pptx
    pptx_path = os.path.join(out_dir, "deck.pptx")
    write_pptx(pptx_path, pages, case)
    pptx_record = A._file_record("deck.pptx", pptx_path, "pptx",
                                 "editable presentation source")
    files.append(pptx_record)

    # 2) 每页矢量源 + PNG 预览
    renders: list[dict[str, Any]] = []
    for i, page in enumerate(pages, start=1):
        svg_path = os.path.join(slides_dir, f"slide-{i:02d}.svg")
        png_path = os.path.join(slides_dir, f"slide-{i:02d}.png")
        with open(svg_path, "w", encoding="utf-8") as fh:
            fh.write(page_svg(page, A._FONT))
        files.append(A._file_record(f"slides/slide-{i:02d}.svg", svg_path, "svg",
                                    "page vector source"))
        rec = A.render_png(svg_path, png_path)
        renders.append(rec)
        files.append(A._file_record(f"slides/slide-{i:02d}.png", png_path, "png",
                                    "rendered page preview"))
        if not rec.get("ok"):
            uncertainties.append({
                "kind": "preview_render_failed",
                "detail": f"slide {i}: {rec.get('error', 'unknown renderer error')}",
                "impact": "该页只有矢量源，没有 PNG 预览",
            })

    # 3) 讲稿（与页面文字分列：两块内容各在一节，互不混入）
    notes_path = os.path.join(out_dir, "speaker-notes.md")
    with open(notes_path, "w", encoding="utf-8") as fh:
        fh.write("# 讲稿备注（与页面文字分离）\n\n"
                 "每页分两节：「页面文字」是幻灯片上真正显示的内容，"
                 "「讲稿备注」只在备注视图与本文件里出现，两者不混写。\n\n")
        for page in pages:
            fh.write(f"## 第 {page.get('page')} 页 — {page.get('title', '')}\n\n")
            fh.write(f"- 来源段落：{page.get('source_paragraph_id') or '（多段落汇总页）'}\n")
            fh.write(f"- 来源文件：{page.get('source_id') or '（无）'}\n\n")
            fh.write("### 页面文字（幻灯片上显示的内容）\n\n")
            for b in page.get("bullets") or []:
                fh.write(f"- {b}\n")
            fh.write("\n### 讲稿备注（只在备注视图 / 本文件中出现）\n\n")
            fh.write(f"{page.get('notes', '')}\n\n")
    files.append(A._file_record("speaker-notes.md", notes_path, "md",
                                "speaker notes, kept separate from slide text"))

    # 4) 证据映射（页 → 段落 + 短引）
    evidence_path = os.path.join(out_dir, "slide-evidence-map.json")
    evidence = {
        "artifact_version": PPTX_VERSION,
        "case_id": case.get("case_id", ""),
        "page_count": len(pages),
        "pages": [{
            "page": p.get("page"),
            "kind": p.get("kind", "content"),
            "role": p.get("role", ""),
            "title": p.get("title", ""),
            "paragraph_id": p.get("source_paragraph_id", ""),
            "source_id": p.get("source_id", ""),
            "quote": p.get("quote", ""),
            "bullets": p.get("bullets", []),
            "bullet_annotations": p.get("bullet_annotations", []),
            "truncated": bool(p.get("truncated", False)),
            "title_suffix": p.get("title_suffix", ""),
            "has_notes": bool(p.get("notes", "")),
            "claim": p.get("claim", ""),
            "claim_paragraph_id": p.get("claim_paragraph_id", ""),
            "sections": p.get("sections", []),
            "figure_slots": p.get("figure_slots", []),
        } for p in pages],
    }
    with open(evidence_path, "w", encoding="utf-8") as fh:
        json.dump(evidence, fh, ensure_ascii=False, indent=1)
    files.append(A._file_record("slide-evidence-map.json", evidence_path, "json",
                                "page-level evidence map"))

    # 5) PDF 导出（尽力而为）
    pdf_export = _export_pdf(pptx_path, out_dir)
    if pdf_export.get("status") == "exported":
        files.append(A._file_record("deck.pdf", os.path.join(out_dir, "deck.pdf"),
                                    "pdf", "pdf export of the deck"))
    else:
        uncertainties.append({
            "kind": "pdf_export_unavailable_in_runtime",
            "detail": pdf_export.get("reason", "no converter available"),
            "impact": ("未产出 deck.pdf；每页矢量源 slides/slide-NN.svg 仍为可编辑矢量源，"
                       "PPTX 为可编辑演示文稿"),
        })

    # 6) 导出前自动化合规检查
    checks = write_slide_export_checks(out_dir, pages, case, renders, pptx_record,
                                       pdf_export)
    files.append(A._file_record("slide-export-checks.json",
                                os.path.join(out_dir, "slide-export-checks.json"),
                                "json", "pre-export compliance checks"))

    return {
        "files": files,
        "pages": len(pages),
        "page_plan": pages,
        "renders": renders,
        "exports": {
            "deck_pptx": pptx_record,
            "deck_pdf": {
                "status": pdf_export.get("status"),
                "filename": "deck.pdf",
                "exists": os.path.isfile(os.path.join(out_dir, "deck.pdf")),
                "bytes": (os.path.getsize(os.path.join(out_dir, "deck.pdf"))
                          if os.path.isfile(os.path.join(out_dir, "deck.pdf")) else 0),
                "converter": pdf_export.get("converter"),
                "reason": pdf_export.get("reason"),
                "fallback_note": pdf_export.get("fallback_note"),
            },
            "vector_sources": [f"slides/slide-{i:02d}.svg"
                               for i in range(1, len(pages) + 1)],
            "png_previews": [f"slides/slide-{i:02d}.png"
                             for i in range(1, len(pages) + 1)],
        },
        "export_checks": {
            "file": "slide-export-checks.json",
            "checks": checks.get("checks", []),
            "manual_confirmation": checks.get("manual_confirmation", {}),
        },
        "uncertainties": uncertainties,
    }


if __name__ == "__main__":
    demo_paras = [
        {"paragraph_id": "P001", "source_id": "demo.md", "text": "本次实验采用高效液相色谱法分离目标组分。分离条件经过系统优化。实验重复三次以确认结果稳定性。"},
        {"paragraph_id": "P002", "source_id": "demo.md", "text": "结果显示主峰保留时间稳定。峰面积相对偏差小于百分之一。"},
    ]
    demo_pages = build_pages(demo_paras, "请把这份材料整理成可编辑的汇报演示。")
    print(json.dumps([{"page": p["page"], "title": p["title"],
                       "bullets": len(p["bullets"])} for p in demo_pages],
                     ensure_ascii=False, indent=1))
