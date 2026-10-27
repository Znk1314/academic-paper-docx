#!/usr/bin/env python
"""启发式「AI 味」检测：只给段落级证据，不改稿。

规则编号与 `references/anti-ai-academic-zh.md` 一一对应：

    S2.1 空转的意义句        S2.2 数据缺席的拔高     S2.3 假对比与假递进
    S2.4 段落末尾复述收尾    A2.5 起跑式铺垫/元话语   A2.6 连接词铺路
    A2.7 句长均匀            B2.9 名词化与"进行+动词" B2.10 无主被动
    B2.11 四字排比           S2.12 引用装饰/借权威    A3.1 段落首句复读标题
    A3.4 全文零数字

设计取舍（重要）：

* **只报警，不判身份**：分数高不等于 AI 写的，模板化公文、翻译腔同样会高分。
* **词表不是黑名单**：每条命中都要回到上下文；文档 §4 的学术豁免（方法段被动句、
  术语化名词化、必要限定、引用格式）默认不报。
* **数据可以压倒拔高**：同一句里带数字/百分比/对比时，S2.2 不报。
* **引用与引语先剥离**：`[3]`、"…"、「…」、`code` 里的内容不参与匹配。

用法：

    python scripts/ai_tone_check.py --docx out/report.docx
    python scripts/ai_tone_check.py --json examples/demo/content_model.json
    python scripts/ai_tone_check.py --text draft.md --max-score 40   # 超阈值退出码 1
"""

from __future__ import annotations

import argparse
import json
import os
import re
import statistics
import sys
from dataclasses import dataclass, field

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import _console  # noqa: E402

# ---------------------------------------------------------------- 规则表

CJK = r"\u4e00-\u9fa5"
_NUM = r"\d+(?:\.\d+)?"


@dataclass
class Rule:
    id: str
    level: str          # S / A / B
    title: str
    pattern: str
    advice: str
    weight: int
    # True 时：命中句里若含数字/百分比则视为"有数据"，不报（S2.2 用）
    needs_no_number: bool = False
    # True 时：只在该段最后一句命中才算（S2.4 用）
    paragraph_final: bool = False
    # True 时：结构级规则，不参与逐句正则匹配（A2.7 / A3.1 / A3.4）
    structural: bool = False


RULES: list[Rule] = [
    Rule(
        "S2.1", "S", "空转的意义句",
        r"具有(?:重要|深远|重大)?的?(?:理论|现实|实践)?(?:意义|价值|作用)"
        r"|广阔的?(?:应用)?前景|提供了?(?:新的|有益的)?(?:思路|参考|借鉴|方向)"
        r"|奠定了?坚实的基础|具有一定的?指导意义",
        "换成这个结论真正约束了什么（对谁、在什么条件下、限制了什么）",
        6,
    ),
    Rule(
        "S2.2", "S", "数据缺席的拔高",
        rf"(?:显著|大幅|有效|明显|极大)地?(?:提升|提高|改善|增强|降低|减少|缩短)"
        rf"|性能优异|效果良好|取得了?(?:良好|显著|丰硕)的?(?:效果|成果|进展)"
        rf"|{CJK}{{0,6}}表现(?:优异|突出)",
        "补上指标、数值、对比对象、条件中的至少两项；没有数据就写\"尚未量化\"",
        6,
        needs_no_number=True,
    ),
    Rule(
        "S2.3", "S", "假对比与假递进",
        r"不是[^。；\n]{0,24}(?:而是|更是)|不仅[^。；\n]{0,24}更是"
        r"|并非[^。；\n]{0,24}而是|与其说[^。；\n]{0,24}不如说",
        "直接陈述肯定项；只有否定项在纠正读者的真实误解时才保留对比",
        6,
    ),
    Rule(
        "S2.4", "S", "段落末尾复述收尾",
        r"这(?:就)?(?:说明|表明|意味着|正是)|由此可见|可见|这正是"
        r"|(?:充分)?(?:彰显|体现)了",
        "删掉复述；若有推论，写成带条件的新句子",
        6,
        paragraph_final=True,
    ),
    Rule(
        "A2.5", "A", "起跑式铺垫与元话语",
        r"本节(?:将|主要)|下面(?:我们)?(?:来)?(?:看|介绍|讨论|分析)"
        r"|接下来(?:我们)?|众所周知|不难(?:发现|看出)"
        r"|值得(?:注意|一提|强调)的是|需要注意的是|首先需要说明的是"
        r"|随着[^。；\n]{0,16}(?:的)?(?:不断|快速|迅速|蓬勃)?(?:发展|进步|推进|普及)",
        "删掉铺垫直接给内容；范围界定类句子保留",
        4,
    ),
    Rule(
        "A2.6", "A", "连接词铺路（每段都用）",
        r"^(?:因此|然而|此外|同时|首先|其次|再次|最后|另外|总之|综上|与此同时)[，,]",
        "把关系写进句子（因果用动词与条件），保留真正需要的转折",
        4,
    ),
    Rule(
        "B2.9", "B", "名词化与\"进行＋动词\"",
        r"进行(?:了|着)?(?:[\u4e00-\u9fa5]{1,4})(?:分析|研究|测试|优化|处理|评估|讨论|验证|比较|测量|改进)"
        r"|对[^。；\n]{0,12}的(?:优化|分析|实现|设计|改进|评估)",
        "换成直接动词：\"测试了…\"而不是\"对…进行了测试\"",
        2,
    ),
    Rule(
        "B2.10", "B", "无主被动与\"被\"字堆叠",
        r"被(?:认为|视为|广泛|普遍|发现|证明)|被[^。；\n]{0,8}(?:认为|视为)",
        "施事已知就写出来；确属未知才用被动，并保留不确定性",
        2,
    ),
    Rule(
        "B2.11", "B", "四字排比与形容词堆叠",
        rf"(?:[{CJK}]{{2,4}}、){{2,}}[{CJK}]{{2,4}}",
        "换成可测的表述（时间、次数、阈值）",
        2,
    ),
    Rule(
        "S2.12", "S", "引用装饰与借权威",
        r"专家(?:认为|指出|表示)|业内(?:普遍)?认为|大量(?:研究|实践)(?:表明|证明)"
        r"|(?:研究|实践)表明(?!\s*\[)|有(?:研究|学者)(?:指出|认为)",
        "点名、点条件、点结论（\"文献[5]在 20 个节点上测得…\"），否则删掉",
        6,
    ),
    Rule(
        "B2.13", "B", "装饰性标点与格式",
        # 数值范围（60—90 s、图1—图3）里的连接号是规范用法，必须放过
        rf"(?<!\d)[—–](?!\d)|[{CJK}][→←↑↓]|[\U0001F300-\U0001FAFF\u2705\u274C]|\*\*[^*\n]+\*\*",
        "论文正文不用破折号堆叠、箭头装饰、emoji 与加粗强调",
        3,
    ),
    # —— 结构级规则：不参与逐句匹配，由 analyze() 在段落/全篇层面判定 ——
    Rule(
        "A2.7", "A", "句长均匀",
        "", "关键结论压成短句，限定条件并成长句；不要为了变化而硬拆硬合",
        4, structural=True,
    ),
    Rule(
        "A3.1", "A", "段落首句复读标题",
        "", "删掉复述标题的那句，直接给条件、定义或数据",
        3, structural=True,
    ),
    Rule(
        "A3.4", "A", "全文零数字",
        "", "至少给一个具体数字、区间或条件（样本量、阈值、指标值）",
        3, structural=True,
    ),
]

CONNECTIVES = r"因此|然而|此外|同时|首先|其次|最后|综上|总之|另外|需要指出|由此可见"
CITATION = re.compile(r"\[\s*\d+(?:\s*[-,，]\s*\d+)*\s*\]")
QUOTED = re.compile(r"「[^」]{0,80}」|“[^”]{0,80}”|\"[^\"]{0,80}\"|‘[^’]{0,80}’|`[^`]{0,80}`")
SENT_SPLIT = re.compile(r"[。！？；!?;]+")
CAPTION = re.compile(r"^(?:图|表)\s*\d+|^注：")
REF_HEADING = re.compile(r"^(?:参考文献|References)$", re.I)
HEADING_STYLE = re.compile(r"^(?:Heading|标题)\s*\d", re.I)


@dataclass
class Hit:
    rule_id: str
    level: str
    title: str
    para_no: int
    excerpt: str
    advice: str


@dataclass
class Report:
    source: str
    chars: int = 0
    paragraphs: int = 0
    sentences: int = 0
    hits: list = field(default_factory=list)
    metrics: dict = field(default_factory=dict)

    def _weighted(self, levels: set[str]) -> int:
        if not self.chars:
            return 0
        weight = 0.0
        counted: dict[str, int] = {}
        for hit in self.hits:
            if hit.level not in levels:
                continue
            rule = RULE_BY_ID[hit.rule_id]
            n = counted.get(hit.rule_id, 0)
            counted[hit.rule_id] = n + 1
            # 同一规则第 3 次起降权，避免一个坏习惯把分数拉爆
            weight += rule.weight if n < 2 else rule.weight * 0.5
        # 短文本不做无限放大：分母至少按 500 字计
        denominator = max(self.chars, 500) / 1000.0
        return int(min(100, round(weight / denominator * 4)))

    @property
    def score(self) -> int:
        """AI 味分（S/A/B 级规则）。"""
        return self._weighted({"S", "A", "B"})

    @property
    def register_score(self) -> int:
        """语体分（R 级规则）：越高越大白话。"""
        return self._weighted({"R"})

    @property
    def defensive_score(self) -> int:
        """防御性分（D 级规则）：越高越像在为想象中的批评者写作。"""
        return self._weighted({"D"})


# ---------------------------------------------------------------- 语体（大白话）

# 「去 AI 味」的反面不是「大白话」。论文要的是学术语体：客观、有依据、有限定。
# 这组规则单独计分（level="R"），与 AI 味分数并列报告，两者都低才算合格。
REGISTER_RULES: list[Rule] = [
    Rule(
        "R1", "R", "口语动词/程度词",
        r"搞(?!活)|弄(?!虚)|整(?!理|体|合)|挺|超级|贼|巨(?!大|型)|不错|好用",
        "换成书面语：搞→开展/实施；弄→处理/构造；挺/超级→较为/显著（有数据才用）",
        3,
    ),
    Rule(
        "R2", "R", "语气词与口语衔接",
        r"(?:呢|吧|啊|嘛|啦|哈)(?=[，。；！？]|$)|然后呢|就是说|你看|你想想|说白了|反正",
        "删掉语气词与口语衔接词；用“因此/据此/另一方面/相比之下”承接",
        3,
    ),
    Rule(
        "R3", "R", "主观随意的表述",
        r"我觉得|我们觉得|我认为|我看|咱|咱们|笔者觉得|个人感觉",
        "改为客观陈述：本文认为／结果表明／数据支持……；避免第一人称口语",
        4,
    ),
    Rule(
        "R4", "R", "模糊不严谨",
        r"好像是|似乎吧|大概是|估计是|貌似|应该差不多|可能是吧|差不多是",
        "改为可核查表述：估计值 ± 不确定度，或注明“在……假设下”",
        3,
    ),
    Rule(
        "R5", "R", "口语化顺承",
        r"所以说|这不就|于是乎|接下来我们|下面我们就|我们可以看到",
        "改为“因此/由此可见/下文依次分析”；避免把论文写成讲课口吻",
        2,
    ),
    Rule(
        "R6", "R", "情绪化与感叹",
        r"太好了|太难了|棒极了|非常非常|极其极其|！{1,}",
        "客观陈述即可；感叹号在学术正文中基本不用",
        4,
    ),
]

# ---------------------------------------------------------------- 防御性写作

# 第三轴：为"想象中的批评者"写作。表现为免责声明、铺垫式让步、模糊限定叠加、
# 预设质疑的"不是…而是…"、结尾强行平衡。
# 与 AI 味（拔高/空转）方向相反但同源：都是"措辞强度 ≠ 证据强度"。
DEFENSIVE_RULES: list[Rule] = [
    Rule(
        "D1", "D", "免责声明与自我辩护",
        r"并非否认|并不是否认|无意贬低|无意否定|并非要否定|并不是要|不否认|"
        r"这并不意味着|但这并不代表|事先声明|需要声明的是",
        "删掉声明本身，直接写事实与结论；读者默认会善意理解",
        4,
    ),
    Rule(
        "D2", "D", "铺垫式让步",
        r"诚然|固然|不可否认|毋庸置疑|话虽如此|虽然如此|当然，也|也有其道理",
        "只在让步会实质改变结论或适用范围时保留，否则直接给判断",
        3,
    ),
    Rule(
        "D3", "D", "模糊限定叠加",
        r"(?:可能|或许|一定程度上|某种意义上|某种程度上|在某种程度上)"
        r"[^。；！？]{0,14}"
        r"(?:可能|或许|一定程度上|某种意义上|某种程度上|在某种程度上)",
        "叠加限定会稀释结论。确有不确定性时，明确指出不确定的是什么",
        3,
    ),
    Rule(
        "D4", "D", "预设质疑的“不是…而是…”",
        r"不是[^。；！？]{0,20}而是",
        "若文中没有人质疑过这一点，直接说它是什么、怎么运作、为何成立",
        2,
    ),
    Rule(
        "D5", "D", "结尾强行平衡",
        r"不能一概而论|仍需理性看待|应辩证|有待商榷|仁者见仁|见仁见智|"
        r"不能简单地说",
        "删掉收尾平衡句；让读者自己判断，或把边界条件写进方法与讨论",
        3,
    ),
    Rule(
        "D6", "D", "弱化到看不出结论",
        r"(?:似乎|或许|大致上|基本上)(?:可以)?(?:认为|说明|表明|证明了)",
        "把能确定的写成确定；不能确定的写成“在……条件下观察到”",
        2,
    ),
]

ALL_RULES = RULES + REGISTER_RULES + DEFENSIVE_RULES

RULE_BY_ID = {r.id: r for r in ALL_RULES}


# ---------------------------------------------------------------- 文本处理

def strip_noise(text: str) -> str:
    """剥离引用编号、引语、行内代码，避免误报。"""
    text = CITATION.sub("", text)
    text = QUOTED.sub("", text)
    return text


def split_sentences(text: str) -> list[str]:
    parts = [p.strip() for p in SENT_SPLIT.split(text)]
    return [p for p in parts if len(p) >= 4]


def collect_paragraphs_docx(path: str) -> tuple[list[tuple[str, bool]], str]:
    """返回 [(段落文本, 是否标题)]，跳过题注、表注与参考文献之后的内容。"""
    from docx import Document

    doc = Document(path)
    out: list[tuple[str, bool]] = []
    in_refs = False
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style = (p.style.name if p.style is not None else "") or ""
        is_heading = bool(HEADING_STYLE.match(style))
        if is_heading and REF_HEADING.match(text):
            in_refs = True
            continue
        if in_refs:
            continue
        if CAPTION.match(text) or text.startswith("关键词") or text.startswith("Keywords"):
            continue
        out.append((text, is_heading))
    return out, os.path.basename(path)


def collect_paragraphs_model(path: str) -> tuple[list[tuple[str, bool]], str]:
    with open(path, encoding="utf-8") as fh:
        model = json.load(fh)

    out: list[tuple[str, bool]] = []
    for key in ("abstract", "abstract_en"):
        block = model.get(key) or {}
        for text in block.get("paragraphs", []) or []:
            out.append((str(text).strip(), False))
    for item in model.get("sections", []) or []:
        if not isinstance(item, dict):
            continue
        if item.get("level") in (1, 2) or item.get("type") in ("h1", "h2"):
            out.append((str(item.get("text", "")).strip(), True))
        elif item.get("type") in (None, "p", "paragraph"):
            text = item.get("text", "")
            for chunk in (text if isinstance(text, list) else [text]):
                out.append((str(chunk).strip(), False))
        elif item.get("type") == "list":
            for chunk in item.get("items", []) or []:
                out.append((str(chunk).strip(), False))
    return [(t, h) for t, h in out if t], os.path.basename(path)


def collect_paragraphs_text(path: str) -> tuple[list[tuple[str, bool]], str]:
    with open(path, encoding="utf-8") as fh:
        raw = fh.read()
    out = []
    for block in re.split(r"\n\s*\n", raw):
        block = block.strip()
        if not block:
            continue
        if block.startswith("#"):
            out.append((block.lstrip("# ").strip(), True))
        else:
            out.append((block.replace("\n", ""), False))
    return out, os.path.basename(path)


# ---------------------------------------------------------------- 检测

def analyze(paragraphs: list[tuple[str, bool]], source: str) -> Report:
    report = Report(source=source)
    body = [(t, h) for t, h in paragraphs]
    report.paragraphs = len(body)

    all_chars = 0
    all_sentences: list[int] = []
    connective_count = 0
    number_count = 0
    heading_buffer: str | None = None

    for idx, (raw, is_heading) in enumerate(body, start=1):
        if is_heading:
            heading_buffer = raw
            continue
        text = strip_noise(raw)
        all_chars += len(text)
        sentences = split_sentences(text)
        all_sentences.extend(len(s) for s in sentences)
        connective_count += len(re.findall(CONNECTIVES, text))
        number_count += len(re.findall(_NUM, text))

        # A3.1 段落首句复读标题
        if heading_buffer and sentences:
            first = sentences[0]
            overlap = _overlap(heading_buffer, first)
            if overlap >= 0.6 and len(first) <= 40:
                report.hits.append(Hit(
                    "A3.1", "A", "段首复读标题", idx,
                    _excerpt(first),
                    "删掉复述标题的那句，直接给条件、定义或数据"))
        heading_buffer = None

        for rule in ALL_RULES:
            if rule.structural:
                continue      # 结构级规则由段落/全篇逻辑判定，不做逐句匹配
            for s_idx, sentence in enumerate(sentences):
                if rule.needs_no_number and re.search(_NUM, sentence):
                    continue
                if rule.paragraph_final and s_idx != len(sentences) - 1:
                    continue
                m = re.search(rule.pattern, sentence)
                if not m:
                    continue
                if rule.id == "A2.6" and not _dense_connectives(sentences):
                    continue      # 只有"每段都用"才报
                report.hits.append(Hit(rule.id, rule.level, rule.title, idx,
                                       _excerpt(sentence), rule.advice))
                break             # 同一段同一规则只报一次

    report.chars = all_chars
    report.sentences = len(all_sentences)
    report.metrics = {
        "句长均值": round(statistics.mean(all_sentences), 1) if all_sentences else 0,
        "句长变异": round(_cv(all_sentences), 2),
        "连接词/千字": round(connective_count / all_chars * 1000, 1) if all_chars else 0,
        "数字/千字": round(number_count / all_chars * 1000, 1) if all_chars else 0,
    }

    # A3.4 全文零数字
    if all_chars >= 400 and report.metrics["数字/千字"] < 1.0:
        report.hits.append(Hit("A3.4", "A", "全文零数字", 0, "(全篇)",
                               "至少给一个具体数字、区间或条件"))
    # A2.7 句长均匀（结构级）
    if len(all_sentences) >= 8 and report.metrics["句长变异"] < 0.35:
        report.hits.append(Hit("A2.7", "A", "句长均匀", 0,
                               "句长变异 %.2f（< 0.35）" % report.metrics["句长变异"],
                               "关键结论压成短句，限定条件并成长句；不要硬拆硬合"))
    return report


def _cv(values: list[int]) -> float:
    if len(values) < 2:
        return 1.0
    mean = statistics.mean(values)
    return statistics.pstdev(values) / mean if mean else 1.0


def _overlap(heading: str, sentence: str) -> float:
    """标题与句子的二元组重合率；以两者中较短的长度归一，避免长句被稀释。"""
    h = re.sub(rf"[^\w{CJK}]", "", heading)
    s = re.sub(rf"[^\w{CJK}]", "", sentence)
    if not h or not s:
        return 0.0
    h_len = max(len(h) - 1, 1)
    s_len = max(len(s) - 1, 1)
    grams = {h[i:i + 2] for i in range(h_len)}
    hits = sum(1 for i in range(s_len) if s[i:i + 2] in grams)
    return hits / min(h_len, s_len)


def _dense_connectives(sentences: list[str]) -> bool:
    if len(sentences) < 2:
        return False
    started = sum(1 for s in sentences if re.match(CONNECTIVES, s.strip()))
    return started >= max(2, len(sentences) // 2)


def _excerpt(text: str, limit: int = 46) -> str:
    text = text.strip()
    return text if len(text) <= limit else text[:limit] + "…"


# ---------------------------------------------------------------- 输出

LEVEL_LABEL = [(60, "重"), (40, "明显"), (20, "轻微"), (0, "清爽")]


REGISTER_LABEL = [(60, "大白话"), (40, "偏口语"), (20, "略松"), (0, "学术语体")]
DEFENSIVE_LABEL = [(60, "通篇防御"), (40, "明显防御"), (20, "略防御"), (0, "克制")]


def grade_defensive(score: int) -> str:
    for threshold, label in DEFENSIVE_LABEL:
        if score >= threshold:
            return label
    return "克制"


def grade_register(score: int) -> str:
    for threshold, label in REGISTER_LABEL:
        if score >= threshold:
            return label
    return "学术语体"


def grade(score: int) -> str:
    for threshold, label in LEVEL_LABEL:
        if score >= threshold:
            return label
    return "清爽"


def render(report: Report) -> str:
    lines = []
    lines.append("AI 味检测：%s" % report.source)
    lines.append("分析 %d 字 / %d 段 / %d 句" % (report.chars, report.paragraphs, report.sentences))
    lines.append("-" * 72)
    score = report.score
    reg = report.register_score
    dfe = report.defensive_score
    lines.append("AI 味 %d/100（%s）｜语体 %d/100（%s）｜防御性 %d/100（%s）"
                 % (score, grade(score), reg, grade_register(reg),
                    dfe, grade_defensive(dfe)))
    lines.append("三轴都要低（各 ≤20）：AI 味=空转拔高；语体=大白话；"
                 "防御性=为想象中的批评者写作。")
    lines.append("共同目标只有一个：措辞强度 = 证据强度。")
    m = report.metrics
    lines.append("句长均值 %s ｜ 句长变异 %s ｜ 连接词 %s/千字 ｜ 数字 %s/千字"
                 % (m.get("句长均值"), m.get("句长变异"),
                    m.get("连接词/千字"), m.get("数字/千字")))
    lines.append("-" * 72)

    if not report.hits:
        lines.append("没有命中可枚举的规则。仍建议按 §5 协议通读一遍：")
        lines.append("  references/anti-ai-academic-zh.md")
        return "\n".join(lines)

    grouped: dict[str, list[Hit]] = {}
    for hit in report.hits:
        grouped.setdefault(hit.rule_id, []).append(hit)

    for rule in ALL_RULES:
        hits = grouped.get(rule.id)
        if not hits:
            continue
        lines.append("[%s] %s（%s 级）×%d" % (rule.id, rule.title, rule.level, len(hits)))
        for hit in hits[:6]:
            where = "全篇" if hit.para_no == 0 else "¶%d" % hit.para_no
            lines.append("   %-6s %s" % (where, hit.excerpt))
        if len(hits) > 6:
            lines.append("   … 另有 %d 处" % (len(hits) - 6))
        if rule.advice:
            lines.append("   建议：%s" % rule.advice)
        lines.append("")

    by_level: dict[str, int] = {}
    for hit in report.hits:
        by_level[hit.level] = by_level.get(hit.level, 0) + 1
    lines.append("S 级 %d / A 级 %d / B 级 %d / R 级（大白话）%d / D 级（防御性）%d；"
                 "规则说明见 references/anti-ai-academic-zh.md"
                 % (by_level.get("S", 0), by_level.get("A", 0), by_level.get("B", 0),
                    by_level.get("R", 0), by_level.get("D", 0)))
    return "\n".join(lines)


def main(argv=None):
    _console.setup()
    ap = argparse.ArgumentParser(
        description="启发式检测学术中文的「AI 味」（只报警，不改稿）")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--docx", help="已生成的 .docx")
    src.add_argument("--json", help="内容模型 content.json")
    src.add_argument("--text", help="纯文本/Markdown（空行分段，# 开头视为标题）")
    ap.add_argument("--max-defensive", type=int, default=None,
                    help="防御性分阈值（越高越防御）；超阈值退出码 1")
    ap.add_argument("--max-register", type=int, default=None,
                    help="语体分阈值（越大白越大）；超阈值退出码 1")
    ap.add_argument("--max-score", type=int, default=None,
                    help="分数超过该值时退出码为 1（用于 CI/交付门禁）")
    ap.add_argument("--json-out", action="store_true", help="输出 JSON 而不是报告文本")
    args = ap.parse_args(argv)

    if args.docx:
        paragraphs, name = collect_paragraphs_docx(args.docx)
    elif args.json:
        paragraphs, name = collect_paragraphs_model(args.json)
    else:
        paragraphs, name = collect_paragraphs_text(args.text)

    if not paragraphs:
        print("没有可分析的正文段落（检查输入文件）", file=sys.stderr)
        return 2

    report = analyze(paragraphs, name)
    if args.json_out:
        print(json.dumps({
            "source": report.source,
            "score": report.score,
            "grade": grade(report.score),
            "register_score": report.register_score,
            "defensive_score": report.defensive_score,
            "chars": report.chars,
            "sentences": report.sentences,
            "metrics": report.metrics,
            "hits": [h.__dict__ for h in report.hits],
        }, ensure_ascii=False, indent=2))
    else:
        print(render(report))

    if args.max_score is not None and report.score > args.max_score:
        print("\n超过阈值 %d（实际 %d）。" % (args.max_score, report.score), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
