#!/usr/bin/env python3
"""Lint reader-first Chinese research reports.

The script has no third-party dependencies. It checks plain text/Markdown and
performs additional structural checks for DOCX files.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}


@dataclass
class Issue:
    severity: str
    code: str
    message: str
    context: str = ""


BANNED_PATTERNS = [
    (
        "BANNED_NOT_BUT",
        re.compile(r"不是[^。！？；\n]{0,80}(?:，|,)?\s*而是"),
        "使用了禁用结构“不是……而是……”。直接陈述需要保留的事实。",
    ),
    (
        "BANNED_NOT_BUT",
        re.compile(r"并非[^。！？；\n]{0,80}(?:，|,)?\s*而是"),
        "使用了禁用结构“并非……而是……”。直接陈述需要保留的事实。",
    ),
    (
        "BANNED_NOT_BUT",
        re.compile(r"并不是[^。！？；\n]{0,80}(?:，|,)?\s*(?:而是|而在于)"),
        "使用了禁用结构“并不是……而是/而在于……”。",
    ),
    (
        "BANNED_FOCUS_CONTRAST",
        re.compile(r"重点不在[^。！？；\n]{0,80}(?:，|,)?\s*而在"),
        "使用了禁用的二分转折。直接写出重点内容。",
    ),
]


SCOPE_PATTERNS = [
    (re.compile(r"不具有代表性|缺乏代表性"), "写明当前数据规模、重复次数和下一步验证计划。"),
    (re.compile(r"(?:结果|优势|表现)不稳定"), "改写为具体数据表现、变化范围和下一步重复验证。"),
    (re.compile(r"存在(?:明显|一定)?(?:不足|局限)"), "改写为当前证据覆盖的条件和待补充的验证。"),
    (
        re.compile(r"(?:实验|验证)仅在[^。！？；\n]{0,40}(?:数据集|样本|场景)"),
        "明确当前实验范围，并写出扩大数据或重复验证的计划。",
    ),
    (
        re.compile(r"(?:当前|现阶段|目前)(?:尚)?不(?:判断|评价|讨论|支持|确定)"),
        "改写为现有证据或统计量，并写出用于完成判断的下一步工作。",
    ),
    (
        re.compile(r"仍需[^。！？；\n]{0,40}才能"),
        "改写为下一步将补充的内容及其用途。",
    ),
]


META_PATTERNS = [
    re.compile(r"推进逻辑|写作思路|组织框架|分析路径|叙述结构"),
    re.compile(r"下文将|接下来(?:将|介绍|说明|讨论)|下面(?:将|介绍|说明|讨论)"),
    re.compile(r"为了便于理解[^。！？；\n]{0,20}(?:先|首先)"),
]


DERIVED_CONTENT_PATTERN = re.compile(
    r"(?:训练|测试|候选|样本|数据|模型)?(?:名单顺序|名单|排序结果|排名结果|分组结果|筛选结果)(?:如下|为)?"
)
GENERATION_CUE_PATTERN = re.compile(
    r"(?:按|按照|根据|基于|依据)[^。！？；\n]{0,60}"
    r"|(?:升序|降序|由高到低|从高到低|由低到高|从低到高)"
    r"|(?:脚本|程序|配置|人工|自动生成|生成方式|并列|同分)"
)


CN_FONT_ALIASES = {"宋体": "宋体", "SimSun": "宋体", "黑体": "黑体", "SimHei": "黑体"}


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def w_attr(element: ET.Element | None, name: str) -> str | None:
    if element is None:
        return None
    return element.get(f"{W}{name}")


def visible_border(element: ET.Element | None) -> bool:
    return element is not None and (w_attr(element, "val") or "single") not in {"none", "nil"}


def context_for(text: str, start: int, end: int) -> str:
    line = text.count("\n", 0, start) + 1
    left = max(0, start - 28)
    right = min(len(text), end + 28)
    snippet = re.sub(r"\s+", " ", text[left:right]).strip()
    return f"第 {line} 行：{snippet}"


def extract_docx(path: Path) -> tuple[str, dict[str, ET.Element]]:
    with zipfile.ZipFile(path) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
        parts = {"document": document}
        if "word/styles.xml" in archive.namelist():
            parts["styles"] = ET.fromstring(archive.read("word/styles.xml"))

    paragraphs = []
    for paragraph in document.iter(f"{W}p"):
        pieces = [node.text or "" for node in paragraph.iter(f"{W}t")]
        if pieces:
            paragraphs.append("".join(pieces))
    return "\n".join(paragraphs), parts


def read_source(path: Path | None, direct_text: str | None) -> tuple[str, dict[str, ET.Element]]:
    if direct_text is not None:
        return direct_text, {}
    if path is None:
        raise ValueError("需要提供文件路径或 --text。")
    if path.suffix.lower() == ".docx":
        return extract_docx(path)
    return path.read_text(encoding="utf-8"), {}


def lint_language(text: str) -> list[Issue]:
    issues: list[Issue] = []
    for code, pattern, message in BANNED_PATTERNS:
        for match in pattern.finditer(text):
            issues.append(Issue("error", code, message, context_for(text, *match.span())))

    for pattern, message in SCOPE_PATTERNS:
        for match in pattern.finditer(text):
            issues.append(Issue("warning", "SCOPE_NEGATIVE", message, context_for(text, *match.span())))

    for pattern in META_PATTERNS:
        for match in pattern.finditer(text):
            issues.append(
                Issue(
                    "warning",
                    "META_DISCOURSE",
                    "发现元话语。直接陈述研究对象、实验动作、结果或计划。",
                    context_for(text, *match.span()),
                )
            )

    for match in DERIVED_CONTENT_PATTERN.finditer(text):
        left = max(0, match.start() - 120)
        right = min(len(text), match.end() + 120)
        if not GENERATION_CUE_PATTERN.search(text[left:right]):
            issues.append(
                Issue(
                    "warning",
                    "DERIVED_CONTENT_SOURCE",
                    "名单、顺序、分组或筛选结果缺少生成依据。说明对象、规则、排列方向和生成方式。",
                    context_for(text, *match.span()),
                )
            )

    duplicate_pattern = re.compile(r"([\u4e00-\u9fff]{2,6})\1")
    seen_spans: set[tuple[int, int]] = set()
    for match in duplicate_pattern.finditer(text):
        span = match.span()
        if span in seen_spans:
            continue
        seen_spans.add(span)
        issues.append(
            Issue(
                "error",
                "DUPLICATE_WORD",
                f"发现连续重复词“{match.group(1)}”。",
                context_for(text, *span),
            )
        )

    for match in re.finditer(r"\[待补充：[^\]]+\]", text):
        issues.append(Issue("warning", "MISSING_FIELD", "报告仍包含待补充字段。", context_for(text, *match.span())))
    return issues


def load_terms(path: Path) -> list[dict]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, dict):
        entries = []
        for canonical, value in payload.items():
            value = value if isinstance(value, dict) else {}
            entries.append(
                {
                    "canonical": canonical,
                    "allowed": value.get("allowed", []),
                    "avoid": value.get("avoid", []),
                }
            )
        return entries
    if isinstance(payload, list):
        return payload
    raise ValueError("术语文件必须是对象或数组。")


def lint_terms(text: str, entries: Iterable[dict]) -> list[Issue]:
    issues: list[Issue] = []
    for entry in entries:
        canonical = str(entry.get("canonical", "")).strip()
        if not canonical:
            continue
        for variant in entry.get("avoid", []):
            variant = str(variant).strip()
            if not variant:
                continue
            for match in re.finditer(re.escape(variant), text):
                issues.append(
                    Issue(
                        "error",
                        "TERM_VARIANT",
                        f"术语“{variant}”应统一为“{canonical}”。",
                        context_for(text, *match.span()),
                    )
                )
        canonical_pos = text.find(canonical)
        for abbreviation in entry.get("allowed", []):
            abbreviation = str(abbreviation).strip()
            if not abbreviation:
                continue
            short_pos = text.find(abbreviation)
            if short_pos >= 0 and (canonical_pos < 0 or short_pos < canonical_pos):
                issues.append(
                    Issue(
                        "warning",
                        "TERM_BEFORE_DEFINITION",
                        f"“{abbreviation}”出现在“{canonical}”定义之前。",
                        context_for(text, short_pos, short_pos + len(abbreviation)),
                    )
                )
    return issues


def check_default_fonts(styles: ET.Element | None) -> list[Issue]:
    issues: list[Issue] = []
    if styles is None:
        return [Issue("warning", "DOCX_STYLES", "DOCX 缺少 styles.xml，无法检查默认字体。")]
    fonts = styles.find("./w:docDefaults/w:rPrDefault/w:rPr/w:rFonts", NS)
    if fonts is None:
        return [Issue("warning", "DOCX_DEFAULT_FONT", "未设置默认字体。")]
    east_asia = w_attr(fonts, "eastAsia")
    ascii_font = w_attr(fonts, "ascii")
    hansi = w_attr(fonts, "hAnsi")
    if CN_FONT_ALIASES.get(east_asia or "") != "宋体":
        issues.append(Issue("error", "DOCX_CN_FONT", f"默认中文字体为“{east_asia or '未设置'}”，要求使用宋体。"))
    if ascii_font != "Times New Roman" or hansi != "Times New Roman":
        issues.append(Issue("error", "DOCX_EN_FONT", "默认英文和数字字体要求使用 Times New Roman。"))
    return issues


def check_page(document: ET.Element) -> list[Issue]:
    issues: list[Issue] = []
    sections = list(document.iter(f"{W}sectPr"))
    if not sections:
        return [Issue("error", "DOCX_PAGE", "未找到页面设置。")]
    for index, section in enumerate(sections, start=1):
        size = section.find("w:pgSz", NS)
        margins = section.find("w:pgMar", NS)
        width = int(w_attr(size, "w") or 0)
        height = int(w_attr(size, "h") or 0)
        if abs(width - 11906) > 50 or abs(height - 16838) > 50:
            issues.append(Issue("error", "DOCX_A4", f"第 {index} 节未使用 A4 纵向页面。"))
        expected = {"top": 1440, "bottom": 1440, "left": 1800, "right": 1800}
        for name, target in expected.items():
            actual = int(w_attr(margins, name) or 0)
            if abs(actual - target) > 30:
                issues.append(
                    Issue(
                        "error",
                        "DOCX_MARGIN",
                        f"第 {index} 节{name}页边距为 {actual / 1440:.2f} 英寸，要求 {target / 1440:.2f} 英寸。",
                    )
                )
    return issues


def check_runs(document: ET.Element) -> list[Issue]:
    issues: list[Issue] = []
    seen: set[tuple[str, str]] = set()
    for run in document.iter(f"{W}r"):
        text = "".join(node.text or "" for node in run.iter(f"{W}t"))
        if not text.strip():
            continue
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            continue
        fonts = rpr.find("w:rFonts", NS)
        if fonts is not None:
            if re.search(r"[\u4e00-\u9fff]", text):
                font = w_attr(fonts, "eastAsia")
                theme_font = w_attr(fonts, "eastAsiaTheme")
                if font and font not in CN_FONT_ALIASES and ("CN_FONT", font) not in seen:
                    seen.add(("CN_FONT", font))
                    issues.append(Issue("error", "DOCX_CN_FONT", f"发现中文字体“{font}”，只允许宋体或黑体。", text[:40]))
                if not font and theme_font and ("CN_THEME_FONT", theme_font) not in seen:
                    seen.add(("CN_THEME_FONT", theme_font))
                    issues.append(
                        Issue(
                            "error",
                            "DOCX_THEME_FONT",
                            f"发现中文主题字体“{theme_font}”，请显式设置为宋体或黑体。",
                            text[:40],
                        )
                    )
            if re.search(r"[A-Za-z0-9]", text):
                font = w_attr(fonts, "ascii") or w_attr(fonts, "hAnsi")
                theme_font = w_attr(fonts, "asciiTheme") or w_attr(fonts, "hAnsiTheme")
                if font and font != "Times New Roman" and ("EN_FONT", font) not in seen:
                    seen.add(("EN_FONT", font))
                    issues.append(
                        Issue("error", "DOCX_EN_FONT", f"发现英文或数字字体“{font}”，要求使用 Times New Roman。", text[:40])
                    )
                if not font and theme_font and ("EN_THEME_FONT", theme_font) not in seen:
                    seen.add(("EN_THEME_FONT", theme_font))
                    issues.append(
                        Issue(
                            "error",
                            "DOCX_THEME_FONT",
                            f"发现英文或数字主题字体“{theme_font}”，请显式设置为 Times New Roman。",
                            text[:40],
                        )
                    )
        color = rpr.find("w:color", NS)
        value = (w_attr(color, "val") or "auto").upper()
        if value not in {"AUTO", "000000"} and ("COLOR", value) not in seen:
            seen.add(("COLOR", value))
            issues.append(Issue("error", "DOCX_COLOR", f"发现非黑色文字 #{value}。", text[:40]))
    return issues


def border_in(parent: ET.Element, name: str, path: str) -> ET.Element | None:
    return parent.find(f"{path}/w:{name}", NS)


def check_table(table: ET.Element, index: int) -> list[Issue]:
    issues: list[Issue] = []
    rows = table.findall("w:tr", NS)
    if not rows:
        return issues

    for shading in table.iter(f"{W}shd"):
        fill = (w_attr(shading, "fill") or "auto").upper()
        if fill not in {"AUTO", "FFFFFF", "NONE"}:
            issues.append(Issue("error", "TABLE_SHADING", f"表 {index} 使用了底色 #{fill}。"))
            break

    for border in table.iter():
        if local_name(border.tag) in {"left", "right", "insideV", "start", "end"} and visible_border(border):
            issues.append(Issue("error", "TABLE_VERTICAL_BORDER", f"表 {index} 存在竖线，要求使用无竖线三线表。"))
            break

    first_cells = rows[0].findall("w:tc", NS)
    last_cells = rows[-1].findall("w:tc", NS)
    table_top = border_in(table, "top", "w:tblPr/w:tblBorders")
    table_bottom = border_in(table, "bottom", "w:tblPr/w:tblBorders")
    has_top = visible_border(table_top) or all(
        visible_border(border_in(cell, "top", "w:tcPr/w:tcBorders")) for cell in first_cells
    )
    has_header_rule = all(
        visible_border(border_in(cell, "bottom", "w:tcPr/w:tcBorders")) for cell in first_cells
    )
    has_bottom = visible_border(table_bottom) or all(
        visible_border(border_in(cell, "bottom", "w:tcPr/w:tcBorders")) for cell in last_cells
    )
    if not has_top:
        issues.append(Issue("error", "TABLE_TOP_RULE", f"表 {index} 缺少顶线。"))
    if not has_header_rule:
        issues.append(Issue("error", "TABLE_HEADER_RULE", f"表 {index} 缺少表头下边线。"))
    if not has_bottom:
        issues.append(Issue("error", "TABLE_BOTTOM_RULE", f"表 {index} 缺少底线。"))

    for cell_index, cell in enumerate(first_cells, start=1):
        runs = [run for run in cell.iter(f"{W}r") if any((node.text or "").strip() for node in run.iter(f"{W}t"))]
        if runs and not all(run.find("w:rPr/w:b", NS) is not None for run in runs):
            issues.append(Issue("warning", "TABLE_HEADER_BOLD", f"表 {index} 表头第 {cell_index} 列未显式加粗。"))
        for run in runs:
            text = "".join(node.text or "" for node in run.iter(f"{W}t"))
            if not re.search(r"[\u4e00-\u9fff]", text):
                continue
            font = w_attr(run.find("w:rPr/w:rFonts", NS), "eastAsia")
            if CN_FONT_ALIASES.get(font or "") != "黑体":
                issues.append(Issue("warning", "TABLE_HEADER_FONT", f"表 {index} 表头第 {cell_index} 列中文未显式使用黑体。"))
                break

    column_count = max((len(row.findall("w:tc", NS)) for row in rows), default=0)
    expected_size = "21" if column_count <= 5 else "18"
    for run in table.iter(f"{W}r"):
        text = "".join(node.text or "" for node in run.iter(f"{W}t"))
        if not text.strip():
            continue
        size = run.find("w:rPr/w:sz", NS)
        value = w_attr(size, "val")
        if value and value != expected_size:
            issues.append(
                Issue(
                    "warning",
                    "TABLE_FONT_SIZE",
                    f"表 {index} 含 {int(value) / 2:g} 磅文字；当前列数要求 {int(expected_size) / 2:g} 磅。",
                    text[:30],
                )
            )
            break
    return issues


def lint_docx(parts: dict[str, ET.Element]) -> list[Issue]:
    document = parts["document"]
    issues = check_default_fonts(parts.get("styles"))
    issues.extend(check_page(document))
    issues.extend(check_runs(document))
    for index, table in enumerate(document.iter(f"{W}tbl"), start=1):
        issues.extend(check_table(table, index))
    return issues


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="检查科研报告的语言、术语和 DOCX 基础格式。")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("source", nargs="?", type=Path, help="UTF-8 文本、Markdown 或 DOCX 文件")
    group.add_argument("--text", help="直接检查给定文本")
    parser.add_argument("--terms", type=Path, help="JSON 术语表")
    parser.add_argument("--strict", action="store_true", help="将警告也视为失败")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出结果")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.source is not None and not args.source.exists():
        print(f"文件不存在：{args.source}", file=sys.stderr)
        return 2
    try:
        text, parts = read_source(args.source, args.text)
        issues = lint_language(text)
        if args.terms:
            issues.extend(lint_terms(text, load_terms(args.terms)))
        if parts:
            issues.extend(lint_docx(parts))
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError, json.JSONDecodeError) as error:
        print(f"检查失败：{error}", file=sys.stderr)
        return 2

    errors = sum(issue.severity == "error" for issue in issues)
    warnings = sum(issue.severity == "warning" for issue in issues)
    if args.json:
        print(json.dumps({"errors": errors, "warnings": warnings, "issues": [asdict(i) for i in issues]}, ensure_ascii=False, indent=2))
    else:
        label = str(args.source) if args.source is not None else "直接文本"
        print(f"检查对象：{label}")
        print(f"错误：{errors}；警告：{warnings}")
        for issue in issues:
            suffix = f" | {issue.context}" if issue.context else ""
            print(f"[{issue.severity.upper()}] {issue.code}: {issue.message}{suffix}")
    return 1 if errors or (args.strict and warnings) else 0


if __name__ == "__main__":
    raise SystemExit(main())
