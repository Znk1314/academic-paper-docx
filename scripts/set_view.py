#!/usr/bin/env python
"""只改 word/settings.xml：强制默认页面视图（print）+ 缩放 100%。

为什么需要单独一个脚本：
    Word COM 保存后可能把视图改回 Web 版式；规范要求"最后一次保存之后"
    再设置 w:view=print。本脚本只重写 settings.xml，其它部件字节不变，
    因此在 Word 保存之后运行是安全的（不会破坏已更新的目录域）。

用法：
    python set_view.py report.docx
    python set_view.py report.docx --zoom 100 --check
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SETTINGS_PART = "word/settings.xml"

# settings.xml 里需要保持在正确位置的元素（只列关键部分）
ORDER = [
    "writeProtection", "view", "zoom", "removePersonalInformation",
    "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape",
    "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData",
    "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData",
    "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign",
    "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
    "stylePaneSortMethod", "documentType", "mailMerge", "revisionView",
    "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting",
    "documentProtection", "autoFormatOverride", "styleLockTheme",
    "styleLockQFSet", "defaultTabStop", "autoHyphenation",
    "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
    "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
    "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
    "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
    "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
    "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
    "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
    "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
    "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
    "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
]


def _local(el) -> str:
    return el.tag.split("}")[-1] if isinstance(el.tag, str) else ""


def _insert_ordered(root, el):
    tag = _local(el)
    if tag not in ORDER:
        root.append(el)
        return
    idx = ORDER.index(tag)
    for child in list(root):
        ctag = _local(child)
        if ctag in ORDER and ORDER.index(ctag) > idx:
            child.addprevious(el)
            return
    root.append(el)


def _drop(root, tag):
    for el in root.findall("{%s}%s" % (W, tag)):
        root.remove(el)


def set_print_view(path: str, zoom: int = 100) -> None:
    path = os.path.abspath(path)
    with zipfile.ZipFile(path) as zf:
        infos = zf.infolist()
        blobs = {info.filename: zf.read(info.filename) for info in infos}

    if SETTINGS_PART not in blobs:
        raise SystemExit("不是有效的 docx（缺少 %s）：%s" % (SETTINGS_PART, path))

    root = etree.fromstring(blobs[SETTINGS_PART])
    _drop(root, "view")
    _drop(root, "zoom")
    view = etree.SubElement(root, "{%s}view" % W)
    view.set("{%s}val" % W, "print")
    _insert_ordered(root, view)

    z = etree.Element("{%s}zoom" % W)
    z.set("{%s}percent" % W, str(int(zoom)))
    _insert_ordered(root, z)

    blobs[SETTINGS_PART] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)

    fd, tmp = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for info in infos:
                out.writestr(info, blobs[info.filename])
        shutil.move(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def read_view(path: str):
    with zipfile.ZipFile(os.path.abspath(path)) as zf:
        root = etree.fromstring(zf.read(SETTINGS_PART))
    view = root.find("{%s}view" % W)
    zoom = root.find("{%s}zoom" % W)
    return (
        view.get("{%s}val" % W) if view is not None else None,
        zoom.get("{%s}percent" % W) if zoom is not None else None,
    )


def main(argv=None):
    ap = argparse.ArgumentParser(description="把 docx 默认视图固定为页面视图 + 100%")
    ap.add_argument("docx")
    ap.add_argument("--zoom", type=int, default=100)
    ap.add_argument("--check", action="store_true", help="只检查，不修改")
    args = ap.parse_args(argv)

    if args.check:
        view, zoom = read_view(args.docx)
        ok = (view == "print" and str(zoom) == "100")
        print("%s: view=%s zoom=%s %s" % (args.docx, view, zoom,
                                          "OK" if ok else "需要修复"))
        return 0 if ok else 1

    set_print_view(args.docx, args.zoom)
    view, zoom = read_view(args.docx)
    print("已设置 %s: view=%s zoom=%s" % (args.docx, view, zoom))
    return 0


if __name__ == "__main__":
    sys.exit(main())
