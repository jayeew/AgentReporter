from __future__ import annotations

import re
import shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree
from PIL import Image, ImageDraw, ImageFont


ROOT = Path("/mnt/c/Users/jayee/Documents/肿瘤agent汇报")
OUT = ROOT / "outputs"
WORK = OUT / "work_ppt12_15"
EXTRACTED = WORK / "extracted"
PPTX = OUT / "final_12_15_adjusted.pptx"
HANDOUT = OUT / "oncology_agent_report_v4_逐页讲义与版式设计.md"

W, H = 1672, 941
NAVY = (6, 35, 82)
TEAL = (4, 124, 138)
TEAL_DARK = (0, 104, 120)
TEAL_LIGHT = (167, 215, 224)
TEXT = (20, 38, 77)
MUTED = (78, 96, 126)
BG = (248, 252, 254)
LINE = (97, 185, 195)

FONT_REG = "/mnt/c/Windows/Fonts/msyh.ttc"
FONT_BOLD = "/mnt/c/Windows/Fonts/msyhbd.ttc"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, size)


def text_size(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.FreeTypeFont) -> tuple[int, int]:
    box = draw.textbbox((0, 0), text, font=fnt)
    return box[2] - box[0], box[3] - box[1]


def wrap_text(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    lines: list[str] = []
    for para in text.split("\n"):
        if not para:
            lines.append("")
            continue
        buf = ""
        for ch in para:
            trial = buf + ch
            if text_size(draw, trial, fnt)[0] <= max_width:
                buf = trial
            else:
                if buf:
                    lines.append(buf)
                buf = ch
        if buf:
            lines.append(buf)
    return lines


def draw_wrapped(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    fnt: ImageFont.FreeTypeFont,
    fill=TEXT,
    max_width: int = 300,
    line_gap: int = 10,
) -> int:
    x, y = xy
    for line in wrap_text(draw, text, fnt, max_width):
        draw.text((x, y), line, font=fnt, fill=fill)
        y += text_size(draw, line or " ", fnt)[1] + line_gap
    return y


def make_background() -> Image.Image:
    img = Image.new("RGB", (W, H), BG)
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    # Subtle DNA-like motif on the right.
    for i in range(0, 8):
        y = 20 + i * 86
        d.arc((1480, y, 1688, y + 170), start=95, end=265, fill=(148, 190, 211, 34), width=8)
        d.arc((1515, y + 35, 1723, y + 205), start=275, end=85, fill=(148, 190, 211, 28), width=8)
        d.line((1585, y + 45, 1640, y + 80), fill=(148, 190, 211, 25), width=5)

    # Light molecule network bottom corners.
    for base_x, flip in [(22, 1), (1505, -1)]:
        pts = [(base_x, 720), (base_x + 70 * flip, 680), (base_x + 140 * flip, 725),
               (base_x + 105 * flip, 805), (base_x + 190 * flip, 850), (base_x + 255 * flip, 805)]
        for a, b in zip(pts, pts[1:]):
            d.line((a[0], a[1], b[0], b[1]), fill=(148, 190, 211, 45), width=4)
        for p in pts:
            d.ellipse((p[0] - 8, p[1] - 8, p[0] + 8, p[1] + 8), fill=(148, 190, 211, 50))

    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    return img


def draw_page_header(draw: ImageDraw.ImageDraw, page: str, title: str, subtitle: str | None = None):
    draw.rounded_rectangle((-9, 0, 63, 50), radius=10, fill=TEAL)
    tw, th = text_size(draw, page, font(28, True))
    draw.text((31 - tw / 2, 10), page, font=font(28, True), fill="white")
    draw.rounded_rectangle((94, 47, 105, 142), radius=6, fill=TEAL)
    draw.text((142, 66), title, font=font(57, True), fill=NAVY)
    if subtitle:
        draw.text((143, 154), subtitle, font=font(29, False), fill=TEXT)


def icon_circle(draw: ImageDraw.ImageDraw, cx: int, cy: int, r: int = 40):
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), outline=TEAL, width=3, fill=(255, 255, 255))


def draw_check(draw: ImageDraw.ImageDraw, cx: int, cy: int, r: int = 28):
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=TEAL)
    draw.line((cx - 13, cy, cx - 3, cy + 11, cx + 16, cy - 13), fill="white", width=7, joint="curve")


def draw_doc_icon(draw: ImageDraw.ImageDraw, x: int, y: int, s: int = 90):
    draw.rounded_rectangle((x, y, x + s * 0.74, y + s), radius=8, outline=TEAL, width=4)
    draw.polygon([(x + s * 0.52, y), (x + s * 0.74, y + s * 0.22), (x + s * 0.52, y + s * 0.22)],
                 outline=TEAL, fill=BG)
    for i in range(3):
        yy = y + 28 + i * 18
        draw.line((x + 16, yy, x + s * 0.55, yy), fill=TEAL, width=4)
    draw.ellipse((x + 18, y + s - 36, x + 42, y + s - 12), fill=TEAL)
    draw.rectangle((x + 48, y + s - 31, x + 62, y + s - 26), fill=TEAL)


def draw_simple_icon(draw: ImageDraw.ImageDraw, kind: str, cx: int, cy: int, scale: int = 1):
    icon_circle(draw, cx, cy, 38)
    if kind == "book":
        draw.rectangle((cx - 25, cy - 16, cx - 3, cy + 18), outline=TEAL, width=3)
        draw.rectangle((cx + 3, cy - 16, cx + 25, cy + 18), outline=TEAL, width=3)
        draw.line((cx, cy - 18, cx, cy + 20), fill=TEAL, width=3)
    elif kind == "trial":
        draw.rounded_rectangle((cx - 20, cy - 24, cx + 20, cy + 24), radius=5, outline=TEAL, width=3)
        for i in range(3):
            draw.line((cx - 9, cy - 10 + i * 12, cx + 14, cy - 10 + i * 12), fill=TEAL, width=3)
        draw.ellipse((cx - 29, cy + 10, cx - 12, cy + 27), outline=TEAL, width=3)
        draw.ellipse((cx + 12, cy + 10, cx + 29, cy + 27), outline=TEAL, width=3)
    elif kind == "dna":
        for t in range(-22, 25, 11):
            draw.line((cx - 18, cy + t, cx + 18, cy - t), fill=TEAL, width=2)
        draw.arc((cx - 25, cy - 28, cx + 25, cy + 28), 90, 270, fill=TEAL, width=3)
        draw.arc((cx - 25, cy - 28, cx + 25, cy + 28), 270, 90, fill=TEAL, width=3)
    elif kind == "mdt":
        for dx, dy in [(-17, 0), (0, -8), (17, 0)]:
            draw.ellipse((cx + dx - 8, cy + dy - 8, cx + dx + 8, cy + dy + 8), fill=TEAL)
            draw.rectangle((cx + dx - 12, cy + dy + 10, cx + dx + 12, cy + dy + 26), fill=TEAL)
        draw.rounded_rectangle((cx - 25, cy - 30, cx + 25, cy - 12), radius=8, outline=TEAL, width=3)
    elif kind == "multi":
        draw.rounded_rectangle((cx - 22, cy - 17, cx + 22, cy + 20), radius=8, outline=TEAL, width=3)
        draw.arc((cx - 12, cy - 4, cx + 12, cy + 20), 180, 360, fill=TEAL, width=3)
        draw.ellipse((cx - 8, cy - 28, cx + 8, cy - 12), outline=TEAL, width=3)
    elif kind == "split":
        draw.rounded_rectangle((cx - 23, cy - 23, cx + 23, cy + 23), radius=8, outline=TEAL, width=3)
        draw.line((cx, cy - 6, cx, cy + 15), fill=TEAL, width=3)
        draw.line((cx, cy + 15, cx - 16, cy + 25), fill=TEAL, width=3)
        draw.line((cx, cy + 15, cx + 16, cy + 25), fill=TEAL, width=3)
    elif kind == "cite":
        draw_doc_icon(draw, cx - 26, cy - 31, 62)
        draw.ellipse((cx + 8, cy + 8, cx + 31, cy + 31), outline=TEAL, width=4)
        draw.line((cx + 25, cy + 25, cx + 38, cy + 38), fill=TEAL, width=5)
    elif kind == "warn":
        draw.polygon([(cx, cy - 28), (cx - 31, cy + 26), (cx + 31, cy + 26)], outline=TEAL, fill=(235, 250, 252), width=4)
        draw.line((cx, cy - 10, cx, cy + 10), fill=TEAL, width=5)
        draw.ellipse((cx - 3, cy + 17, cx + 3, cy + 23), fill=TEAL)


def slide12() -> Path:
    img = make_background()
    d = ImageDraw.Draw(img)
    draw_page_header(d, "12", "Demo 简介：Onco-MDT Evidence Brief", "一个用于 MDT 会前证据整理的 Agent 工作流演示")

    # Three-column information flow.
    columns = [
        (96, 232, 440, "使用场景", [
            ("MDT 会前准备", "把病例信息转成可讨论问题"),
            ("证据快速整理", "聚合指南、文献和试验线索"),
            ("人工复核前置", "提前暴露不确定性和缺失资料"),
        ]),
        (610, 232, 455, "输入示意", []),
        (1210, 232, 365, "预期输出", []),
    ]
    for x, y, w, title, _ in columns:
        d.rounded_rectangle((x, y, x + w, 672), radius=22, outline=TEAL_LIGHT, width=2, fill=(255, 255, 255))
        d.text((x + 34, y + 28), title, font=font(30, True), fill=NAVY)
        d.line((x + 34, y + 82, x + w - 34, y + 82), fill=LINE, width=2)

    # Left: use cases.
    use_items = columns[0][4]
    icon_kinds = ["mdt", "cite", "warn"]
    for i, ((head, body), kind) in enumerate(zip(use_items, icon_kinds)):
        y = 340 + i * 105
        draw_simple_icon(d, kind, 150, y + 22)
        d.text((205, y), head, font=font(25, True), fill=NAVY)
        draw_wrapped(d, (205, y + 36), body, font(20), fill=MUTED, max_width=285, line_gap=4)

    # Middle: cached chat input.
    d.rounded_rectangle((654, 340, 1022, 550), radius=18, outline=(190, 224, 232), width=2, fill=(244, 251, 252))
    d.text((686, 372), "模拟肿瘤病例聊天输入", font=font(25, True), fill=TEAL_DARK)
    input_lines = ["黑色素瘤病史", "分期 / 治疗经过", "BRAF 状态", "当前 MDT 关注问题"]
    for i, line in enumerate(input_lines):
        yy = 424 + i * 34
        d.ellipse((690, yy + 8, 700, yy + 18), fill=TEAL)
        d.text((718, yy), line, font=font(22, True), fill=TEXT)
    d.rounded_rectangle((700, 594, 975, 632), radius=18, fill=TEAL)
    d.text((735, 601), "去身份化模拟病例", font=font(20, True), fill="white")

    # Arrows toward output.
    d.line((1065, 452, 1170, 452), fill=TEAL, width=7)
    d.polygon([(1170, 424), (1210, 452), (1170, 480)], fill=TEAL)

    # Right: expected output modules.
    outputs = ["病例摘要", "MDT 关键问题", "证据摘要", "临床试验预筛线索", "缺失信息与不确定性", "MDT 讨论清单", "人工复核点", "TXT 简报导出"]
    for i, label in enumerate(outputs):
        yy = 325 + i * 38
        d.rounded_rectangle((1240, yy, 1538, yy + 30), radius=9, outline=(194, 224, 230), width=1, fill=(248, 252, 254))
        d.ellipse((1255, yy + 10, 1265, yy + 20), fill=TEAL)
        d.text((1282, yy + 1), label, font=font(19, True), fill=TEXT)

    checks = [
        "模拟病例，不用真实患者数据",
        "输出证据简报，不输出治疗医嘱",
        "所有结果均需人工复核",
    ]
    for i, label in enumerate(checks):
        x = 130 + i * 490
        d.rounded_rectangle((x, 765, x + 438, 843), radius=18, outline=TEAL, width=2, fill=(255, 255, 255))
        draw_check(d, x + 52, 804, 29)
        d.text((x + 96, 786), label, font=font(24, True), fill=NAVY)

    out = WORK / "new_slide12.png"
    img.save(out)
    return out


def slide13_from_old() -> Path:
    img = Image.open(EXTRACTED / "slide15_old.png").convert("RGB")
    d = ImageDraw.Draw(img)
    # Replace page badge 14 -> 13 while preserving the original workflow graphic.
    d.rectangle((0, 0, 92, 112), fill=BG)
    d.rounded_rectangle((-9, 42, 88, 105), radius=12, fill=TEAL)
    tw, th = text_size(d, "13", font(34, True))
    d.text((40 - tw / 2, 54), "13", font=font(34, True), fill="white")
    out = WORK / "new_slide13.png"
    img.save(out)
    return out


def slide14() -> Path:
    img = make_background()
    d = ImageDraw.Draw(img)
    draw_page_header(d, "14", "总结：面向人机协同的肿瘤学证据工作流")

    summary_items = [
        ("01", "不是替代医生，而是整理证据", "doctor"),
        ("02", "不是单次回答，而是可追踪工作流", "workflow"),
        ("03", "不是越自动越好，而是越可复核越安全", "shield"),
    ]
    for i, (num, text, kind) in enumerate(summary_items):
        y = 190 + i * 180
        d.rounded_rectangle((96, y, 1018, y + 130), radius=18, outline=TEAL_LIGHT, width=2, fill=(255, 255, 255))
        d.ellipse((124, y + 26, 204, y + 106), fill=TEAL)
        tw, _ = text_size(d, num, font(35, True))
        d.text((164 - tw / 2, y + 45), num, font=font(35, True), fill="white")
        if kind == "doctor":
            d.ellipse((274, y + 35, 314, y + 75), outline=TEAL, width=4)
            d.line((294, y + 75, 294, y + 100), fill=TEAL, width=4)
            d.arc((256, y + 78, 332, y + 134), 205, 335, fill=TEAL, width=4)
            draw_check(d, 337, y + 76, 22)
        elif kind == "workflow":
            for j, xx in enumerate([260, 310, 360]):
                d.rounded_rectangle((xx, y + 65, xx + 35, y + 100), radius=6, outline=TEAL, width=4)
                if j:
                    d.line((xx - 15, y + 82, xx, y + 82), fill=TEAL, width=4)
            d.rounded_rectangle((309, y + 32, 354, y + 58), radius=8, outline=TEAL, width=4)
            draw_check(d, 389, y + 82, 22)
        else:
            d.polygon([(300, y + 28), (355, y + 50), (345, y + 106), (300, y + 123), (255, y + 106), (245, y + 50)], outline=TEAL, fill=(235, 250, 252), width=4)
            d.line((276, y + 76, 294, y + 95, 326, y + 55), fill=TEAL, width=6, joint="curve")
        d.line((390, y + 34, 390, y + 96), fill=LINE, width=2)
        d.text((420, y + 45), text, font=font(36, True), fill=NAVY)

    # Right evidence chain.
    d.rounded_rectangle((1050, 184, 1560, 700), radius=20, outline=TEAL_LIGHT, width=2, fill=(255, 255, 255))
    d.text((1196, 218), "肿瘤学证据链条", font=font(28, True), fill=NAVY)
    d.line((1105, 253, 1180, 253), fill=TEAL, width=2)
    d.line((1430, 253, 1506, 253), fill=TEAL, width=2)
    cx, cy = 1305, 457
    chain = [
        ("证据检索", 1305, 315, "book"),
        ("证据摘要", 1450, 400, "cite"),
        ("试验匹配与排序", 1412, 590, "trial"),
        ("证据解读", 1196, 590, "split"),
        ("人工复核与决策", 1156, 400, "shield"),
    ]
    for _, x, y, _ in chain:
        d.line((cx, cy, x, y), fill=(97, 185, 195), width=2)
    d.rounded_rectangle((1251, 403, 1359, 511), radius=20, outline=TEAL, width=4, fill=(244, 251, 252))
    d.text((1278, 435), "AI", font=font(37, True), fill=TEAL_DARK)
    for label, x, y, kind in chain:
        draw_simple_icon(d, kind if kind != "shield" else "warn", x, y)
        tw, _ = text_size(d, label, font(18, True))
        d.text((x - tw / 2, y + 49), label, font=font(18, True), fill=TEXT)

    d.rounded_rectangle((174, 765, 1460, 850), radius=35, fill=TEAL)
    d.text((305, 788), "近期重点：Evidence Brief / Trial Pre-screening / MDT Preparation", font=font(33, True), fill="white")
    out = WORK / "new_slide14.png"
    img.save(out)
    return out


def future_card(
    d: ImageDraw.ImageDraw,
    x: int,
    y: int,
    w: int,
    label: str,
    title: str,
    bullets: list[str],
    metrics: str,
):
    d.rounded_rectangle((x, y, x + w, y + 500), radius=18, outline=TEAL_LIGHT, width=2, fill=(255, 255, 255))
    d.ellipse((x + 34, y + 28, x + 102, y + 96), fill=TEAL)
    tw, th = text_size(d, label, font(24, True))
    d.text((x + 68 - tw / 2, y + 47), label, font=font(24, True), fill="white")
    d.text((x + 125, y + 37), title, font=font(30, True), fill=NAVY)
    d.line((x + 34, y + 122, x + w - 34, y + 122), fill=LINE, width=2)
    yy = y + 155
    for b in bullets:
        d.ellipse((x + 40, yy + 8, x + 51, yy + 19), fill=TEAL)
        yy = draw_wrapped(d, (x + 68, yy), b, font(22), fill=TEXT, max_width=w - 100, line_gap=7) + 10
    d.rounded_rectangle((x + 34, y + 382, x + w - 34, y + 462), radius=13, outline=(210, 232, 238), width=1, fill=(242, 250, 252))
    d.text((x + 58, y + 397), "评价重点", font=font(18, True), fill=TEAL_DARK)
    draw_wrapped(d, (x + 58, y + 424), metrics, font(18), fill=MUTED, max_width=w - 110, line_gap=5)


def slide15() -> Path:
    img = make_background()
    d = ImageDraw.Draw(img)
    draw_page_header(d, "15", "未来展望：从单点工具走向临床级协作系统", "三个面向研究与临床转化的同级目标")

    cards = [
        ("01", "多模态协同与前瞻验证", ["指南、病理、影像、分子信息", "病程文本进入同一证据链"], "跨模态冲突如何处理？"),
        ("02", "多角色 MDT 协作智能体", ["内科、外科、放疗、病理", "影像、分子、试验匹配分工"], "角色意见如何对齐？"),
        ("03", "跨中心部署与持续治理", ["不同医院和数据系统下可用", "持续监测漂移、公平性与审计"], "稳定性和责任边界如何保证？"),
    ]
    for i, (num, title, bullets, question) in enumerate(cards):
        x = 108 + i * 502
        d.rounded_rectangle((x, 242, x + 456, 708), radius=22, outline=TEAL_LIGHT, width=2, fill=(255, 255, 255))
        d.ellipse((x + 34, 270, x + 102, 338), fill=TEAL)
        tw, _ = text_size(d, num, font(24, True))
        d.text((x + 68 - tw / 2, 289), num, font=font(24, True), fill="white")
        title_font = font(25 if len(title) > 11 else 28, True)
        draw_wrapped(d, (x + 126, 278), title, title_font, fill=NAVY, max_width=335, line_gap=6)
        d.line((x + 36, 365, x + 420, 365), fill=LINE, width=2)
        yy = 408
        for b in bullets:
            d.ellipse((x + 45, yy + 10, x + 57, yy + 22), fill=TEAL)
            d.text((x + 76, yy), b, font=font(22, True), fill=TEXT)
            yy += 42
        d.rounded_rectangle((x + 36, 582, x + 420, 662), radius=15, outline=(198, 225, 232), width=1, fill=(242, 250, 252))
        d.text((x + 60, 599), "关键问题", font=font(18, True), fill=TEAL_DARK)
        draw_wrapped(d, (x + 60, 625), question, font(19, True), fill=MUTED, max_width=340, line_gap=5)

    d.rounded_rectangle((130, 773, 1540, 850), radius=24, fill=TEAL)
    d.text((235, 795), "未来重点不是更会回答，而是可验证、可协作、可治理的肿瘤学证据系统。",
           font=font(29, True), fill="white")

    source = "Sources: oncology AI agent, Nat Cancer 2025; HONeYBEE, npj Digit Med 2025; TrialGPT, Nat Commun 2024; prospective evaluation principles"
    d.text((142, 898), source, font=font(15), fill=(93, 111, 137))

    out = WORK / "new_slide15.png"
    img.save(out)
    return out


def extract_speaker_notes() -> dict[int, str]:
    text = HANDOUT.read_text(encoding="utf-8")
    notes = {}
    for page in [12, 13, 14, 15]:
        m = re.search(
            rf"## 第 {page} 页：.*?### 讲述文案\s*(.*?)\s*### 数据来源",
            text,
            flags=re.S,
        )
        if not m:
            raise RuntimeError(f"Could not find speaker notes for page {page}")
        notes[page] = re.sub(r"\n{3,}", "\n\n", m.group(1).strip())
    return notes


def image_only_slide_xml(rel_id: str, name: str = "Full-slide image") -> bytes:
    xml = f'''<?xml version="1.0" encoding="utf-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:srgbClr val="FFFFFF"/>
        </a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
      </p:grpSpPr>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="2" name="{name}"/>
          <p:cNvPicPr>
            <a:picLocks noChangeAspect="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
          </p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="{rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
          <a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        </p:blipFill>
        <p:spPr>
          <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:off x="0" y="0"/>
            <a:ext cx="12192000" cy="6858000"/>
          </a:xfrm>
          <a:prstGeom prst="rect" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>'''
    return xml.encode("utf-8")


def replace_note_text(data: bytes, note: str) -> bytes:
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    root = etree.fromstring(data)
    nodes = root.xpath(".//a:t", namespaces=ns)
    if not nodes:
        return data
    for node in nodes:
        node.text = ""
    nodes[0].text = note
    nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=False)


def update_pptx():
    slide12()
    slide14()
    slide15()

    notes = extract_speaker_notes()
    replacements = {
        "ppt/media/image12.png": (WORK / "new_slide12.png").read_bytes(),
        "ppt/media/image15.png": (WORK / "new_slide14.png").read_bytes(),
        "ppt/media/image14.png": (WORK / "new_slide15.png").read_bytes(),
        "ppt/slides/slide14.xml": image_only_slide_xml("R3a1a618dfefa44f6", "Future outlook slide"),
    }
    note_replacements = {
        "ppt/notesSlides/notesSlide12.xml": notes[12],
        # Presentation order: actual slide 14 is slide15.xml/notesSlide15.xml.
        "ppt/notesSlides/notesSlide15.xml": notes[14],
        # Presentation order: actual slide 15 is slide14.xml/notesSlide14.xml.
        "ppt/notesSlides/notesSlide14.xml": notes[15],
    }

    backup = OUT / "final_12_15_adjusted_before_regen_12_14_15.pptx"
    if not backup.exists():
        shutil.copy2(PPTX, backup)

    tmp = PPTX.with_suffix(".tmp.pptx")
    with ZipFile(PPTX, "r") as zin, ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in replacements:
                data = replacements[item.filename]
            elif item.filename in note_replacements:
                data = replace_note_text(data, note_replacements[item.filename])
            zout.writestr(item, data)
    try:
        tmp.replace(PPTX)
        print("Wrote primary deck:", PPTX)
    except PermissionError:
        alternate = OUT / "final_12_15_regenerated.pptx"
        tmp.replace(alternate)
        print("Primary deck is locked; wrote alternate deck:", alternate)


if __name__ == "__main__":
    WORK.mkdir(parents=True, exist_ok=True)
    update_pptx()
    print("Updated", PPTX)
    for p in ["new_slide12.png", "new_slide13.png", "new_slide14.png", "new_slide15.png"]:
        print(WORK / p)
