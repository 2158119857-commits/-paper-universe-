# -*- coding: utf-8 -*-
"""生成提交用的实验报告与 AI 交互记录（docx）"""
import json, os, re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"D:\工作\一页纸的宇宙"
OUT = os.path.join(ROOT, "提交材料")
SHOTS = os.path.join(OUT, "截图")
os.makedirs(OUT, exist_ok=True)

INK = (0x1a, 0x1a, 0x1a)
GRAY = (0x66, 0x66, 0x66)
SAND = (0x8a, 0x6a, 0x3a)


def font(run, cn="宋体", en="Times New Roman", size=12, bold=False, color=INK):
    run.font.name = en
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), cn)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color)


def para(doc, text="", size=12, bold=False, cn="宋体", color=INK, align=None,
         first_indent=True, space_after=6, line=1.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line
    if first_indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    if align is not None:
        p.alignment = align
    if text:
        font(p.add_run(text), cn=cn, size=size, bold=bold, color=color)
    return p


def h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(10)
    font(p.add_run(text), cn="黑体", en="Arial", size=16, bold=True)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    font(p.add_run(text), cn="黑体", en="Arial", size=13, bold=True)
    return p


def bullet(doc, text, size=11.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.4
    font(p.add_run(text), size=size)
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    font(p.add_run(text), cn="宋体", size=9, color=GRAY)


def picture(doc, name, width=15.2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(os.path.join(SHOTS, name), width=Cm(width))


def make_table(doc, headers, rows, widths=None, size=9.5):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        font(p.add_run(h), cn="黑体", en="Arial", size=size, bold=True)
        shade(hdr[i], "EFE9DF")
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            font(p.add_run(str(v)), size=size)
    if widths:
        for r in t.rows:
            for i, w in enumerate(widths):
                r.cells[i].width = Cm(w)
    return t


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def setup(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.6)
    sec.right_margin = Cm(2.6)
    sec.top_margin = Cm(2.4)
    sec.bottom_margin = Cm(2.2)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")


# ------------------------------------------------------------------ 报告
def make_report():
    doc = Document()
    setup(doc)

    # 封面
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run("（此处粘贴学校 LOGO）"), cn="宋体", size=10, color=GRAY)
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run("《解密DeekSeek：从基础到创意实践》"), cn="黑体", en="Arial", size=18, bold=True)
    for _ in range(2):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run("实  验  报  告"), cn="黑体", en="Arial", size=34, bold=True)
    for _ in range(4):
        doc.add_paragraph()

    info = [
        ("实验名称", "一页纸的宇宙 —— 基于 WebGL 的尺度旅行交互网站"),
        ("学    号", "____________________"),
        ("姓    名", "____________________"),
        ("班    级", "____________________"),
        ("日    期", "2026 年 10 月"),
    ]
    for k, v in info:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(10)
        font(p.add_run(k + "："), cn="黑体", en="Arial", size=13, bold=True)
        font(p.add_run(v), size=13)
    doc.add_page_break()

    # 一、项目概述
    h1(doc, "一、项目概述")
    para(doc, "作品名称：一页纸的宇宙（Powers of Ten）。作品形式：单页 WebGL 交互网站，已部署至 GitHub Pages，"
              "无需安装即可在浏览器中完整访问。在线地址：https://2158119857-commits.github.io/-paper-universe-/。")

    h2(doc, "1. 选题背景")
    para(doc, "本作品源于课程中的「个人网站开发与部署」案例。最初的目标只是照着案例做一个可以上线的个人主页，"
              "但在讨论中我意识到：个人网站最容易变成模板的堆叠，真正难的是找到一个只有我能做、并且值得别人停留十秒的表达。"
              "我选择了「尺度」这个主题——人类对数字有概念，对尺度却没有直觉。10²⁶ 米这个数字毫无意义，"
              "只有连续地退后、亲眼看着沙粒变成星系，尺度才会变成一种体感。")
    para(doc, "灵感来自两个方向：一是经典的科普影片《Powers of Ten》，它用连续缩放讲清了宇宙的层级；"
              "二是我在短视频平台上看到的一个 AI 生成的三维星系演示，视觉效果很强，但它只是「好看」，没有叙事。"
              "我想把两者合起来：保留《Powers of Ten》的线性尺度旅程，同时让其中几个节点可以真正走进去、转动视角。")

    h2(doc, "2. 目标观众与使用场景")
    para(doc, "目标观众是对科学可视化、宇宙题材感兴趣的普通观众，也包括课堂演示与科普展览场景。"
              "作品不要求任何前置知识：不需要按钮教学，滚轮或拖动就能前进；每一站都有 100 字以内的说明与真实出处；"
              "愿意深入的观众可以点击热点、进入三维站台，或者复制某个尺度的链接分享给别人。")

    h2(doc, "3. 主题与作品价值")
    para(doc, "作品想表达的主题是：所有我们熟悉的东西——沙粒、城市、海岸、地球——都只是同一段尺度轴上的一个切片。"
              "结尾落在「你刚才，是从一粒沙出发的」，把宏大的数字重新拉回到个人经验。"
              "在价值上，它既是一次前端技术练习（无框架手写 WebGL 引擎），也是一次 AI 协作方法的实践："
              "如何把审美判断翻译成可执行的技术要求，如何用截图和真机验收每一步。")

    # 二、项目目标与方案设计
    h1(doc, "二、项目目标与方案设计")

    h2(doc, "1. 项目目标")
    bullet(doc, "把 11 张真实影像串成一条连续的尺度轴，覆盖 10⁻² 米到 10²⁶·⁵ 米，共 28 个数量级。")
    bullet(doc, "在尺度轴上设置三个可以「走进去」的三维站台：地球轨道、太阳系、银河系。")
    bullet(doc, "让静态照片具备三维感：用深度图为每张照片生成曲面，随指针产生真实视差与光照变化。")
    bullet(doc, "为跨度特别大的段落设计独立的视觉叙事，而不是简单切黑场。")
    bullet(doc, "保证手机可打开、可操作，并提供可分享的尺度深链。")

    h2(doc, "2. 核心功能与模块结构")
    para(doc, "作品是一个单文件应用（index.html 约 3100 行，140 KB），所有 HTML、CSS、JavaScript 与着色器内联，"
              "不依赖任何第三方库或 CDN。逻辑上分为四层：")
    bullet(doc, "数据层：STAGES 数组描述 11 站的尺度、影像、文案、热点；SOLAR 数组描述八大行星的轨道与外观。")
    bullet(doc, "渲染层：自写透视投影与 lookAt 矩阵、球体/圆环/网格几何、7 组 GLSL 着色器；"
                "影像平面使用 96×54 网格，顶点按深度图抬升。")
    bullet(doc, "交互层：滚轮速度累积、指针拖拽、方向键逐站、空格播放、点击热点与天体、进入/退出站台。")
    bullet(doc, "界面层：顶部品牌与帮助、右侧站点导轨、热点标注、底部信息面板、跨尺度对数尺、信息抽屉。")

    h2(doc, "3. 界面风格")
    para(doc, "整体是深空黑底色，文字用接近宣纸的暖白，强调色取沙色。中文标题用宋体，数字与英文用等宽字体，"
              "让「科学仪器」和「一本书」两种气质同时存在。所有面板都不做重边框，只靠留白与极细的分隔线，"
              "避免遮挡影像本身。")

    h2(doc, "4. 主要操作流程")
    para(doc, "开场页 → 点击「开始旅程」→ 滚轮或拖拽控制尺度 → 每站可点击热点查看细节 → 在地球、太阳、银河三站出现"
              "「进入轨道 / 进入太阳系 / 进入银河系」入口 → 三维站台内拖动环绕、滚轮缩放、点击天体 → 返回尺度轴 → "
              "在结尾选择「回到一粒沙」或「留在这里」。")

    # 三、制作工具与实施过程
    h1(doc, "三、制作工具与实施过程")

    h2(doc, "1. 使用的工具")
    make_table(doc,
               ["工具 / 平台", "用途"],
               [
                   ["Cherry Studio + Claude", "对话式编程：需求拆解、WebGL 引擎与着色器编写、调试、文档撰写"],
                   ["Depth Anything V2（small, ONNX）", "本地单目深度估计，为 15 张影像生成深度图"],
                   ["Python + uv + onnxruntime + Pillow", "批量推理与深度图后处理（tools/make_depth.py）"],
                   ["Bun", "本地静态服务器、脚本运行与语法检查"],
                   ["无头 Edge + CDP", "自动化截图与视觉验收（tools/capture.js）"],
                   ["Git + GitHub Pages", "版本管理与线上部署"],
                   ["NASA / ESA 公共影像库", "全部站点影像素材（公共领域）"],
               ], widths=[6.2, 9.5])

    h2(doc, "2. 实施步骤")
    para(doc, "整个项目从 10 月 21 日上午开始，到 10 月 27 日完成，大致分为七个阶段：")
    bullet(doc, "（1）需求与概念：分析课程案例，确定「尺度旅程」的方向，先做出一版 SVG 插画风的概念稿。")
    bullet(doc, "（2）第一次推翻：概念稿被评价为「太卡通、可互动性弱」，于是改用真实 NASA 影像，并引入 WebGL。")
    bullet(doc, "（3）三维化：实现透视投影、球体与粒子系统，加入地球轨道、太阳系、银河系三个站台。")
    bullet(doc, "（4）可用性修复：热点跟随、移动端黑屏、贴图按需加载、降分辨率。")
    bullet(doc, "（5）部署：创建 GitHub 仓库、配置 Pages、处理 HTML 缓存导致的「旧版本」问题。")
    bullet(doc, "（6）转场与叙事：解决切图时透出星空的问题，为跨数量级段落设计推镜、对数尺、量级水印与速度线。")
    bullet(doc, "（7）深度与动态：用深度估计把静态照片变成三维曲面，加入方向光、云层漂移、星野慢转与灯火呼吸。")

    h2(doc, "3. 本人工作与 AI 辅助的划分")
    para(doc, "本人承担的工作：选题与创意方向、对每一版效果的判断与取舍（例如否定插画风、要求保留线性尺度轴）、"
              "需求拆解与验收（在手机和桌面浏览器上实际打开检查）、GitHub 账号与部署授权、全部文案的事实核对。")
    para(doc, "AI 辅助完成的工作：WebGL 引擎与着色器实现、深度图脚本、布局与动效、调试（通过报错与截图定位问题）、"
              "部署命令与文档初稿。整个过程中 AI 也出现过错误判断（例如认为「切图露星空」只是透明度过低），"
              "最后是通过我提供的截图和「还是星空」的反馈才定位到黑底板方案。")

    # 四、成果展示
    h1(doc, "四、作品成果展示")
    para(doc, "作品已部署至 GitHub Pages，可直接访问：https://2158119857-commits.github.io/-paper-universe-/。"
              "仓库地址：https://github.com/2158119857-commits/-paper-universe-。")
    para(doc, "以下截图均来自线上版本（1280×800，无头浏览器实拍）：")

    shots = [
        ("01-cover.jpg", "图 1  开场页：标题、尺度提示与「开始旅程」"),
        ("02-sand.jpg", "图 2  第一站「一粒沙」：照片被深度图抬成曲面，沙粒之间有真实高差"),
        ("03-dunes.jpg", "图 3  「沙丘与海岸」：方向光让沙脊显出明暗，指针移动时高光会变"),
        ("04-city.jpg", "图 4  「城市夜色」：灯火随呼吸微微闪烁，深度视差让城市有厚度"),
        ("05-earth.jpg", "图 5  「地球」：蓝色弹珠影像，云层相对地面缓慢漂移"),
        ("06-earth-orbit.jpg", "图 6  进入轨道：实时昼夜光照的地球、独立云层、月球与杭州定位点"),
        ("07-solar.jpg", "图 7  太阳系站台：行星公转、土星环、底部天体导航"),
        ("08-galaxy.jpg", "图 8  银河系站台：粒子旋臂与「太阳系在这里」标记"),
        ("09-crossing.jpg", "图 9  跨越 9.6 个数量级：推镜、对数尺、量级水印与速度线"),
        ("10-ending.jpg", "图 10  旅程结尾：把宏大的尺度重新拉回个人经验"),
    ]
    for name, cap in shots:
        picture(doc, name)
        caption(doc, cap)

    h2(doc, "使用流程")
    para(doc, "打开链接后点击「开始旅程」，向上滚动（或按住拖动）即开始后退；右侧导轨可直接跳到任意一站；"
              "点击画面中的黄色热点查看细节；在地球、太阳、银河三站点击左下角入口进入三维站台，Esc 或「返回尺度轴」退出；"
              "点击「复制链接」会得到携带当前尺度的地址，例如 #u=14.2 表示 10¹⁴·² 米。")

    # 五、总结
    h1(doc, "五、项目总结与学习体会")

    h2(doc, "1. 完成情况")
    para(doc, "项目目标基本完成：11 个尺度站点全部上线，3 个可进入的三维站台可用，静态影像实现了深度曲面与动态效果，"
              "跨尺度段落有独立的视觉叙事，手机端可正常打开与操作，作品已部署为公开链接。")

    h2(doc, "2. 不足与改进方向")
    bullet(doc, "单张照片只能还原「浮雕式」三维，无法绕到物体背面；如果要真正的三维，需要多视角采集或生成式重建。")
    bullet(doc, "深空影像的深度是模型推断的，星空部分的三维感有限，局部会出现不自然的起伏。")
    bullet(doc, "银河系一站是艺术想象图，严格来说不是实拍，后续可以换成可交互的观测数据可视化。")
    bullet(doc, "低端移动设备在开启全部效果时可能掉帧，还可以进一步做分级画质。")

    h2(doc, "3. 对 AI 能力与局限的认识")
    para(doc, "这次协作让我最直接的感受是：AI 能把「想法」很快变成「可运行的东西」，但它不会替你做审美判断。"
              "同一个要求，它可以给出十种实现，哪一种「对」需要人来定。它也常常在细节上过于自信，"
              "例如第一次处理转场时，它只调整了透明度，直到我把「还是星空」的截图发过去，才真正找到黑底板的解法。"
              "所以人的角色不是「提需求然后验收」，而是持续地看、持续地指出哪里不对。")

    h2(doc, "4. 提示词与质量判断的收获")
    para(doc, "有效的提示词往往不是更长，而是更具体、可验证：与其说「做得高级一点」，不如说"
              "「切图时不要露出星空」「黄点必须跟着图像移动」「手机打开不能只有星星」。"
              "另外，我学会了用截图和真机作为验收标准，而不是只听描述；也学会了把一次大改拆成几个可回滚的提交，"
              "这样即使某一版更差，也能退回去。")

    # 补充材料
    h1(doc, "补充材料：AI 协作与迭代改进")
    para(doc, "完整交互记录见同目录下的《AI 交互记录.docx》。下面列出 8 个关键迭代节点：")
    make_table(doc,
               ["轮次", "发现的问题", "向 AI 提出的要求（节选）", "AI 的处理", "我的判断与结果"],
               [
                   ["1", "插画风太卡通，只有一个页面，互动弱",
                    "「能不能更精细一点 现在有点卡通了 有没有真实的……可互动性差了一点」",
                    "改用真实 NASA 影像，引入 WebGL 透视相机与影像平面",
                    "方向正确，保留；但仍是平面，继续要求 3D"],
                   ["2", "缺少空间感",
                    "「能不能做成3d的」",
                    "实现球体、粒子星空、深度缩放",
                    "有空间感了，但还缺少可进入的场景"],
                   ["3", "希望参考短视频里的星系演示，但不想放弃线性叙事",
                    "「你有没有办法 做成和他一样的……但是我不想放弃本身的形式」",
                    "提出「线性尺度轴 + 三个可进入站台」的方案",
                    "采纳，成为作品的最终结构"],
                   ["4", "热点黄点不跟随图像，进站后消失",
                    "「有一个问题 黄色小点不会跟着图像移动 这样很突兀」「还是不行 有的点都没了」",
                    "把热点从屏幕坐标改为投影到当前影像平面，并在 3D 模式隐藏",
                    "采纳；连续三轮修复后通过"],
                   ["5", "手机打开只有星空",
                    "（截图反馈）",
                    "按需加载 ±2 站贴图、移动端降分辨率、加入载入提示",
                    "采纳，手机恢复正常"],
                   ["6", "切图时中间透出星空",
                    "「怎么让一个图片切到另一个图片的时候更加丝滑 而不是还是星空」",
                    "先只调透明度，未解决；后改为满屏叠化 + 黑色底板 + 未就绪时托住上一张图",
                    "第二次方案有效，采纳"],
                   ["7", "跨尺度段落不好看、读数不清",
                    "「然后中间的跨越尺度有没有更加好看的方法」「感觉不够清晰 跨越尺度」",
                    "推镜 + 对数尺 + 量级水印 + 速度线，并放大读数",
                    "采纳；成为作品最有辨识度的一段"],
                   ["8", "静态照片仍然像照片",
                    "「让每一个静态图都成为动态的或者3d的可以吗」「其他的能不能也改成3d的 我们可以突破一下」",
                    "用 Depth Anything V2 生成深度图，网格抬升 + 指针旋转 + 环境光遮蔽；再加强方向光与云层/星野动态",
                    "采纳；完成从 2.5D 到三维曲面的升级"],
               ], widths=[1.1, 3.3, 4.4, 4.0, 3.2], size=8.5)

    doc.save(os.path.join(OUT, "实验报告.docx"))
    print("saved report")


# ------------------------------------------------------------ 交互记录
def clean(s):
    s = re.sub(r"github_pat_[A-Za-z0-9_]+", "github_pat_（已隐去）", s)
    s = s.replace("Attached files (read them with your tools using these absolute paths):", "（附文件）")
    s = re.sub(r"\n\s*-\s*\"[^\"]+\":\s*[^\n]+", "", s)
    s = s.strip()
    if s.startswith("<task-notification>"):
        return ""
    if s.startswith("github_pat_"):
        return "（提交 GitHub 个人访问令牌，内容已隐去）"
    return s


def make_log():
    doc = Document()
    setup(doc)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(60)
    font(p.add_run("《一页纸的宇宙》"), cn="黑体", en="Arial", size=24, bold=True)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run("AI 交互记录"), cn="黑体", en="Arial", size=20, bold=True)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24)
    font(p.add_run("记录时间：2026-10-21 10:26 — 2026-10-27 09:40"), size=11, color=GRAY)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run("工具：Cherry Studio · Claude（对话式编程）"), size=11, color=GRAY)
    doc.add_page_break()

    h1(doc, "一、记录说明")
    para(doc, "本记录来自本次作业的真实会话，按时间顺序整理。整理后共 50 条有效用户输入、216 段可读消息"
              "（原始日志含系统通知与重复记录，已剔除），时间跨度为 2026 年 10 月 21 日至 10 月 27 日。"
              "正文标注了 10 个关键迭代节点，附录 A 列出全部有效输入（隐去敏感信息），"
              "附录 B 为每条输入对应的 AI 回复要点。所有效果均以线上版本与真机截图为准。")

    h1(doc, "二、关键迭代节点")

    nodes = [
        ("节点 1：从「做个个人网站」到「做一段尺度旅程」",
         "分析完课程案例后，我确认代码不是难点，难点是选题。我让 AI 列出更有创意的方向，"
         "在若干个方案里选中了第 5 个——用连续缩放讲尺度。原始需求是：",
         "「有没有更有意思或创意的 给我列下来几个 让我选择一下」",
         "AI 给出了 6 个方向（尺度旅程、时间胶囊、声音地图等），并说明了每个方案的实现成本。"
         "我选择了尺度旅程，因为它有明确的视觉主线，而且素材可以全部使用真实的 NASA 影像。",
         "判断：这个方向的价值在于「连续」——不是一页页翻卡片，而是一条可以从头滚到尾的轴。"),

        ("节点 2：推翻插画风，改用真实影像",
         "第一版是 SVG 插画风格，颜色干净但明显「卡通」。我给出的反馈是：",
         "「能不能更精细一点 现在有点卡通了 有没有真实的 而且我觉得可互动性差了一点 只是一个页面 感觉有点简陋」",
         "AI 把 11 张站点图换成真实 NASA 影像，引入 WebGL 透视相机，影像不再是一张铺满屏幕的图片，"
         "而是有景深的平面；同时加入星空粒子层。",
         "判断：真实影像立刻提升了可信度，但平面仍然是平面。这一轮让我意识到「真实」不只是素材，还包括空间。"),

        ("节点 3：做成 3D，并确定「线性轴 + 站台」的结构",
         "我提出能不能做成 3D，并给了一个短视频里的星系演示作为参考，但明确不想放弃原来的线性叙事：",
         "「你有没有办法 做成和他一样的 你先提供点思路 但是我不想放弃本身的形式」",
         "AI 提出：保留 11 站的线性尺度轴，只在三个关键节点（地球、太阳、银河）提供「进入」入口，"
         "进入后是真正的三维场景，退出后回到原来的尺度位置。",
         "判断：这是整个作品最重要的结构决定——视频里的演示是「展厅」，我们做的是「旅程 + 三个可以走进去的房间」。"),

        ("节点 4：热点必须跟着图像走",
         "三维站台做好后，画面上的黄色热点不会跟随影像移动，进入轨道模式后还会残留：",
         "「有一个问题 黄色小点不会跟着图像移动 这样很突兀」「还是不行 有的点都没了」",
         "前两次修复只改了 CSS 定位，没有解决根因；第三次把热点从屏幕坐标改为每帧投影到当前影像平面，"
         "并在三维模式与转场期间隐藏，才彻底通过。",
         "判断：这是一个典型的「表面现象 vs 根因」案例。我坚持用截图指出具体哪一站、哪个点不对，才让问题收敛。"),

        ("节点 5：手机端只有星空",
         "部署后用手机打开，页面只有星空，没有影像。我发了手机截图。",
         "（截图反馈：页面全黑，仅有星点）",
         "AI 判断是首屏一次性加载 6.5 MB 贴图导致移动端内存不足，改为按当前站 ±2 懒加载、移动端贴图上限 1024px，"
         "并加入「载入影像…」提示。",
         "判断：修复有效。这个问题的教训是——桌面端的「能跑」不等于移动端的「能用」。"),

        ("节点 6：切图时中间透出星空",
         "两站之间切换时，画面中间会短暂露出星空，非常出戏：",
         "「怎么让一个图片切到另一个图片的时候更加丝滑 而不是还是星空」",
         "第一次只调整了透明度曲线，没有解决；第二次分析出两个原因：转场时画面上退导致不满屏，"
         "以及两张半透明图叠加后背景仍透出约 25%。最终方案是满屏叠化 + 每张影像背后垫一块同尺寸黑底板 + "
         "新图未就绪时托住上一张图。",
         "判断：AI 第一次的判断是错的，是我反复用「还是星空」的截图把它拉回正确方向。"),

        ("节点 7：为跨尺度段落设计叙事",
         "相邻站点之间的过渡修好后，跨度特别大的两段（太阳→恒星 9.6 个数量级、银河→宇宙网 3 个数量级）仍然只是「淡出淡入」：",
         "「然后中间的跨越尺度有没有更加好看的方法」→「按你的倾向做 发挥你自己的最大想象」",
         "AI 实现了「推镜穿越」：旧图放大送出、新图从 1.16 倍收回；顶部对数尺随尺度滑动；中央巨号量级水印逐级跳动；"
         "画面中心向外飞出速度线。",
         "判断：这一段后来成为作品最有辨识度的部分。随后我指出读数不够清楚，AI 又放大了刻度、提高了对比度，"
         "并加了「每一格十倍」的说明。"),

        ("节点 8：把静态照片变成三维曲面",
         "我提出一个更大的想法：除了三个真正的 3D 站台，其他照片能不能也变成动态或 3D 的。",
         "「让每一个静态图都成为动态的或者3d的可以吗」「除了那三张真正的3D场景 其他的能不能也改成3d的 我们可以突破一下」",
         "AI 先用单目深度估计模型（Depth Anything V2）离线为每张照片生成深度图，再把影像平面换成 96×54 网格，"
         "顶点按深度抬起；指针带动曲面轻微旋转，并用深度邻域计算环境光遮蔽。",
         "判断：第一次实现时着色器因为精度声明不一致而链接失败，画面只剩巨大的三角形；"
         "通过检查 LINK_STATUS 定位并修复后，沙丘与沙粒的立体感非常明显。"),

        ("节点 9：让画面真正动起来",
         "我要求「两个都做」——既更真实，也更有动态。",
         "「两个都做 可以的」",
         "AI 把起伏幅度提高 40%，加入会随指针偏转的方向光；同时实现云层相对地面漂移（只作用于白而不鲜艳的像素）、"
         "星野极慢旋转、灯火呼吸。",
         "判断：动态幅度控制在「看得见但不喧宾夺主」，避免变成廉价的闪烁特效。"),

        ("节点 10：部署与交付",
         "作品完成后需要上线。过程中遇到 GitHub 令牌权限不足、仓库改名、GitHub Pages 的 10 分钟 HTML 缓存等问题：",
         "「过一会那个网址会更新吗」「没有暗下来啊」",
         "AI 解释了 Pages 的缓存策略，并给出带 ?v= 参数的访问方式用于强刷；最终作品稳定运行在 GitHub Pages 上。",
         "判断：线上验证必须用带参数的地址，否则看到的可能是十分钟前的版本。"),
    ]
    for title, problem, prompt, output, judge in nodes:
        h2(doc, title)
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
        font(p.add_run("问题："), cn="黑体", en="Arial", size=11, bold=True)
        font(p.add_run(problem), size=11)
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Cm(0.6)
        font(p.add_run(prompt), cn="楷体", size=11, color=SAND)
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
        font(p.add_run("AI 的处理："), cn="黑体", en="Arial", size=11, bold=True)
        font(p.add_run(output), size=11)
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(8)
        font(p.add_run("结果与反思："), cn="黑体", en="Arial", size=11, bold=True)
        font(p.add_run(judge.replace("判断：", "", 1)), size=11)

    h1(doc, "三、版本与迭代统计")
    make_table(doc,
               ["提交", "时间", "内容"],
               [
                   ["6cc5128", "10-21 22:39", "尺度旅程 + 地球/太阳系/银河 3D 站台"],
                   ["696a8f1", "10-21 23:23", "按需加载贴图、移动端降分辨率、载入提示"],
                   ["9ede84f", "10-21 23:34", "加入旅程结尾与地球云层"],
                   ["1fa91a9", "10-21 23:44", "转场垫黑底板并托住上一张图"],
                   ["a3efa57", "10-21 23:53", "跨尺度段落改为满屏推镜穿越"],
                   ["ce820fa", "10-22 00:07", "每张影像接入深度图"],
                   ["0dbd97f", "10-22 00:19", "静态影像升级为真正的三维曲面"],
                   ["67054ce", "10-27 17:34", "加强三维起伏与方向光；云层漂移、星野慢转"],
               ], widths=[2.4, 3.2, 9.8], size=9.5)

    # 附录
    data = json.load(open(os.path.join(ROOT, "tools", "log_timeline.json"), encoding="utf-8"))
    turns = []
    prev = None
    for e in data:
        if e["role"] == "user":
            t = clean(e["text"])
            if not t or t == prev:
                continue
            if "system-reminder" in t or t.startswith("[Image:"):
                continue
            if "This session is being continued" in t:
                continue
            prev = t
            turns.append({"ts": e["ts"], "user": t, "ai": ""})
        elif e["role"] == "assistant" and turns and not turns[-1]["ai"]:
            turns[-1]["ai"] = clean(e["text"])

    h1(doc, "附录 A：完整用户输入")
    para(doc, "以下为全部用户输入（去重、隐去令牌与附件路径），按时间排序。", first_indent=False, size=10.5)
    for i, t in enumerate(turns, 1):
        ts = t["ts"].replace("T", " ")[:16]
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.left_indent = Cm(0.4)
        font(p.add_run(f"{i:02d}. [{ts}] "), cn="宋体", size=9.5, color=GRAY)
        font(p.add_run(t["user"].replace("\n", " ")[:220]), size=9.5)

    doc.add_page_break()
    h1(doc, "附录 B：AI 回复要点")
    para(doc, "每条用户输入对应的 AI 首次回复摘要（取前 90 字），用于呈现完整的往复过程。",
         first_indent=False, size=10.5)
    for i, t in enumerate(turns, 1):
        if not t["ai"]:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.4)
        font(p.add_run(f"{i:02d}. "), cn="宋体", size=9, color=GRAY)
        font(p.add_run(t["ai"].replace("\n", " ")[:90]), size=9)

    doc.save(os.path.join(OUT, "AI交互记录.docx"))
    print("saved log, turns:", len(turns))


make_report()
make_log()
