const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, 
        ShadingType, LevelFormat, PageNumber, PageBreak, VerticalAlign } = require('docx');

// ========== Shared Helpers ==========
const tb = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const cb = { top: tb, bottom: tb, left: tb, right: tb };
const shadeHeader = { fill: "1A1A2E", type: ShadingType.CLEAR };
const shadeAlt = { fill: "F8F9FA", type: ShadingType.CLEAR };

function hCell(text, width) {
  return new TableCell({
    borders: cb, width: { size: width, type: WidthType.DXA },
    shading: shadeHeader, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text, bold: true, color: "FFFFFF", size: 20, font: "Arial" })
    ]})]
  });
}

function cell(text, width, opts = {}) {
  return new TableCell({
    borders: cb, width: { size: width, type: WidthType.DXA },
    shading: opts.shade ? shadeAlt : undefined,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ 
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text, size: 20, font: "Arial", bold: opts.bold, color: opts.color })]
    })]
  });
}

function mcell(texts, width, opts = {}) {
  return new TableCell({
    borders: cb, width: { size: width, type: WidthType.DXA },
    shading: opts.shade ? shadeAlt : undefined,
    verticalAlign: VerticalAlign.TOP,
    children: texts.map(t => new Paragraph({ 
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text: t, size: 20, font: "Arial" })]
    }))
  });
}

function heading1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, font: "Arial" })] });
}

function heading2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, font: "Arial" })] });
}

function heading3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 180, after: 60 },
    children: [new TextRun({ text, font: "Arial" })] });
}

function p(text, opts = {}) {
  return new Paragraph({ spacing: { before: 60, after: 60 }, indent: opts.indent ? { left: 360 } : undefined,
    children: [new TextRun({ text, size: 22, font: "Arial", bold: opts.bold, color: opts.color, italics: opts.italic })] });
}

function bullet(text, ref) {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })] });
}

function divider() {
  return new Paragraph({ spacing: { before: 200, after: 200 }, 
    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" } },
    children: [new TextRun({ text: "", size: 4 })] });
}

// ========== Document ==========
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: "1A1A2E", font: "Arial" },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "2D3748", font: "Arial" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: "4A5568", font: "Arial" },
        paragraph: { spacing: { before: 180, after: 60 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "b1", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b2", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b3", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b4", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b5", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b6", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b7", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b8", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b9", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b10", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b11", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "b12", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "n1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "n2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "n3", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [
    // ==================== COVER PAGE ====================
    {
      properties: { 
        page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "CONFIDENTIAL", size: 16, color: "999999", font: "Arial" })] })] })
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "999999" }), 
                     new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" }),
                     new TextRun({ text: " / ", size: 16, font: "Arial", color: "999999" }),
                     new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: "Arial", color: "999999" })] })] })
      },
      children: [
        new Paragraph({ spacing: { before: 2400 } , children: [] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [new TextRun({ text: "腾讯广告", size: 28, color: "718096", font: "Arial" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "H5 获客矩阵", size: 52, bold: true, color: "1A1A2E", font: "Arial" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [new TextRun({ text: "市场需求文档（MRD）", size: 36, bold: true, color: "2D3748", font: "Arial" })] }),
        new Paragraph({ spacing: { before: 600 }, children: [] }),
        
        // Info table
        new Table({
          columnWidths: [2400, 6960],
          rows: [
            new TableRow({ children: [
              cell("项目名称", 2400, { bold: true }), cell("腾讯广告 H5 获客矩阵 — 新客三件套", 6960)
            ]}),
            new TableRow({ children: [
              cell("文档版本", 2400, { bold: true }), cell("v1.0", 6960)
            ]}),
            new TableRow({ children: [
              cell("编制日期", 2400, { bold: true }), cell("2026 年 3 月 19 日", 6960)
            ]}),
            new TableRow({ children: [
              cell("目标受众", 2400, { bold: true }), cell("新客中小广告主（从未投过微信广告）", 6960)
            ]}),
            new TableRow({ children: [
              cell("文档状态", 2400, { bold: true }), cell("Draft — 待评审", 6960)
            ]}),
          ]
        }),

        new Paragraph({ spacing: { before: 800 }, children: [] }),
        divider(),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100 },
          children: [new TextRun({ text: "本文档为内部使用，未经许可不得外传", size: 18, color: "A0AEC0", font: "Arial", italics: true })] }),
      ]
    },

    // ==================== MAIN CONTENT ====================
    {
      properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "腾讯广告 · H5 获客矩阵 MRD", size: 16, color: "999999", font: "Arial" })] })] })
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Page ", size: 16, font: "Arial", color: "999999" }),
                     new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "999999" }),
                     new TextRun({ text: " / ", size: 16, font: "Arial", color: "999999" }),
                     new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: "Arial", color: "999999" })] })] })
      },
      children: [

        // ===== 1. 项目概述 =====
        heading1("1. 项目概述"),

        heading2("1.1 背景与市场机会"),
        p("大量中小商家对微信广告存在认知门槛：以为投广告「很贵」「很难」「需要专业团队」。实际上腾讯广告已大幅降低准入门槛，但信息不对称导致大量潜在客户流失。"),
        p("本项目通过三个独立 H5 互动页面，以「轻游戏化 + 认知纠正」的方式，让新客在 1-3 分钟内完成一次互动体验，自然过渡到获客转化（加企微 / 领行业数据）。"),

        heading2("1.2 项目定位"),
        new Table({
          columnWidths: [2800, 6560],
          rows: [
            new TableRow({ children: [hCell("维度", 2800), hCell("说明", 6560)] }),
            new TableRow({ children: [cell("目标用户", 2800, { bold: true }), cell("从未投过微信广告的中小商家主 / 创业者", 6560)] }),
            new TableRow({ children: [cell("核心策略", 2800, { bold: true }), cell("游戏化互动 → 认知纠正 → 降低心理门槛 → 获客转化", 6560)] }),
            new TableRow({ children: [cell("投放场景", 2800, { bold: true }), cell("朋友圈广告 / 公众号文章 / 视频号信息流 / 社群裂变", 6560)] }),
            new TableRow({ children: [cell("转化目标", 2800, { bold: true }), cell("添加企业微信领取行业数据报告，进入 SDR 跟进链路", 6560)] }),
            new TableRow({ children: [cell("合规约束", 2800, { bold: true }), cell("不涉及效果承诺（无到达率 / 转化率 / ROI 等数据）", 6560)] }),
          ]
        }),

        heading2("1.3 三个 H5 全景总览"),
        new Table({
          columnWidths: [600, 2400, 2400, 1800, 2160],
          rows: [
            new TableRow({ children: [
              hCell("#", 600), hCell("H5 名称", 2400), hCell("核心玩法", 2400), hCell("互动时长", 1800), hCell("转化钩子", 2160)
            ]}),
            new TableRow({ children: [
              cell("1", 600), cell("30s 模拟朋友圈广告", 2400, { bold: true }),
              cell("选行业 → 选目标 → 输名称 → 上传照片（可选）→ 模拟朋友圈广告预览", 2400),
              cell("1-2 min", 1800), cell("「想让这条广告真的出现在朋友圈？」", 2160)
            ]}),
            new TableRow({ children: [
              cell("2", 600), cell("你不知道的微信广告热知识", 2400, { bold: true }),
              cell("6 道选择题 → 纠正新客误解 → 每题有「老炮说」专家解读 → 6个真相总结", 2400),
              cell("2-3 min", 1800), cell("「想在微信投广告？先看看你的行业数据」", 2160)
            ]}),
            new TableRow({ children: [
              cell("3", 600), cell("测测你的微信生意潜力", 2400, { bold: true }),
              cell("5 道维度题 → 多维度分析 → 潜力分评分 + 个性化建议", 2400),
              cell("1-2 min", 1800), cell("「想知道同行都怎么投的？」", 2160)
            ]}),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. H5 #1 朋友圈广告模拟器 =====
        heading1("2. H5 #1 — 30s 模拟朋友圈广告"),

        heading2("2.1 产品概述"),
        p("用户选择行业、投放目标、输入店铺名称后，系统即时生成一条高仿真的朋友圈广告预览。支持「单图模式」和「九宫格模式」两种展示形态，让用户直观感受自己的广告出现在朋友圈是什么效果。"),
        p("核心心理机制：「打不过就加入」——当用户看到自己的广告出现在朋友圈时，会产生强烈的代入感和行动意愿。", { italic: true }),

        heading2("2.2 用户流程"),
        new Table({
          columnWidths: [1200, 2400, 2400, 3360],
          rows: [
            new TableRow({ children: [hCell("步骤", 1200), hCell("页面", 2400), hCell("用户动作", 2400), hCell("系统响应", 3360)] }),
            new TableRow({ children: [
              cell("Step 1", 1200), cell("封面页", 2400), cell("点击「开始生成我的广告」", 2400),
              cell("引导语说明玩法 + 底部标注「内含腾讯广告行业数据」", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 2", 1200), cell("选行业", 2400), cell("从 7 个行业中选择一个", 2400),
              cell("行业选项：餐饮美食 / 丽人美业 / 家装建材 / 教育培训 / 电商零售 / 本地生活 / 其他行业", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 3", 1200), cell("选目标", 2400), cell("选择投放核心目标", 2400),
              cell("目标选项：客户来咨询 / 到店里消费 / 加我微信 / 直接下单", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 4", 1200), cell("填名称 + 上传照片", 2400), cell("输入店铺名（可跳过）、上传照片（可选，最多 9 张）", 2400),
              cell("最多 10 字，有默认占位符 | 不上传则使用模板图", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 5", 1200), cell("广告预览", 2400), cell("查看模拟广告 / 切换单图/九宫格模式", 2400),
              cell("生成仿真朋友圈卡片，含虚拟点赞和评论，支持模式切换", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 6", 1200), cell("CTA 页", 2400), cell("扫码 / 分享", 2400),
              cell("领取腾讯广告行业数据 + 企微二维码", 3360)
            ]}),
          ]
        }),

        heading2("2.3 行业模板配置 & 目标分类"),
        p("每个行业有独立的广告模板（7 个行业），包含：头像 emoji、默认店铺名、4 种投放目标对应的广告文案、主图配色渐变、九宫格 9 个素材配置。"),
        p("投放目标统一为：📞 客户来咨询 / 🚶 到店里消费 / 💬 加我微信 / 🛒 直接下单", { bold: true }),
        new Table({
          columnWidths: [1400, 1800, 2200, 3360],
          rows: [
            new TableRow({ children: [hCell("行业", 1400), hCell("默认名称", 1800), hCell("主图文案", 2200), hCell("示例广告文案（到店目标）", 3360)] }),
            new TableRow({ children: [
              cell("🍵 餐饮美食", 1400), cell("你的餐饮店", 1800), cell("到店立减 · 新客专享", 2200),
              cell("📍 离你 500 米，到店消费满 50 减 15", 3360)
            ]}),
            new TableRow({ children: [
              cell("💇 丽人美业", 1400), cell("你的美业店", 1800), cell("变美不贵 · 限时体验", 2200),
              cell("📍 新店开业，到店体验免费送小样礼包", 3360)
            ]}),
            new TableRow({ children: [
              cell("🏠 家装建材", 1400), cell("你的家装店", 1800), cell("免费设计 · 实景案例", 2200),
              cell("📍 周末样板间开放日，到店送全屋设计方案", 3360)
            ]}),
            new TableRow({ children: [
              cell("🎓 教育培训", 1400), cell("你的教育机构", 1800), cell("免费试听 · 名师授课", 2200),
              cell("📍 免费试听课，到校即送学习大礼包 🎁", 3360)
            ]}),
            new TableRow({ children: [
              cell("🛒 电商零售", 1400), cell("你的网店", 1800), cell("限时特惠 · 工厂直供", 2200),
              cell("📍 线下体验店开业，到店购买全场 8 折", 3360)
            ]}),
            new TableRow({ children: [
              cell("🏪 本地生活", 1400), cell("你的店", 1800), cell("附近好店 · 新客体验", 2200),
              cell("📍 就在你附近，到店体验还有惊喜礼包", 3360)
            ]}),
            new TableRow({ children: [
              cell("📦 其他行业", 1400), cell("你的店", 1800), cell("新客专享 · 免费咨询", 2200),
              cell("📍 新客首单立减，在线预约立享优惠", 3360)
            ]}),
          ]
        }),

        heading2("2.4 关键特性"),
        bullet("仿真朋友圈卡片：白底卡片 + 蓝色用户名 + 广告标签 + 时间戳，高度还原真实朋友圈广告样式", "b1"),
        bullet("虚拟社交互动：自动生成虚拟点赞（3-6 人）+ 评论（2-3 条），评论文案按行业定制，增强代入感", "b1"),
        bullet("双模式展示：单图模式（16:9 渐变主图）+ 九宫格模式（3×3 素材展示），用户可通过 Tab 自由切换", "b1"),
        bullet("照片上传功能：支持上传最多 9 张照片，用于替换模板图；上传不足 9 张时，剩余位置用模板素材填充", "b1"),
        bullet("行业定制化：7 个行业 × 4 种目标 = 28 套广告文案组合，各有独立的九宫格模板", "b1"),
        bullet("返回编辑：任意步骤都可点击「上一页」返回前一步，返回时重置相关状态（如取消行业选择）", "b1"),
        bullet("分享生成：点击分享后，Canvas 生成包含朋友圈卡片 + 二维码的图片，支持原生 Share API 或下载", "b1"),
        bullet("进度引导：多步骤圆点指示器，降低用户流失", "b1"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 3. H5 #2 认知纠正选择题 =====
        heading1("3. H5 #2 — 你不知道的微信广告热知识"),

        heading2("3.1 产品概述"),
        p("通过 6 道精心设计的选择题，逐条击破新客对微信广告的常见误解。每道题选完后即时展示「老炮说」专家解读，以「行业老炮」的口吻给出通俗易懂的正确认知。最终根据得分给出用户画像分级。"),
        p("核心心理机制：认知纠正 + 权威背书——选错不丢人，因为「很多人跟你一样被误解耽误了」。", { italic: true }),

        heading2("3.2 用户流程"),
        new Table({
          columnWidths: [1200, 2400, 2400, 3360],
          rows: [
            new TableRow({ children: [hCell("步骤", 1200), hCell("页面", 2400), hCell("用户动作", 2400), hCell("系统响应", 3360)] }),
            new TableRow({ children: [
              cell("Step 1", 1200), cell("封面页", 2400), cell("点击「开始测试」", 2400),
              cell("引导语说明答题机制 + 老炮解读 + 底部标注「完成测试可领取行业投放数据」", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 2", 1200), cell("答题页", 2400), cell("逐题选择答案", 2400),
              cell("选中后标绿 ✓（正确）/ 标红 ✗（错误），弹出「老炮说」", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 3", 1200), cell("结果页", 2400), cell("查看得分 + 分级结果", 2400),
              cell("得分条动画 + 每题答对/答错汇总 + 结果人格卡", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 4", 1200), cell("CTA 页", 2400), cell("扫码 / 分享", 2400),
              cell("领取腾讯广告行业报告 + 企微二维码", 3360)
            ]}),
          ]
        }),

        heading2("3.3 题库详情（6 道认知纠正题）"),

        heading3("Q1: 在微信投广告，最少要花多少钱？"),
        bullet("❌ 至少 1 万起步", "b2"),
        bullet("✅ 几十块就能开始（正确答案）", "b2"),
        bullet("❌ 要看行业，最少也得几千", "b2"),
        p("🔥 老炮说：没有起投门槛，几十块就能跑。很多小店老板第一次投 100 块试水，效果好了再加。别被「烧钱」两个字吓到。", { italic: true }),

        heading3("Q2: 朋友圈广告，能指定给谁看吗？"),
        bullet("❌ 不能，系统随机投放", "b3"),
        bullet("❌ 只能选城市，不能更精准", "b3"),
        bullet("✅ 能按地区、年龄、兴趣等精准定向（正确答案）", "b3"),
        p("🔥 老炮说：「附近 3 公里的 25-40 岁女性」都能精准投。本地商家一定要用地域定向，半径越小越精准。", { italic: true }),

        heading3("Q3: 没有设计师，能投广告吗？"),
        bullet("❌ 不行，必须有专业素材", "b4"),
        bullet("✅ 可以，有工具能自动生成素材（正确答案）", "b4"),
        bullet("❌ 可以，但效果会很差", "b4"),
        p("🔥 老炮说：有智能创意工具，选模板传产品图就行。手机拍的真实照片反而比精修海报效果好，用户觉得真实。", { italic: true }),

        heading3("Q4: 微信广告只能投朋友圈？"),
        bullet("❌ 对，只有朋友圈", "b5"),
        bullet("✅ 朋友圈、公众号、视频号、小程序都能投（正确答案）", "b5"),
        p("🔥 老炮说：朋友圈只是其中一个位置。视频号现在流量红利大，有短视频素材的优先试。", { italic: true }),

        heading3("Q5: 广告投出去了，钱没花完能退吗？"),
        bullet("❌ 不能，投出去就是泼出去的水", "b6"),
        bullet("✅ 能退，随时能停，余额可退（正确答案）", "b6"),
        p("🔥 老炮说：随时能暂停投放，没花完的余额可以退。预算你说了算，不是系统说了算。", { italic: true }),

        heading3("Q6: 投微信广告需要专业团队吗？"),
        bullet("❌ 需要，至少得有个投手", "b7"),
        bullet("❌ 最好有，不然搞不定", "b7"),
        bullet("✅ 一个人就能搞定（正确答案）", "b7"),
        p("🔥 老炮说：现在有行业智投模式，选行业、传素材、设预算三步就搞定。把省下来的钱投广告不香吗？", { italic: true }),

        heading2("3.4 结果页设计（去掉题目，只展示 6 大真相）"),
        p("完成所有 6 题后，不再逐题回顾选对/选错，而是直接展示「微信广告 6 个真相」——6 条核心认知的一句话总结，帮助用户快速理解最重要的内容。", { bold: true }),
        new Table({
          columnWidths: [3360, 6000],
          rows: [
            new TableRow({ children: [hCell("序号", 3360), hCell("6 大真相（结果页展示）", 6000)] }),
            new TableRow({ children: [
              cell("1", 3360), cell("✅ 几十块就能开始投，没有起投门槛", 6000)
            ]}),
            new TableRow({ children: [
              cell("2", 3360), cell("✅ 能按地区、年龄、兴趣精准定向", 6000)
            ]}),
            new TableRow({ children: [
              cell("3", 3360), cell("✅ 没有设计师也能投，有工具自动生成素材", 6000)
            ]}),
            new TableRow({ children: [
              cell("4", 3360), cell("✅ 不止朋友圈，公众号、视频号、小程序都能投", 6000)
            ]}),
            new TableRow({ children: [
              cell("5", 3360), cell("✅ 随时能停，余额可退，预算你说了算", 6000)
            ]}),
            new TableRow({ children: [
              cell("6", 3360), cell("✅ 一个人就能搞定，有行业智投模式", 6000)
            ]}),
          ]
        }),

        heading2("3.5 得分分级"),
        new Table({
          columnWidths: [1800, 1800, 2400, 3360],
          rows: [
            new TableRow({ children: [hCell("得分", 1800), hCell("Emoji", 1800), hCell("称号", 2400), hCell("结语", 3360)] }),
            new TableRow({ children: [
              cell("5-6 题", 1800), cell("🧠", 1800), cell("投放老司机", 2400),
              cell("你对微信广告的了解超过大多数人，可以直接上手了。", 3360)
            ]}),
            new TableRow({ children: [
              cell("3-4 题", 1800), cell("🤔", 1800), cell("有点底子", 2400),
              cell("知道一些，但还有几个关键认知需要纠正。了解一下行业数据会帮你少走弯路。", 3360)
            ]}),
            new TableRow({ children: [
              cell("0-2 题", 1800), cell("😯", 1800), cell("被误解耽误了", 2400),
              cell("很多人跟你一样，对投广告有不少误解。其实没那么难，也没那么贵。", 3360)
            ]}),
          ]
        }),

        heading2("3.6 关键特性"),
        bullet("即时反馈：选中后立即标绿/标红，正确答案始终标绿高亮", "b8"),
        bullet("老炮说模块：每题选完后弹出黄色调卡片，以行业老炮口吻解读正确认知", "b8"),
        bullet("进度条：顶部进度条 + 题号显示（1/6, 2/6...），减少中途流失", "b8"),
        bullet("题目切换动画：fade out + translateX(-15px) → 加载新题 → fade in", "b8"),
        bullet("结果页简化：只展示 6 个核心真相总结，不再逐题回顾，提升浏览体验", "b8"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 4. H5 #3 投放老板类型 =====
        heading1("4. H5 #3 — 测测你的微信生意潜力"),

        heading2("4.1 产品概述"),
        p("5 道多维度评估题，从「客户来源、内容能力、接客能力、产品口碑、同行感知」五个维度全面诊断用户的生意在微信上的获客潜力。最终生成一个「潜力分」（50-98 分），以及针对弱项的个性化改进建议。"),
        p("核心心理机制：多维度诊断 + 个性化建议——让用户了解自己的优势和待改进点，为投广告做好准备。", { italic: true }),

        heading2("4.2 用户流程"),
        new Table({
          columnWidths: [1200, 2400, 2400, 3360],
          rows: [
            new TableRow({ children: [hCell("步骤", 1200), hCell("页面", 2400), hCell("用户动作", 2400), hCell("系统响应", 3360)] }),
            new TableRow({ children: [
              cell("Step 1", 1200), cell("封面页", 2400), cell("点击「开始测试」", 2400),
              cell("引导语展示测试目标 + 底部标注「内含腾讯广告行业数据」", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 2", 1200), cell("答题页", 2400), cell("逐题选择（支持返回修改）", 2400),
              cell("5 道维度题，每题 4 选项，选中后记录得分，可点返回按钮修改", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 3", 1200), cell("加载页", 2400), cell("等待 2 秒", 2400),
              cell("loading 动画 +「正在分析你的生意潜力…」", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 4", 1200), cell("结果页", 2400), cell("查看潜力分 + 建议", 2400),
              cell("潜力分（50-98）+ 等级（🚀/💪/🌱）+ 5 维度的个性化建议", 3360)
            ]}),
            new TableRow({ children: [
              cell("Step 5", 1200), cell("CTA 页", 2400), cell("扫码", 2400),
              cell("领取腾讯广告行业数据 + 企微二维码", 3360)
            ]}),
          ]
        }),

        heading2("4.3 题库详情（5 道维度题）"),

        heading3("维度 1：客户来源 — 你现在的客户主要从哪来？"),
        bullet("✅ 3 分：老客转介绍 / 口碑", "b9"),
        bullet("✅ 3 分：路过进店 / 到店消费", "b9"),
        bullet("⚡ 2 分：线上平台（如 58、大众点评、小红书）", "b9"),
        bullet("🌱 1 分：还没稳定客源", "b9"),
        p("💡 反馈：3分优势 | 已有稳定客源，广告能帮你放大现有优势，获客成本更低", { italic: true }),

        heading3("维度 2：内容能力 — 你平时会在朋友圈发产品/服务内容吗？"),
        bullet("✅ 3 分：经常发，有人看了来问价", "b10"),
        bullet("⚡ 2 分：偶尔发", "b10"),
        bullet("⚡ 2 分：很少发", "b10"),
        bullet("🌱 1 分：基本不发", "b10"),
        p("💡 反馈：3分优势 | 已有内容基础，做广告素材不费劲，朋友圈的内容稍微改改就能投", { italic: true }),

        heading3("维度 3：接客能力 — 如果突然多来 10 个客户咨询，你能接得住吗？"),
        bullet("✅ 3 分：完全没问题", "b11"),
        bullet("✅ 3 分：忙一点但能应付", "b11"),
        bullet("⚡ 2 分：可能要找人帮忙", "b11"),
        bullet("🌱 1 分：还没想过", "b11"),
        p("💡 反馈：3分优势 | 接得住，投了不浪费，广告来的每一个咨询都能转化成生意", { italic: true }),

        heading3("维度 4：产品口碑 — 买过的客户一般怎么评价你？"),
        bullet("✅ 3 分：回头客多，经常被推荐", "b12"),
        bullet("✅ 3 分：还不错，偶尔复购", "b12"),
        bullet("⚡ 2 分：跟同行差不多", "b12"),
        bullet("🌱 1 分：刚起步", "b12"),
        p("💡 反馈：3分优势 | 回头客多是投广告最大的底气，新客进来也容易留住", { italic: true }),

        heading3("维度 5：同行感知 — 你觉得同行有在微信上投广告吗？"),
        bullet("✅ 3 分：有，经常刷到", "b9"),
        bullet("⚡ 2 分：好像有", "b9"),
        bullet("⚡ 2 分：应该没有", "b9"),
        bullet("🌱 1 分：没关注过", "b9"),
        p("💡 反馈：3分优势 | 同行已经在投，说明你的行业在微信上有机会，现在入场正当时", { italic: true }),

        heading2("4.4 评分体系与潜力分"),
        new Table({
          columnWidths: [2000, 2000, 2400, 3360],
          rows: [
            new TableRow({ children: [hCell("得分范围（总分）", 2000), hCell("潜力分（换算）", 2000), hCell("等级 Emoji", 2400), hCell("等级称号", 3360)] }),
            new TableRow({ children: [
              cell("13-15 分", 2000), cell("85-98 分", 2000), cell("🚀", 2400),
              cell("天然适配：基础条件非常好，投起来大概率能跑正", 3360)
            ]}),
            new TableRow({ children: [
              cell("10-12 分", 2000), cell("70-84 分", 2000), cell("💪", 2400),
              cell("很有潜力：基础不错，优化几个小点效果翻倍", 3360)
            ]}),
            new TableRow({ children: [
              cell("5-9 分", 2000), cell("55-69 分", 2000), cell("🌱", 2400),
              cell("有前景：不是不能投，准备好了效果更好", 3360)
            ]}),
          ]
        }),
        p("潜力分计算公式：50 + (总分 - 5) × 4.8 = 最终分数（向上取整至 98）", { italic: true, color: "718096" }),

        heading2("4.5 结果页个性化建议"),
        p("每个用户的 5 道题得分都会生成对应的改进建议。例如：", { bold: true }),
        bullet("客户来源得 1 分 → 「有潜力，广告正好帮你找到第一批精准客户，很多老板都是从 0 开始的」", "b9"),
        bullet("内容能力得 2 分 → 「可以固定每周发 3 条产品内容，积累的素材直接就能当广告用」", "b9"),
        bullet("接客能力得 1 分 → 「可以先设好自动回复和话术模板，准备好了投起来事半功倍」", "b9"),

        heading2("4.6 关键特性"),
        bullet("多维度评估：5 个不同维度（客户来源、内容能力、接客能力、产品口碑、同行感知），全面诊断用户的生意状态", "b12"),
        bullet("即时得分反馈：每题选完后记录该维度得分（1/2/3 分），计入总分", "b12"),
        bullet("返回编辑功能：支持返回修改前一题答案（弹出上一次选择，返回时清除该题得分）", "b12"),
        bullet("仪式感加载页：2 秒 loading 动画 +「正在分析你的生意潜力…」文案，制造期待感", "b12"),
        bullet("潜力分与等级：5-15 总分换算为 50-98 潜力分，配合 🚀/💪/🌱 三等级展示", "b12"),
        bullet("个性化建议：根据每个维度的单个得分生成对应的改进建议，而非通用结果", "b12"),
        bullet("进度条透明化：题号显示（1/5, 2/5...）+ 水平进度条，用户明确了解答题进度", "b12"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 5. 技术方案 =====
        heading1("5. 技术方案"),

        heading2("5.1 技术栈"),
        new Table({
          columnWidths: [2400, 6960],
          rows: [
            new TableRow({ children: [hCell("维度", 2400), hCell("方案", 6960)] }),
            new TableRow({ children: [cell("前端框架", 2400, { bold: true }), cell("纯 HTML + CSS + Vanilla JS，零框架依赖", 6960)] }),
            new TableRow({ children: [cell("样式系统", 2400, { bold: true }), cell("shared.css 共享设计系统 + 每个 H5 内联 <style> 自定义样式", 6960)] }),
            new TableRow({ children: [cell("布局基准", 2400, { bold: true }), cell("移动优先，375px 基准宽度，暗色主题（#0e0e0e / #111）", 6960)] }),
            new TableRow({ children: [cell("页面导航", 2400, { bold: true }), cell("Screen-based 单页模式：.screen.active 切换 + goTo(n) 函数", 6960)] }),
            new TableRow({ children: [cell("服务部署", 2400, { bold: true }), cell("Python http.server 8081（nohup 持久化）", 6960)] }),
            new TableRow({ children: [cell("后端依赖", 2400, { bold: true }), cell("无后端，纯静态页面，所有逻辑在客户端完成", 6960)] }),
          ]
        }),

        heading2("5.2 共享设计系统（shared.css）"),
        bullet("全局暗色主题：#0e0e0e 背景 + rgba 白色分层文字", "n1"),
        bullet("按钮系统：.btn-primary（白底黑字 / hover 缩放）、.share-btn（透明底白描边）", "n1"),
        bullet("标签系统：.tag-green / .tag-yellow（圆角小标签）", "n1"),
        bullet("进度条：.progress-bar + .fill（绿色 #4ADE80 填充）", "n1"),
        bullet("卡片底色：#1a1a1a + 1px rgba 白色描边 + 24px 圆角", "n1"),
        bullet("动画：fadeInUp（从下向上渐现）", "n1"),
        bullet("企微二维码占位：.qr-placeholder（虚线框 + 灰色描边）", "n1"),

        heading2("5.3 目录结构"),
        p("h5-matrix-v2/"),
        p("├── index.html          （Hub 导航页 — 3 张卡片入口）", { indent: true }),
        p("├── shared.css          （共享设计系统）", { indent: true }),
        p("├── roast/index.html    （H5 #1 朋友圈广告模拟器）", { indent: true }),
        p("├── wechat-ads/index.html （H5 #2 热知识测试 — 6 道认知纠正题）", { indent: true }),
        p("└── boss-type/index.html  （H5 #3 生意潜力测试 — 5 维度多维诊断）", { indent: true }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 6. 统一转化设计 =====
        heading1("6. 统一转化设计"),

        heading2("6.1 CTA 策略"),
        p("三个 H5 使用统一转化口径，降低用户决策成本："),
        new Table({
          columnWidths: [2400, 3600, 3360],
          rows: [
            new TableRow({ children: [hCell("要素", 2400), hCell("内容", 3600), hCell("说明", 3360)] }),
            new TableRow({ children: [cell("主 CTA 按钮", 2400, { bold: true }), cell("领取行业投放数据", 3600), cell("统一文案，三个 H5 一致", 3360)] }),
            new TableRow({ children: [cell("二维码", 2400, { bold: true }), cell("企业微信二维码", 3600), cell("扫码添加企微，进入 SDR 跟进链路", 3360)] }),
            new TableRow({ children: [cell("副 CTA", 2400, { bold: true }), cell("分享给朋友", 3600), cell("Web Share API 优先，降级为截图提示", 3360)] }),
          ]
        }),

        heading2("6.2 合规红线"),
        bullet("所有页面不涉及效果承诺：无到达率、转化率、ROI、「X 天见效」等数据", "n2"),
        bullet("不出现「保证」「确保」「一定」等承诺性措辞", "n2"),
        bullet("模拟器页面标注「效果模拟 · 仅供参考」提示", "n2"),
        bullet("所有互动数据（行业、目标、得分）不做持久化存储，不涉及用户隐私", "n2"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 7. 投放建议 =====
        heading1("7. 投放建议与分发策略"),

        heading2("7.1 渠道矩阵"),
        new Table({
          columnWidths: [2400, 3360, 3600],
          rows: [
            new TableRow({ children: [hCell("渠道", 2400), hCell("推荐 H5", 3360), hCell("投放建议", 3600)] }),
            new TableRow({ children: [
              cell("朋友圈广告", 2400, { bold: true }), cell("H5 #1（朋友圈模拟器）", 3360),
              cell("「在朋友圈看你的广告长什么样」——场景强关联，代入感最强", 3600)
            ]}),
            new TableRow({ children: [
              cell("公众号文章", 2400, { bold: true }), cell("H5 #2（热知识测试）", 3360),
              cell("「测测你对投广告了解多少」——文末嵌入，内容型流量自然过渡", 3600)
            ]}),
            new TableRow({ children: [
              cell("视频号信息流", 2400, { bold: true }), cell("H5 #3（生意潜力）", 3360),
              cell("「你的生意在微信上有多大潜力？」——社交属性强，适合裂变传播", 3600)
            ]}),
            new TableRow({ children: [
              cell("社群裂变", 2400, { bold: true }), cell("三个 H5 交替使用", 3360),
              cell("商家社群 / 行业交流群分发，结果卡截图分享率高", 3600)
            ]}),
          ]
        }),

        heading2("7.2 AB 测试建议"),
        bullet("封面文案 AB 测试：分别测试不同 hook 文案的点击率", "n3"),
        bullet("CTA 按钮文案 AB 测试：「领取行业数据」vs「免费获取行业报告」", "n3"),
        bullet("H5 #2 题目数量：6 题 vs 4 题（测试完成率差异）", "n3"),
        bullet("H5 #3 结果样式：带维度详情 vs 仅潜力分（测试分享率差异）", "n3"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 8. 待办与风险 =====
        heading1("8. 待办事项与风险提示"),

        heading2("8.1 待办事项"),
        new Table({
          columnWidths: [600, 3600, 2400, 2760],
          rows: [
            new TableRow({ children: [hCell("#", 600), hCell("事项", 3600), hCell("负责方", 2400), hCell("优先级", 2760)] }),
            new TableRow({ children: [
              cell("1", 600), cell("替换企微二维码占位符为真实二维码", 3600), cell("运营", 2400), cell("P0 — 上线前必须", 2760, { color: "E53E3E" })
            ]}),
            new TableRow({ children: [
              cell("2", 600), cell("H5 #3 五维度题库 + 评分 + 个性化建议", 3600), cell("策划", 2400), cell("✅ 已完成", 2760, { color: "38A169" })
            ]}),
            new TableRow({ children: [
              cell("3", 600), cell("埋点方案：页面 PV、按钮点击、完成率、分享率", 3600), cell("数据", 2400), cell("P1 — 首版可后补", 2760, { color: "DD6B20" })
            ]}),
            new TableRow({ children: [
              cell("4", 600), cell("合规审核：文案终审 + 广告标识合规", 3600), cell("法务", 2400), cell("P0 — 上线前必须", 2760, { color: "E53E3E" })
            ]}),
            new TableRow({ children: [
              cell("5", 600), cell("正式域名部署 + HTTPS + CDN 加速", 3600), cell("技术", 2400), cell("P0 — 上线前必须", 2760, { color: "E53E3E" })
            ]}),
            new TableRow({ children: [
              cell("6", 600), cell("微信 JS-SDK 接入（分享卡片自定义标题/图片）", 3600), cell("技术", 2400), cell("P1 — 体验优化", 2760, { color: "DD6B20" })
            ]}),
            new TableRow({ children: [
              cell("7", 600), cell("H5 #1 照片上传 + 社交互动 + Canvas 分享图", 3600), cell("策划", 2400), cell("✅ 已完成", 2760, { color: "38A169" })
            ]}),
            new TableRow({ children: [
              cell("8", 600), cell("H5 #2 结果页改为 6 大真相总结（不再逐题回顾）", 3600), cell("策划", 2400), cell("✅ 已完成", 2760, { color: "38A169" })
            ]}),
            new TableRow({ children: [
              cell("9", 600), cell("H5 #1 / #3 返回按钮（上一页/上一题）", 3600), cell("策划", 2400), cell("✅ 已完成", 2760, { color: "38A169" })
            ]}),
          ]
        }),

        heading2("8.2 风险提示"),
        new Table({
          columnWidths: [3600, 2400, 3360],
          rows: [
            new TableRow({ children: [hCell("风险", 3600), hCell("影响", 2400), hCell("应对方案", 3360)] }),
            new TableRow({ children: [
              cell("模拟器被误认为真实广告", 3600), cell("用户投诉", 2400),
              cell("已标注「效果模拟 · 仅供参考」，法务确认合规", 3360)
            ]}),
            new TableRow({ children: [
              cell("答题页中途流失率高", 3600), cell("获客效率低", 2400),
              cell("每题即时反馈 + 进度条 + 控制题数 ≤6", 3360)
            ]}),
            new TableRow({ children: [
              cell("结果分享后无法追踪来源", 3600), cell("裂变效果难量化", 2400),
              cell("后续接入 UTM 参数 + 短链追踪", 3360)
            ]}),
            new TableRow({ children: [
              cell("纯静态无后端，数据无法沉淀", 3600), cell("无法分析用户行为", 2400),
              cell("P1 接入埋点 SDK（如神策 / GrowingIO）", 3360)
            ]}),
          ]
        }),

        new Paragraph({ spacing: { before: 600 }, children: [] }),
        divider(),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100 },
          children: [new TextRun({ text: "— 文档结束 —", size: 20, color: "A0AEC0", font: "Arial", italics: true })] }),
      ]
    }
  ]
});

// ========== Export ==========
const outputPath = '/Users/yuyangcai/.workbuddy/workspace/default_project/h5-matrix-v2/腾讯广告_H5获客矩阵_MRD.docx';
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log('✅ MRD 已生成：' + outputPath);
}).catch(err => {
  console.error('❌ 生成失败：', err);
});
