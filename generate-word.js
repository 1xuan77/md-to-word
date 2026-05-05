import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from "docx";
import { writeFileSync } from "fs";

async function generateDoc() {
  const bodyFont = "Microsoft YaHei";
  const bodySize = 24; // half-points = 12pt
  const smallSize = 22;

  function p(text, { bold = false, size = bodySize, color, spacing } = {}) {
    return new Paragraph({
      spacing: { after: 120, ...spacing },
      children: [new TextRun({ text, bold, size, color, font: bodyFont })],
    });
  }

  function heading(text, level) {
    return new Paragraph({
      text,
      heading: level,
      spacing: { before: level === HeadingLevel.HEADING_1 ? 360 : 240, after: 160 },
    });
  }

  function bullet(text, opts = {}) {
    return new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: `  •  ${text}`, size: bodySize, font: bodyFont, ...opts })],
    });
  }

  function subHeading(text) {
    return new Paragraph({
      text,
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 200, after: 100 },
    });
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // === 标题 ===
          new Paragraph({
            text: "2026年5月5日 中国热点新闻速览",
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
          }),

          // === 副标题 ===
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
            children: [
              new TextRun({
                text: "整理时间：2026年5月5日  |  数据来源：公开新闻报道",
                size: 22,
                color: "888888",
                font: bodyFont,
              }),
            ],
          }),

          // ============ 一、头条要闻 ============
          heading("一、头条要闻", HeadingLevel.HEADING_1),

          subHeading("1. 湖南浏阳烟花厂爆炸事故"),
          p("5月4日下午，湖南省长沙市浏阳市华盛烟花制造燃放有限公司发生爆炸事故。截至当日19时，事故已造成21人死亡、61人受伤。习近平总书记作出重要指示，强调要全力救治伤员，做好善后工作，抓好重点行业领域风险隐患排查整治，加强公共安全管理，确保人民群众生命财产安全。李强总理作出批示。近500名消防救援人员赶赴现场开展救援。"),

          subHeading("2. 习近平向全国青年致以五四节日问候"),
          p("在五四青年节到来之际，习近平给中国青年五四奖章暨新时代青年先锋奖获奖者代表回信，向全国各族青年致以节日祝贺和诚挚问候，勉励广大青年在推进中国式现代化中奋勇争先、建功立业。"),

          // ============ 二、五一假期·交通与旅游 ============
          heading("二、五一假期·交通与旅游", HeadingLevel.HEADING_1),

          subHeading("1. 返程高峰来临"),
          p("5月5日是五一假期最后一天，全国迎来返程客流高峰。5月4日全社会跨区域人员流动量预计超2.93亿人次，同比增长4.2%。铁路方面，4日全国铁路预计发送旅客2030万人次，加开列车1641列。琼州海峡迎来出岛返程高峰，交通部门全力做好出行服务保障。"),

          subHeading("2. 五一假期数据亮眼"),
          bullet("整个五一假期预计全社会跨区域人员流动量达15.2亿人次，同比增长约4%"),
          bullet("5月1日单日出行达3.44亿人次，创历史新高"),
          bullet("铁路单日发送旅客2480万人次，刷新纪录"),
          bullet("新能源汽车高速公路出行量同比增长33%"),
          bullet("文化和旅游部推出13700余场文化活动，发放2.84亿元消费券"),

          subHeading("3. 文旅消费新趋势"),
          bullet("美团平台音乐节搜索量自4月以来增长4倍"),
          bullet("民族旅拍、热气球、体育旅游、徒步、采摘等体验式消费火爆"),
          bullet("北京假期期间上演1400余场演出"),
          bullet("旅游消费从传统观光向沉浸式文化体验转型"),

          // ============ 三、经济数据 ============
          heading("三、经济数据", HeadingLevel.HEADING_1),

          bullet("电影票房：2026年度总票房（含预售）突破135亿元，五一档票房突破6亿元", { bold: true }),
          bullet("农产品电商：一季度全国农产品网上零售额同比增长14.7%", { bold: true }),
          bullet("海洋经济：一季度海洋生产总值达2.6万亿元，同比增长5.4%", { bold: true }),
          bullet("余额宝收益新低：天弘余额宝7日年化收益率报0.9980%，首次跌破1%，创历史新低", { bold: true }),
          bullet("十五五规划：研发支出增长40%，核心数字经济产业目标占GDP的12.5%，军费增长7%（约2770亿美元），2026年经济增长目标设定为4.5%-5%", { bold: true }),

          // ============ 四、科技与产业 ============
          heading("四、科技与产业", HeadingLevel.HEADING_1),

          subHeading("1. 中国加速科技自立自强"),
          p("面对中美竞争加剧，中国在十五五规划中重点布局AI、量子计算、6G、脑机接口和聚变能源等领域。在风险更高的时代，中国押注技术以抵御美国压力，推动关键核心技术自主可控。"),

          subHeading("2. 科技投入持续加大"),
          p("研发支出较上一个五年计划增长40%，推动数字经济核心产业占GDP比重提升至12.5%，科技自立自强上升为国家战略。"),

          // ============ 五、体育捷报 ============
          heading("五、体育捷报", HeadingLevel.HEADING_1),

          subHeading("1. 吴宜泽夺得斯诺克世锦赛冠军"),
          p("中国00后小将吴宜泽以18:17惊险战胜对手，夺得2026年斯诺克世界锦标赛冠军，成为历史上首位00后世锦赛冠军，为中国斯诺克书写新篇章。"),

          subHeading("2. 中国队汤姆斯杯第12次捧杯"),
          p("中国羽毛球队在汤姆斯杯决赛中以3:1战胜法国队，第12次夺得汤姆斯杯冠军，彰显中国羽毛球强大实力。"),

          // ============ 六、社会热点 ============
          heading("六、社会热点", HeadingLevel.HEADING_1),

          bullet("辽宁丹东交通事故：发生严重超载交通事故，已致8人死亡、13人受伤"),
          bullet("五台山救援：五台山突降大雪，300名被困徒步人员成功获救"),
          bullet(`山西“订婚强奸案”：当事人刑满出狱，持续引发社会关注`),
          bullet("汉坦病毒邮轮事件：世卫组织通报大西洋一艘邮轮出现汉坦病毒感染，3人死亡"),

          // ============ 七、国际关注 ============
          heading("七、国际关注", HeadingLevel.HEADING_1),

          bullet("特朗普访华计划：特朗普称期待两周内访华与习近平会面，计划时间约为5月14日至15日，但预测市场对成行概率大幅下调"),
          bullet("中美博弈：中国指示企业无视美国对伊朗石油制裁，中美关系持续紧张"),
          bullet("伊朗局势：伊朗阻止美以驱逐舰驶入霍尔木兹海峡"),
          bullet("历史档案移交：法国友人向南京移交日本侵华历史档案"),

          // === 尾部声明 ===
          new Paragraph({
            spacing: { before: 500, after: 100 },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "本文档由AI自动整理生成，数据来源于公开新闻报道，仅供参考。",
                size: 20,
                color: "999999",
                font: bodyFont,
              }),
            ],
          }),
        ],
      },
    ],
  });

  const title = "2026年5月5日 中国热点新闻速览";
  const filename = `${title}.docx`;
  const buffer = await Packer.toBuffer(doc);
  writeFileSync(filename, buffer);
  console.log(`Word 文档已生成: ${filename}`);
}

generateDoc().catch(console.error);
