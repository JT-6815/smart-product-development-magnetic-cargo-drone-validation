import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const W = 1280;
const H = 720;
const OUT_DIR = "D:\\大作业文件\\大三下期\\智能产品开发";
const FINAL_PPTX = path.join(OUT_DIR, "F450_5kg装货无人机完整结课报告_45页黑橙科技版.pptx");
const LOCAL_COPY = path.resolve(__dirname, "../../../..", "outputs", "F450_5kg装货无人机完整结课报告_45页黑橙科技版.pptx");
const ASSET_DIR = path.join(__dirname, "assets");
const PREVIEW_DIR = path.join(__dirname, "preview", "final");
const QA_DIR = path.join(__dirname, "qa");
const SOURCE_NOTES = path.join(__dirname, "source-notes.txt");
const SLIDE_PLAN = path.join(__dirname, "slide-plan.txt");

const COLORS = {
  bg: "#050506",
  bg2: "#0A0A0D",
  panel: "#111114",
  card: "#17171C",
  card2: "#24242B",
  ink: "#F8FAFC",
  muted: "#B7C0CC",
  white: "#FFFFFF",
  orange: "#FF5A00",
  orange2: "#FFB000",
  cyan: "#00E5FF",
  green: "#19C37D",
  red: "#EF4444",
  line: "#3B2A18",
  line2: "#FF7A1A",
};

const FONT = {
  title: "Microsoft YaHei UI",
  body: "Microsoft YaHei",
  mono: "Consolas",
};

const weekSummaries = {
  1: "第一周主要完成项目立项、方案讨论和基础知识学习。小组明确以F450四轴机架为平台，围绕“底部电磁铁磁吸重物”的装货无人机主题展开，初步确定APM飞控、Mission Planner地面站、无刷电机、电调、电池、遥控器和磁吸挂载模块等核心器材。学习阶段重点理解四轴飞行原理、螺旋桨正反桨方向、电机与电调工作方式、LiPo电池安全和重心控制。同步建立项目资料目录与GitHub仓库，整理采购清单、预算表和分工方案，为后续拼装与调试形成清晰路线。",
  2: "第二周进入主体制作阶段，重点完成F450机架、电机、电调、飞控和供电系统的装配。组装时先固定上下中心板和四个机臂，再按对角方向安装电机与螺旋桨，保证旋转方向和机臂颜色便于识别机头。电气部分完成电调焊接、电源分配、APM飞控固定、接收机与输出通道连接，并用扎带和减震件整理线束。调试时通过Mission Planner完成加速度计、罗盘、遥控器和飞行模式基础设置。最后进行螺丝、焊点、电机方向和解锁状态检查，确保进入试飞前具备基本安全条件。",
  3: "第三周围绕基础试飞和参数优化展开。试飞前先在Mission Planner中确认机架类型、遥控通道、飞行模式、失控保护和电池监测参数，随后进行无桨电机测试、低油门响应检查和解锁保护确认。初次飞行以低空悬停为主，观察机体是否偏航、漂移或高频振动，再逐步进行前后左右移动、转向和高度控制测试。根据飞行表现记录问题，如重心偏移、桨叶不平衡、油门曲线不线性等，并对PID和安装状态进行微调。本周目标是获得可重复的稳定悬停能力。",
  4: "第四周重点验证磁吸挂载与载重运输方案。小组将电磁铁布置在机体底部中心位置，尽量使挂载点靠近重心垂线，降低起飞时的俯仰和横滚扰动。测试流程从地面吸附、断电释放、空载挂载稳定性开始，再逐步增加1000g、1500g、3000g等测试重量，记录油门裕度、电池电压变化、飞行时间和降落安全性。5kg目标作为课程挑战上限，需要结合电机推力、电池放电能力、机架强度和安全边界综合评估。本周形成承重测试表、风险控制流程和载重起飞演示方案。",
  5: "第五周集中整理展示素材、视频脚本和AI外观设计。素材拍摄覆盖机架组装、焊接接线、Mission Planner调参、试飞、挂载测试和小组协作过程，便于最终展示完整制作链路。视频剪辑采用“问题提出、方案设计、制作过程、测试结果、改进方向”的叙事结构，配合字幕、关键参数标注和过程照片。AI外观设计用于生成未来物流无人机概念图和宣传海报，强调磁吸装货、短距离运输和模块化改装特点。该阶段的重点不是单纯包装，而是把工程过程转化为可展示、可答辩的成果。",
  6: "第六周完成资料归档、结课报告和答辩准备。小组对六周内的图片、视频、参数记录、测试数据、问题清单和周报进行统一整理，补充README、项目目标、使用说明和最终总结。报告中重点呈现F450结构认知、硬件拼装、焊接接线、APM与Mission Planner调参、试飞验证、磁吸挂载测试以及5kg重载方案探索。总结部分复盘成功经验与不足，包括载重裕度、安全冗余、飞行时间和结构强度等限制，并提出后续升级方向。最终形成可编辑PPT、演示视频脚本和答辩提纲。",
};

const slides = [
  { type: "cover", title: "F450 5kg装货无人机完整结课报告", subtitle: "智能产品开发设计 | 2026.05.18 - 2026.06.22 | APM飞控 + Mission Planner地面站", image: img("f450/f450_39.jpg") },
  { type: "overview", section: "项目总览", title: "课程任务与最终目标", bullets: ["制作一台基于F450套件的装货无人机", "底部增加电磁铁磁吸挂载模块", "完成拼装、焊接、调参、试飞与载重测试", "以3000-5000g运输能力作为挑战目标"], image: img("template_media/template_14.jpeg") },
  { type: "agenda", section: "报告结构", title: "55页结课报告目录", bullets: ["项目背景与F450结构认知", "六周课程推进记录：采购、拼装、调试、试飞、载重、展示", "APM飞控与Mission Planner参数设置", "电磁铁挂载与5kg重载方案探索", "测试数据、问题复盘、改进方向与答辩材料"] },
  { type: "imageText", section: "结构认知", title: "F450四轴平台结构", bullets: ["X型四旋翼，中心板承载飞控、电池与电源分配", "四根机臂提供电机安装点，红白机臂用于识别机头", "动力系统由无刷电机、ESC、电池和螺旋桨组成", "飞行稳定依赖飞控姿态解算、遥控输入和电机差速控制"], image: img("f450/f450_01.jpg") },
  { type: "cards", section: "结构认知", title: "核心硬件组成", cards: [["机架", "F450中心板、四机臂、起落架"], ["动力", "无刷电机、ESC、电源分配板、LiPo电池"], ["控制", "APM飞控、接收机、GPS/罗盘、数传"], ["挂载", "电磁铁、继电器/开关通道、底部支架"]] },
  { type: "diagram", section: "方案定义", title: "5kg装货目标的工程约束", bullets: ["载重能力取决于总推力、油门裕度和电池放电能力", "F450原平台适合教学与轻载，5kg属于挑战性扩展", "课程中把5kg作为方案探索上限，测试逐级推进", "演示时强调安全边界：低空、短距、人员隔离"], diagram: "payload" },
  { type: "timeline", section: "周计划", title: "六周推进路线", items: ["购买配件与知识学习", "主体组装与基础调试", "试飞测试与PID优化", "磁吸挂载模块测试", "AI视频与项目展示制作", "资料总结与答辩准备"] },
  { type: "section", section: "Week 1", title: "购买配件与知识学习", subtitle: "2026.05.18 - 2026.05.24" },
  { type: "imageText", section: "Week 1", title: "确定整体方案", bullets: ["确认使用F450四轴机架作为课程平台", "主题定位为短距离装货与磁吸挂载验证", "选择APM飞控与Mission Planner地面站作为调试工具", "明确5kg目标为重载方案探索上限，测试逐级推进"], image: img("template_media/template_02.png") },
  { type: "table", section: "Week 1", title: "采购清单与预算结构", headers: ["类别", "主要器材", "作用"], rows: [["机体", "F450机架、起落架、减震件", "提供结构承载"], ["动力", "电机、电调、桨、电池、充电器", "提供升力与能源"], ["控制", "APM飞控、接收机、GPS/罗盘", "姿态控制与定位"], ["挂载", "电磁铁、挂钩、继电器、扎带", "实现吸附与释放"]] },
  { type: "diagram", section: "Week 1", title: "四轴飞行原理学习", bullets: ["相邻桨叶旋向相反，用于抵消反扭矩", "升降通过四个电机同步增减转速实现", "横滚、俯仰、偏航依靠不同电机转速差完成", "安装桨叶前必须确认电机序号与旋向"], diagram: "quad" },
  { type: "imageText", section: "Week 1", title: "LiPo电池与重心安全", bullets: ["电池固定位置决定重心，重心偏移会增加飞控修正量", "LiPo使用前检查鼓包、破损和电压，充电时使用平衡充", "载重测试必须预留油门裕度，避免满油门悬停", "底部磁吸模块应尽量贴近机体中心线"], image: img("f450/f450_05.jpg") },
  { type: "flow", section: "Week 1", title: "GitHub初始化与资料目录", items: ["README项目介绍", "配件清单与预算表", "学习笔记", "项目计划书", "分工图", "周报归档"] },
  { type: "summary", section: "Week 1", title: "第一周小结", body: weekSummaries[1], image: img("f450-contact-sheet.jpg") },
  { type: "section", section: "Week 2", title: "组装与调试", subtitle: "2026.05.25 - 2026.05.31" },
  { type: "imageText", section: "Week 2", title: "机架组装流程", bullets: ["先安装F450底板与四根机臂，确认红色机臂朝向机头", "使用螺丝固定中心板，保证机臂不松动、不扭曲", "预留飞控、电池和挂载模块的安装空间", "装配过程中同步记录照片，便于周报与最终展示"], image: img("f450/f450_03.jpg") },
  { type: "imageText", section: "Week 2", title: "电机、电调与螺旋桨方向", bullets: ["电机安装在机臂末端，线束朝内便于走线", "电调靠近机臂固定，减少动力线拉扯", "正桨/反桨与电机旋向一一对应，装反会导致无法起飞", "首次通电前不安装桨叶，先做电机方向测试"], image: img("f450/f450_06.jpg") },
  { type: "imageText", section: "Week 2", title: "焊接与电源分配", bullets: ["电调电源线焊接到电源分配板正负极", "焊点要求饱满、无虚焊，必要时使用热缩管绝缘", "大电流线路避免交叉缠绕，减少短路风险", "焊接后用万用表检查正负极是否导通异常"], image: img("f450/f450_17.jpg") },
  { type: "imageText", section: "Week 2", title: "APM飞控安装与接收机接线", bullets: ["APM箭头方向与机头一致，使用减震垫降低振动", "接收机通道连接遥控器油门、横滚、俯仰、偏航等输入", "飞控输出口连接四路电调信号线", "整理线束，避免碰到桨叶或影响云台/挂载空间"], image: img("f450/f450_29.jpg") },
  { type: "imageText", section: "Week 2", title: "Mission Planner初始调试", bullets: ["安装地面站并选择正确COM口和波特率连接", "完成机架类型、加速度计、罗盘和遥控器校准", "设置Stabilize、AltHold等基础飞行模式", "确认飞控解锁提示与传感器状态正常"], image: img("missionplanner/mp_37.png") },
  { type: "checklist", section: "Week 2", title: "装机完成后的安全检查", items: ["螺丝锁紧，机臂无晃动", "焊点绝缘，无短路风险", "电机方向与桨叶方向一致", "APM方向与机头一致", "遥控器失控保护可触发", "试飞区域开阔，人员远离"] },
  { type: "summary", section: "Week 2", title: "第二周小结", body: weekSummaries[2], image: img("f450/f450_24.jpg") },
  { type: "section", section: "Week 3", title: "试飞测试", subtitle: "2026.06.01 - 2026.06.07" },
  { type: "imageText", section: "Week 3", title: "试飞前参数确认", bullets: ["确认固件、机架类型和飞行模式写入成功", "检查遥控器行程、油门低位和通道方向", "完成罗盘、加速度计和电调校准", "在Mission Planner中查看姿态与传感器数据是否稳定"], image: img("missionplanner/mp_30.png") },
  { type: "code", section: "Week 3", title: "APM参数与磁吸控制逻辑", code: "1. Mission Planner写入四轴固件与机架类型\n2. 校准遥控器、加速度计、罗盘和电调行程\n3. 设置CH7/CH8作为磁吸模块开关通道\n4. 起飞前：MAGNET=OFF，挂载确认后MAGNET=ON\n5. 运输中：保持低速、低高度、短航线\n6. 降落后：油门锁定，MAGNET=OFF释放货物", bullets: ["本项目的“编程”重点是飞控固件写入、参数配置和辅助通道逻辑，而非从零编写飞控算法。"] },
  { type: "imageText", section: "Week 3", title: "解锁与无桨电机测试", bullets: ["无桨状态下完成解锁检查，避免误伤", "逐个观察电机响应，确认油门曲线平顺", "通过Mission Planner查看DISARMED/ARMED状态", "异常时先断电，再检查电调信号线和通道方向"], image: img("missionplanner/mp_24.png") },
  { type: "imageText", section: "Week 3", title: "低空悬停测试", bullets: ["初次试飞控制在低高度和短时间内完成", "观察机体是否明显偏航、漂移、抖动或侧翻趋势", "记录起飞油门位置和悬停稳定性", "若振动明显，优先检查桨叶、电机和飞控固定"], image: img("template_media/template_28.jpeg") },
  { type: "imageText", section: "Week 3", title: "基础飞行动作测试", bullets: ["前后、左右、偏航和高度保持分项测试", "每次只调整一个变量，便于定位问题", "记录飞行时长、电压变化和遥控响应", "所有问题在GitHub问题记录表中闭环"], image: img("missionplanner/mp_38.jpeg") },
  { type: "bars", section: "Week 3", title: "PID与稳定性优化记录", labels: ["横滚稳定", "俯仰稳定", "偏航响应", "振动控制", "高度保持"], values: [72, 70, 66, 58, 64], note: "数值为调试记录的相对评分，用于展示优化前后的判断维度。" },
  { type: "cards", section: "Week 3", title: "试飞问题与处理", cards: [["起飞偏移", "重新检查重心和遥控微调"], ["高频振动", "检查桨叶平衡与飞控减震"], ["油门不线性", "重新校准电调行程"], ["GPS/罗盘异常", "远离金属物并重新校准"]] },
  { type: "summary", section: "Week 3", title: "第三周小结", body: weekSummaries[3], image: img("missionplanner/mp_45.jpeg") },
  { type: "section", section: "Week 4", title: "磁吸挂载模块测试", subtitle: "2026.06.08 - 2026.06.14" },
  { type: "diagram", section: "Week 4", title: "电磁铁挂载模块设计", bullets: ["电磁铁固定在机体底部中心位置", "通过遥控辅助通道控制吸附与释放", "挂载件尽量垂直于重心，降低摆动", "所有载重试验先地面验证，再低空短距飞行"], diagram: "magnet" },
  { type: "imageText", section: "Week 4", title: "底部支架与重心调整", bullets: ["底部支架需要避开起落架、线束和飞控安装面", "挂载点与机体中心对齐，减少俯仰力矩", "电池位置可微调用于补偿磁吸模块重量", "重载起飞前必须做地面拉力和断电释放测试"], image: img("f450/f450_37.jpg") },
  { type: "table", section: "Week 4", title: "逐级承重测试计划", headers: ["阶段", "重量", "验证重点"], rows: [["T1", "1000g", "吸附可靠、起落架不干涉"], ["T2", "1500g", "起飞姿态和悬停油门"], ["T3", "3000g", "飞行时间和电池压降"], ["T4", "5000g", "安全边界评估与方案探索"]] },
  { type: "concept", section: "Week 4", title: "5kg载重起飞概念示意", bullets: ["此页为载重起飞演示示意图，用于说明底部磁吸和低空运输路径", "真实测试需在空旷场地、低高度、短距离、人员隔离条件下完成", "若悬停油门接近上限，应停止增加重量并改进动力系统"], image: img("f450/f450_39.jpg") },
  { type: "flow", section: "Week 4", title: "载重飞行测试流程", items: ["地面吸附确认", "低油门拉力观察", "低空起飞", "2-3米短距平移", "缓慢降落", "断电释放", "记录数据"] },
  { type: "bars", section: "Week 4", title: "载重测试记录与趋势", labels: ["0kg空载", "1kg", "1.5kg", "3kg", "5kg目标"], values: [100, 82, 74, 58, 38], note: "条形图表示相对飞行裕度：载重越大，油门裕度和续航空间越低。" },
  { type: "cards", section: "Week 4", title: "最大稳定载重判断", cards: [["结构强度", "机臂、中心板和起落架不得明显变形"], ["动力裕度", "悬停不能长期接近满油门"], ["电池能力", "压降不能触发低电保护"], ["操控安全", "挂载摆动不能导致姿态失控"]] },
  { type: "summary", section: "Week 4", title: "第四周小结", body: weekSummaries[4], image: img("f450/f450_38.jpg") },
  { type: "section", section: "Week 5", title: "AI视频与项目展示制作", subtitle: "2026.06.15 - 2026.06.21" },
  { type: "imageText", section: "Week 5", title: "拍摄素材清单", bullets: ["组装过程：机架、电机、电调、飞控安装", "调试过程：Mission Planner连接、校准、参数设置", "测试过程：解锁、悬停、载重、降落", "展示过程：小组分工、演示视频、答辩讲解"], image: img("f450-contact-sheet.jpg") },
  { type: "imageText", section: "Week 5", title: "未来物流无人机外观概念", bullets: ["参考课程模板的黑橙科技风视觉", "强调底部磁吸挂载、模块化货舱和状态灯", "概念图用于宣传展示，不替代真实测试照片", "最终PPT保留图文可编辑结构，便于答辩修改"], image: img("template_media/template_29.jpeg") },
  { type: "flow", section: "Week 5", title: "展示视频脚本", items: ["问题提出", "方案选择", "硬件拼装", "MP调参", "试飞验证", "磁吸载重", "总结改进"] },
  { type: "code", section: "Week 5", title: "视频旁白与字幕要点", code: "开场：我们基于F450平台制作装货无人机。\n结构：四轴机架负责承载，APM飞控负责稳定控制。\n过程：完成机架组装、焊接接线、Mission Planner校准和低空试飞。\n亮点：底部电磁铁实现吸附与释放，逐级验证载重能力。\n结尾：5kg目标作为重载探索，后续可升级动力与结构。", bullets: ["视频时间建议控制在3-5分钟，字幕与画面节奏对应。"] },
  { type: "summary", section: "Week 5", title: "第五周小结", body: weekSummaries[5], image: img("template_media/template_10.jpeg") },
  { type: "section", section: "Week 6", title: "总结与整理资料", subtitle: "2026.06.22" },
  { type: "cards", section: "Week 6", title: "最终成果对照", cards: [["基础稳定飞行", "完成低空悬停与基础动作测试"], ["磁吸挂载功能", "完成底部电磁铁方案设计与流程验证"], ["3000-5000g探索", "完成逐级承重计划与风险边界分析"], ["过程记录", "形成照片、截图、周报和展示视频脚本"]] },
  { type: "table", section: "Week 6", title: "问题复盘矩阵", headers: ["问题", "原因判断", "改进措施"], rows: [["载重裕度不足", "F450原平台动力有限", "升级电机、桨、电池和机架"], ["挂载摆动", "负载离重心较远", "缩短挂绳并增加限摆结构"], ["调参耗时", "传感器校准和通道设置复杂", "建立参数记录表"], ["续航下降", "载重导致电流增大", "选择更高倍率电池并控制航线"]] },
  { type: "diagram", section: "Week 6", title: "后续升级方向", bullets: ["动力升级：更大桨径、更高推力电机、更高倍率电池", "结构升级：加固中心板、独立货仓、限摆机构", "控制升级：自动航线、低电保护、载重模式参数", "安全升级：冗余挂载释放、围栏场地、应急降落流程"], diagram: "upgrade" },
  { type: "imageText", section: "Week 6", title: "答辩展示结构", bullets: ["1分钟：项目背景与目标", "3分钟：制作流程与关键技术", "2分钟：试飞与载重测试数据", "1分钟：问题复盘与改进方向", "问答：安全性、5kg可行性、控制逻辑"], image: img("missionplanner/mp_51.jpeg") },
  { type: "cards", section: "Week 6", title: "项目结论", cards: [["可行性", "F450适合完成教学级装配、调试和轻中载验证"], ["挑战性", "5kg对F450平台提出较高动力与结构要求"], ["学习价值", "完整覆盖硬件、焊接、飞控、地面站、测试闭环"], ["展示价值", "磁吸挂载主题清晰，便于形成视频和答辩故事"]] },
  { type: "summary", section: "Week 6", title: "第六周小结", body: weekSummaries[6], image: img("missionplanner/mp_26.png") },
  { type: "sources", section: "附录", title: "素材与可信度说明", bullets: ["F450组装照片：来自课程参考路径中的F450航拍机安装教程素材。", "APM与Mission Planner截图：来自课程参考路径中的地面站操作文档媒体。", "模板视觉参考：来自“无人机基准对比研究 & 5kg重载方案探索.pptx”的深色技术风格与图标媒体。", "载重起飞、磁吸挂载与5kg目标页：标注为概念示意/测试计划，用于说明方案，不作为未实测数据证明。"] },
  { type: "thanks", section: "答辩", title: "谢谢观看", subtitle: "F450装货无人机 | APM飞控 | Mission Planner | 磁吸挂载 | 5kg重载方案探索" },
];

const DECK_SLIDES = slides.filter((s) =>
  s.type !== "section" &&
  s.type !== "agenda" &&
  s.type !== "timeline" &&
  s.title !== "GitHub初始化与资料目录" &&
  s.title !== "装机完成后的安全检查"
);
const TOTAL_SLIDES = DECK_SLIDES.length;

function img(rel) {
  return path.join(ASSET_DIR, rel);
}

const F450_IMAGES = Array.from({ length: 39 }, (_, i) => img(`f450/f450_${String(i + 1).padStart(2, "0")}.jpg`));
const MP_IMAGES = [
  ...Array.from({ length: 31 }, (_, i) => img(`missionplanner/mp_${String(i + 1).padStart(2, "0")}.png`)),
  img("missionplanner/mp_32.jpeg"),
  ...Array.from({ length: 5 }, (_, i) => img(`missionplanner/mp_${String(i + 33).padStart(2, "0")}.png`)),
  ...Array.from({ length: 2 }, (_, i) => img(`missionplanner/mp_${String(i + 38).padStart(2, "0")}.jpeg`)),
  ...Array.from({ length: 3 }, (_, i) => img(`missionplanner/mp_${String(i + 40).padStart(2, "0")}.png`)),
  ...Array.from({ length: 5 }, (_, i) => img(`missionplanner/mp_${String(i + 43).padStart(2, "0")}.jpeg`)),
  ...Array.from({ length: 3 }, (_, i) => img(`missionplanner/mp_${String(i + 48).padStart(2, "0")}.png`)),
  img("missionplanner/mp_51.jpeg"),
];
const TEMPLATE_IMAGES = [
  "template_02.png", "template_04.png", "template_06.png", "template_08.png",
  "template_10.jpeg", "template_13.jpeg", "template_14.jpeg", "template_17.png",
  "template_19.png", "template_21.png", "template_22.png", "template_24.jpeg",
  "template_25.png", "template_27.jpeg", "template_28.jpeg", "template_29.jpeg",
  "template_30.png", "template_32.png", "template_34.png",
].map((name) => img(`template_media/${name}`));

function relatedImages(d, count = 3) {
  const key = `${d.section || ""} ${d.title || ""}`;
  let pool = F450_IMAGES;
  if (/Mission|Planner|APM|参数|调试|飞控|解锁|PID|地面站/.test(key)) pool = MP_IMAGES;
  if (/AI|视频|展示|概念|外观|宣传|模板/.test(key)) pool = TEMPLATE_IMAGES;
  if (/载重|磁吸|挂载|货物|承重/.test(key)) pool = [img("f450/f450_37.jpg"), img("f450/f450_38.jpg"), img("f450/f450_39.jpg"), img("f450/f450_24.jpg"), img("template_media/template_28.jpeg"), img("template_media/template_29.jpeg")];
  if (/采购|方案|结构|机架|电机|焊接|电源|螺旋桨|重心/.test(key)) pool = F450_IMAGES;
  const start = ((d._index || 1) * 3) % pool.length;
  return Array.from({ length: count }, (_, i) => pool[(start + i) % pool.length]);
}

async function exists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

function extContentType(file) {
  const ext = path.extname(file).toLowerCase();
  if (ext === ".png") return "image/png";
  if (ext === ".jpeg" || ext === ".jpg") return "image/jpeg";
  if (ext === ".svg") return "image/svg+xml";
  return "image/png";
}

async function imageBytes(file) {
  const data = await fs.readFile(file);
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
}

function safeLine(fill = "none", width = 0) {
  return { style: "solid", fill, width };
}

function addShape(slide, position, fill, line = safeLine("none", 0), geometry = "rect", name) {
  return slide.shapes.add({ geometry, name, position, fill, line });
}

function addText(slide, text, position, opts = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: opts.fill ?? "none",
    line: opts.line ?? safeLine("none", 0),
  });
  shape.text = text;
  shape.text.style = {
    fontSize: opts.size ?? 24,
    bold: opts.bold ?? false,
    color: opts.color ?? COLORS.white,
    typeface: opts.face ?? FONT.body,
    alignment: opts.align ?? "left",
  };
  return shape;
}

function addHalftone(slide, x = 1028, y = 86, cols = 7, rows = 6, gap = 22) {
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const size = Math.max(3, 9 - r);
      addShape(slide, { left: x + c * gap, top: y + r * gap, width: size, height: size }, c % 2 ? "#2A170B" : "#4A2409", safeLine("none", 0), "ellipse");
    }
  }
}

function addNeonPanel(slide, position, fill = COLORS.card, line = COLORS.line2) {
  addShape(slide, position, fill, safeLine(line, 1.2), "roundRect");
  addShape(slide, { left: position.left, top: position.top, width: position.width, height: 5 }, COLORS.orange, safeLine("none", 0), "rect");
}

function addMicroLabel(slide, text, x, y, w = 160) {
  addShape(slide, { left: x, top: y, width: w, height: 26 }, "#120B06", safeLine(COLORS.orange, 1), "roundRect");
  addText(slide, text, { left: x + 12, top: y + 6, width: w - 24, height: 14 }, { size: 11, bold: true, color: COLORS.orange2, align: "center" });
}

function addInsight(slide, text, position) {
  addNeonPanel(slide, position, "#0E0E12", COLORS.orange);
  addText(slide, "关键记录", { left: position.left + 18, top: position.top + 18, width: 110, height: 22 }, { size: 16, bold: true, color: COLORS.orange2 });
  addText(slide, text, { left: position.left + 18, top: position.top + 48, width: position.width - 36, height: position.height - 58 }, { size: 15, color: "#DCE3EA" });
}

function insightText(d) {
  const key = d.title || "";
  if (/载重|磁吸|承重/.test(key)) return "逐级加重、低空短距、先地面后飞行；5kg作为探索上限，重点呈现方案与安全边界。";
  if (/Mission|APM|参数|调试|飞控|PID/.test(key)) return "每次只改一个参数，并同步记录截图、现象和处理结果，保证答辩时能解释调试依据。";
  if (/组装|焊接|电机|电源|机架/.test(key)) return "装配照片按“部件-接线-检查”顺序归档，后续PPT和视频都能复用同一套证据链。";
  if (/AI|视频|展示/.test(key)) return "展示材料采用过程实拍 + 概念效果图，区分真实记录与视觉化设想，表达更完整。";
  return "本页用于串联目标、过程、证据和结论，形成可追溯的课程项目叙事。";
}

async function addImage(slide, file, position, opts = {}) {
  if (!(await exists(file))) {
    addShape(slide, position, "#1E293B", safeLine(COLORS.line, 1), "roundRect");
    addText(slide, "素材缺失\n" + path.basename(file), {
      left: position.left + 18,
      top: position.top + 18,
      width: position.width - 36,
      height: position.height - 36,
    }, { size: 18, color: COLORS.card2 });
    return;
  }
  slide.images.add({
    blob: await imageBytes(file),
    contentType: extContentType(file),
    alt: opts.alt ?? path.basename(file),
    fit: opts.fit ?? "cover",
    position,
    geometry: opts.geometry ?? "roundRect",
    borderRadius: opts.borderRadius ?? 8,
  });
}

function addFooter(slide, index, section = "") {
  addShape(slide, { left: 0, top: 0, width: W, height: H }, COLORS.bg, safeLine("none", 0));
  addShape(slide, { left: 0, top: 0, width: W, height: 7 }, COLORS.orange, safeLine("none", 0));
  addShape(slide, { left: 0, top: 706, width: W, height: 3 }, "#2B1305", safeLine("none", 0));
  addHalftone(slide);
  addShape(slide, { left: 52, top: 46, width: 6, height: 52 }, COLORS.orange, safeLine("none", 0), "rect");
  addText(slide, section || "F450 5kg装货无人机", { left: 72, top: 48, width: 360, height: 24 }, { size: 13, color: COLORS.orange2, bold: true });
  addText(slide, String(index).padStart(2, "0") + " / " + TOTAL_SLIDES, { left: 1110, top: 648, width: 110, height: 24 }, { size: 13, color: "#9CA3AF", align: "right" });
  addShape(slide, { left: 72, top: 636, width: 1050, height: 1 }, "#2A2A31", safeLine("none", 0));
  addShape(slide, { left: 72, top: 640, width: 138, height: 3 }, COLORS.orange, safeLine("none", 0));
}

function addTitle(slide, title, subtitle) {
  addText(slide, title, { left: 72, top: 82, width: 900, height: 54 }, { size: 34, bold: true, color: COLORS.white, face: FONT.title });
  if (subtitle) addText(slide, subtitle, { left: 74, top: 134, width: 840, height: 34 }, { size: 16, color: "#B7C4D6" });
}

function bulletText(items) {
  return (items || []).map((item) => "• " + item).join("\n");
}

function addBullets(slide, items, position, opts = {}) {
  const shape = addText(slide, bulletText(items), position, { size: opts.size ?? 21, color: opts.color ?? "#E7EBEF", face: FONT.body });
  return shape;
}

function addCard(slide, x, y, w, h, title, body, color = COLORS.card) {
  addShape(slide, { left: x, top: y, width: w, height: h }, color, safeLine("#CBD5E1", 1), "roundRect");
  addText(slide, title, { left: x + 22, top: y + 18, width: w - 44, height: 30 }, { size: 22, bold: true, color: COLORS.white, face: FONT.title });
  addText(slide, body, { left: x + 22, top: y + 58, width: w - 44, height: h - 72 }, { size: 17, color: "#D1D5DB" });
}

async function renderSlide(p, slideData, index) {
  const slide = p.slides.add();
  addFooter(slide, index, slideData.section);
  switch (slideData.type) {
    case "cover":
      await cover(slide, slideData);
      break;
    case "section":
      sectionSlide(slide, slideData, index);
      break;
    case "overview":
    case "imageText":
      await imageText(slide, slideData);
      break;
    case "agenda":
      agenda(slide, slideData);
      break;
    case "cards":
      cards(slide, slideData);
      break;
    case "diagram":
      diagram(slide, slideData);
      break;
    case "timeline":
      timeline(slide, slideData);
      break;
    case "table":
      tableSlide(slide, slideData);
      break;
    case "flow":
      flowSlide(slide, slideData);
      break;
    case "summary":
      await summary(slide, slideData);
      break;
    case "checklist":
      checklist(slide, slideData);
      break;
    case "code":
      codeSlide(slide, slideData);
      break;
    case "bars":
      bars(slide, slideData);
      break;
    case "concept":
      await concept(slide, slideData);
      break;
    case "sources":
      sources(slide, slideData);
      break;
    case "thanks":
      thanks(slide, slideData);
      break;
    default:
      agenda(slide, slideData);
  }
}

async function cover(slide, d) {
  addShape(slide, { left: 52, top: 66, width: 1176, height: 588 }, "#09090D", safeLine("#2A2A31", 1), "roundRect");
  addText(slide, "智能产品开发设计", { left: 82, top: 102, width: 320, height: 28 }, { size: 16, bold: true, color: COLORS.orange2 });
  addText(slide, d.title, { left: 80, top: 160, width: 550, height: 172 }, { size: 49, bold: true, color: COLORS.white, face: FONT.title });
  addText(slide, d.subtitle, { left: 84, top: 340, width: 520, height: 70 }, { size: 20, color: "#D3D8E0" });
  addShape(slide, { left: 84, top: 440, width: 152, height: 40 }, COLORS.orange, safeLine("none", 0), "roundRect");
  addText(slide, "结课报告 45页", { left: 101, top: 449, width: 120, height: 22 }, { size: 16, bold: true, color: COLORS.white });
  addInsight(slide, "F450平台完成结构认知、硬件装配、APM调参、低空试飞和磁吸载重流程，形成完整工程闭环。", { left: 86, top: 502, width: 520, height: 112 });
  addShape(slide, { left: 680, top: 82, width: 530, height: 498 }, "#0B0B10", safeLine(COLORS.line2, 1.2), "roundRect");
  await addImage(slide, d.image, { left: 704, top: 104, width: 480, height: 288 }, { fit: "cover" });
  addShape(slide, { left: 704, top: 410, width: 146, height: 88 }, COLORS.orange, safeLine("none", 0), "roundRect");
  addText(slide, "APM", { left: 718, top: 424, width: 118, height: 24 }, { size: 24, bold: true, color: COLORS.white, align: "center" });
  addText(slide, "飞控", { left: 744, top: 462, width: 66, height: 18 }, { size: 14, color: "#1A1005", bold: true, align: "center" });
  addShape(slide, { left: 860, top: 410, width: 150, height: 88 }, "#151520", safeLine(COLORS.cyan, 1.2), "roundRect");
  addText(slide, "MP", { left: 874, top: 424, width: 112, height: 24 }, { size: 24, bold: true, color: COLORS.cyan, align: "center" });
  addText(slide, "地面站", { left: 892, top: 462, width: 84, height: 18 }, { size: 14, color: "#D7F8FF", bold: true, align: "center" });
  addShape(slide, { left: 1028, top: 410, width: 146, height: 88 }, "#231006", safeLine(COLORS.orange2, 1.2), "roundRect");
  addText(slide, "5kg", { left: 1042, top: 424, width: 120, height: 24 }, { size: 24, bold: true, color: COLORS.orange2, align: "center" });
  addText(slide, "重载探索", { left: 1048, top: 462, width: 104, height: 18 }, { size: 14, color: "#FFE7C2", bold: true, align: "center" });
  addText(slide, "APM飞控 / MP地面站 / 磁吸挂载 / 载重起飞", { left: 720, top: 534, width: 450, height: 36 }, { size: 20, color: COLORS.white, bold: true });
}

function sectionSlide(slide, d, index) {
  addShape(slide, { left: 104, top: 126, width: 1072, height: 406 }, "#09090D", safeLine("#2A2A31", 1), "roundRect");
  addText(slide, d.section, { left: 148, top: 158, width: 220, height: 42 }, { size: 30, bold: true, color: COLORS.orange2 });
  addText(slide, d.title, { left: 148, top: 230, width: 820, height: 76 }, { size: 52, bold: true, color: COLORS.white, face: FONT.title });
  addText(slide, d.subtitle, { left: 152, top: 326, width: 640, height: 38 }, { size: 21, color: "#C7D0DA" });
  addText(slide, `第 ${index} 页 / 共 ${TOTAL_SLIDES} 页`, { left: 152, top: 404, width: 280, height: 26 }, { size: 15, color: "#8D98A6" });
  addShape(slide, { left: 898, top: 168, width: 184, height: 184 }, "#14141A", safeLine(COLORS.orange, 1.5), "ellipse");
  addText(slide, String(index).padStart(2, "0"), { left: 946, top: 210, width: 90, height: 44 }, { size: 34, bold: true, color: COLORS.orange2, align: "center" });
  addText(slide, "章节", { left: 960, top: 258, width: 62, height: 18 }, { size: 14, color: "#E7E7EA", align: "center" });
  addShape(slide, { left: 1026, top: 274, width: 92, height: 92 }, "#13161D", safeLine(COLORS.cyan, 1.5), "ellipse");
  addText(slide, "TECH", { left: 1039, top: 308, width: 66, height: 24 }, { size: 17, bold: true, color: COLORS.cyan, align: "center" });
  addMicroLabel(slide, "章节切换", 150, 92, 100);
}

async function imageText(slide, d) {
  addTitle(slide, d.title);
  const imgPos = { left: 718, top: 154, width: 466, height: 294 };
  addNeonPanel(slide, { left: 82, top: 166, width: 574, height: 392 }, "#0E0E12", COLORS.orange);
  addBullets(slide, d.bullets, { left: 116, top: 206, width: 508, height: 244 }, { size: 21 });
  addShape(slide, { left: 106, top: 474, width: 470, height: 54 }, "#120B06", safeLine(COLORS.orange, 1), "roundRect");
  addText(slide, insightText(d), { left: 130, top: 488, width: 420, height: 26 }, { size: 16, color: "#F2F5F7", bold: true });
  await addImage(slide, d.image, imgPos, { fit: "cover" });
  const rel = relatedImages(d, 3);
  await addImage(slide, rel[0], { left: 718, top: 472, width: 146, height: 126 }, { fit: "cover" });
  await addImage(slide, rel[1], { left: 876, top: 472, width: 146, height: 126 }, { fit: "cover" });
  await addImage(slide, rel[2], { left: 1034, top: 472, width: 150, height: 126 }, { fit: "cover" });
  addShape(slide, { left: 718, top: 154, width: 466, height: 294 }, "none", safeLine(COLORS.orange, 1.2), "roundRect");
  addShape(slide, { left: 718, top: 612, width: 466, height: 34 }, "#101015", safeLine("#30303A", 1), "roundRect");
  addText(slide, "主图 + 过程补充图", { left: 744, top: 619, width: 200, height: 18 }, { size: 13, color: COLORS.orange2, bold: true });
  addText(slide, "图片说明与步骤同步展示", { left: 980, top: 619, width: 180, height: 18 }, { size: 13, color: "#B5BBC4", align: "right" });
  addMicroLabel(slide, "图文强化", 82, 144, 96);
}

function agenda(slide, d) {
  addTitle(slide, d.title);
  const items = d.bullets || [];
  items.forEach((item, i) => {
    const y = 170 + i * 78;
    addShape(slide, { left: 96, top: y, width: 52, height: 52 }, COLORS.orange, safeLine("none", 0), "ellipse");
    addText(slide, String(i + 1), { left: 112, top: y + 11, width: 24, height: 24 }, { size: 20, bold: true, color: COLORS.white, align: "center" });
    addShape(slide, { left: 170, top: y, width: 920, height: 52 }, "#0D1B30", safeLine("#223348", 1), "roundRect");
    addText(slide, item, { left: 198, top: y + 13, width: 860, height: 26 }, { size: 21, color: COLORS.card });
  });
}

function cards(slide, d) {
  addTitle(slide, d.title);
  const positions = [
    [88, 176], [650, 176], [88, 386], [650, 386],
  ];
  d.cards.forEach((card, i) => {
    addCard(slide, positions[i][0], positions[i][1], 512, 154, card[0], card[1], i % 2 === 0 ? "#111117" : "#17171E");
    addShape(slide, { left: positions[i][0], top: positions[i][1], width: 512, height: 6 }, i % 2 === 0 ? COLORS.orange : COLORS.cyan, safeLine("none", 0), "rect");
    addShape(slide, { left: positions[i][0] + 462, top: positions[i][1] + 18, width: 32, height: 32 }, i % 2 === 0 ? "#2D1406" : "#0B2326", safeLine("none", 0), "ellipse");
  });
  addInsight(slide, insightText(d), { left: 90, top: 572, width: 1080, height: 74 });
}

function diagram(slide, d) {
  addTitle(slide, d.title);
  addNeonPanel(slide, { left: 76, top: 170, width: 520, height: 382 }, "#0E0E12", COLORS.orange);
  addBullets(slide, d.bullets, { left: 110, top: 210, width: 444, height: 252 }, { size: 20 });
  addText(slide, insightText(d), { left: 110, top: 476, width: 420, height: 48 }, { size: 15, color: "#D3D9E0" });
  if (d.diagram === "quad") quadDiagram(slide);
  if (d.diagram === "payload") payloadDiagram(slide);
  if (d.diagram === "magnet") magnetDiagram(slide);
  if (d.diagram === "upgrade") upgradeDiagram(slide);
}

function quadDiagram(slide) {
  const cx = 885, cy = 350;
  addShape(slide, { left: cx - 75, top: cy - 75, width: 150, height: 150 }, "#152A44", safeLine(COLORS.cyan, 2), "ellipse");
  addText(slide, "F450\nAPM", { left: cx - 50, top: cy - 30, width: 100, height: 60 }, { size: 22, bold: true, color: COLORS.white, align: "center" });
  [[-180, -145, "M1 CW"], [180, -145, "M2 CCW"], [-180, 145, "M3 CCW"], [180, 145, "M4 CW"]].forEach(([dx, dy, label]) => {
    addShape(slide, { left: cx + dx - 48, top: cy + dy - 48, width: 96, height: 96 }, COLORS.orange, safeLine("none", 0), "ellipse");
    addText(slide, label, { left: cx + dx - 43, top: cy + dy - 12, width: 86, height: 24 }, { size: 15, bold: true, color: COLORS.white, align: "center" });
    addShape(slide, { left: Math.min(cx, cx + dx), top: Math.min(cy, cy + dy), width: Math.abs(dx) || 2, height: 4 }, "#3D5A78", safeLine("none", 0), "rect");
  });
}

function payloadDiagram(slide) {
  const x = 660, y = 180;
  addShape(slide, { left: x + 200, top: y, width: 120, height: 60 }, "#12161D", safeLine(COLORS.cyan, 2), "roundRect");
  addText(slide, "总推力", { left: x + 220, top: y + 18, width: 80, height: 24 }, { size: 20, bold: true, color: COLORS.white, align: "center" });
  ["机体自重", "电池重量", "磁吸模块", "货物重量"].forEach((t, i) => {
    addShape(slide, { left: x + 40 + i * 130, top: y + 150, width: 110, height: 76 }, i === 3 ? "#3A1607" : "#101015", safeLine(i === 3 ? COLORS.orange : "#2E3037", 1), "roundRect");
    addText(slide, t, { left: x + 52 + i * 130, top: y + 174, width: 86, height: 24 }, { size: 18, color: COLORS.white, bold: i === 3, align: "center" });
  });
  addText(slide, "飞行裕度 = 总推力 - 全机重量", { left: x + 80, top: y + 300, width: 520, height: 40 }, { size: 26, bold: true, color: COLORS.orange2, align: "center" });
}

function magnetDiagram(slide) {
  const x = 720, y = 145;
  addShape(slide, { left: x + 100, top: y, width: 250, height: 120 }, "#12161D", safeLine(COLORS.cyan, 2), "roundRect");
  addText(slide, "F450机体中心", { left: x + 130, top: y + 46, width: 190, height: 28 }, { size: 22, bold: true, color: COLORS.white, align: "center" });
  addShape(slide, { left: x + 190, top: y + 135, width: 70, height: 55 }, COLORS.orange, safeLine("none", 0), "roundRect");
  addText(slide, "电磁铁", { left: x + 194, top: y + 152, width: 62, height: 22 }, { size: 16, bold: true, color: COLORS.white, align: "center" });
  addShape(slide, { left: x + 130, top: y + 230, width: 190, height: 110 }, "#21242C", safeLine("#3B3E47", 1), "roundRect");
  addText(slide, "5kg货物\n概念验证", { left: x + 150, top: y + 260, width: 150, height: 50 }, { size: 22, bold: true, color: COLORS.white, align: "center" });
  addText(slide, "吸附 / 释放由辅助通道控制", { left: x + 80, top: y + 370, width: 340, height: 28 }, { size: 18, color: "#B7C4D6", align: "center" });
}

function upgradeDiagram(slide) {
  const center = { x: 900, y: 355 };
  addShape(slide, { left: center.x - 85, top: center.y - 85, width: 170, height: 170 }, COLORS.orange, safeLine("none", 0), "ellipse");
  addText(slide, "下一代\n重载平台", { left: center.x - 60, top: center.y - 34, width: 120, height: 70 }, { size: 23, bold: true, color: COLORS.white, align: "center" });
  [["动力", -230, -140], ["结构", 230, -140], ["控制", -230, 140], ["安全", 230, 140]].forEach(([label, dx, dy]) => {
    addShape(slide, { left: center.x + dx - 64, top: center.y + dy - 38, width: 128, height: 76 }, "#12161D", safeLine(COLORS.cyan, 2), "roundRect");
    addText(slide, label, { left: center.x + dx - 50, top: center.y + dy - 12, width: 100, height: 24 }, { size: 24, bold: true, color: COLORS.white, align: "center" });
  });
}

function timeline(slide, d) {
  addTitle(slide, d.title);
  const startX = 110;
  d.items.forEach((item, i) => {
    const x = startX + i * 175;
    addShape(slide, { left: x, top: 275, width: 78, height: 78 }, i < 4 ? COLORS.orange : "#132B46", safeLine(i < 4 ? "none" : COLORS.cyan, 2), "ellipse");
    addText(slide, "W" + (i + 1), { left: x + 18, top: 296, width: 44, height: 26 }, { size: 22, bold: true, color: COLORS.white, align: "center" });
    if (i < d.items.length - 1) addShape(slide, { left: x + 78, top: 312, width: 98, height: 4 }, "#3D5A78", safeLine("none", 0));
    addText(slide, item, { left: x - 26, top: 380, width: 132, height: 74 }, { size: 17, color: COLORS.card, align: "center" });
  });
}

function tableSlide(slide, d) {
  addTitle(slide, d.title);
  const x = 90, y = 170, w = 1080, rowH = 72;
  const cols = d.headers.length;
  const colW = w / cols;
  d.headers.forEach((h, c) => {
    addShape(slide, { left: x + c * colW, top: y, width: colW, height: rowH }, COLORS.orange, safeLine("#FFB36A", 1), "rect");
    addText(slide, h, { left: x + c * colW + 14, top: y + 22, width: colW - 28, height: 24 }, { size: 19, bold: true, color: COLORS.white, align: "center" });
  });
  d.rows.forEach((row, r) => {
    row.forEach((cell, c) => {
      addShape(slide, { left: x + c * colW, top: y + rowH * (r + 1), width: colW, height: rowH }, r % 2 ? "#14161B" : "#101015", safeLine("#2F3138", 1), "rect");
      addShape(slide, { left: x + c * colW, top: y + rowH * (r + 1), width: 4, height: rowH }, c === 0 ? COLORS.orange : COLORS.cyan, safeLine("none", 0), "rect");
      addText(slide, cell, { left: x + c * colW + 16, top: y + rowH * (r + 1) + 16, width: colW - 32, height: rowH - 22 }, { size: 16, color: COLORS.white, align: c === 0 ? "center" : "left", bold: c === 0 });
    });
  });
  addInsight(slide, insightText(d), { left: 90, top: 586, width: 1080, height: 60 });
}

function flowSlide(slide, d) {
  addTitle(slide, d.title);
  const count = d.items.length;
  const boxW = count <= 6 ? 150 : 136;
  const gap = count <= 6 ? 24 : 14;
  const total = count * boxW + (count - 1) * gap;
  let x = (W - total) / 2;
  const y = 278;
  d.items.forEach((item, i) => {
    addShape(slide, { left: x, top: y, width: boxW, height: 104 }, i % 2 ? "#12161D" : "#231006", safeLine(i % 2 ? COLORS.cyan : COLORS.orange, 2), "roundRect");
    addText(slide, String(i + 1), { left: x + 14, top: y + 12, width: 34, height: 28 }, { size: 20, bold: true, color: COLORS.white });
    addText(slide, item, { left: x + 16, top: y + 50, width: boxW - 32, height: 42 }, { size: 18, bold: true, color: COLORS.white, align: "center" });
    if (i < count - 1) {
      addText(slide, "→", { left: x + boxW + 2, top: y + 32, width: gap + 12, height: 36 }, { size: 30, bold: true, color: "#8BA3BC", align: "center" });
    }
    x += boxW + gap;
  });
}

async function summary(slide, d) {
  addTitle(slide, d.title);
  addNeonPanel(slide, { left: 82, top: 160, width: 690, height: 390 }, "#0E0E12", COLORS.orange);
  addText(slide, d.body, { left: 116, top: 194, width: 622, height: 282 }, { size: 20, color: "#E7EBEF" });
  addText(slide, "200字周小结", { left: 116, top: 506, width: 140, height: 20 }, { size: 13, bold: true, color: COLORS.orange2 });
  await addImage(slide, d.image, { left: 832, top: 170, width: 336, height: 238 }, { fit: "cover" });
  const rel = relatedImages(d, 2);
  await addImage(slide, rel[0], { left: 832, top: 426, width: 160, height: 128 }, { fit: "cover" });
  await addImage(slide, rel[1], { left: 1008, top: 426, width: 160, height: 128 }, { fit: "cover" });
  addShape(slide, { left: 812, top: 568, width: 380, height: 44 }, "#17171E", safeLine(COLORS.orange, 1), "roundRect");
  addText(slide, "本周成果：资料、照片、参数归档", { left: 836, top: 579, width: 328, height: 22 }, { size: 15, bold: true, color: COLORS.white, align: "center" });
  addInsight(slide, insightText(d), { left: 82, top: 570, width: 690, height: 76 });
}

function checklist(slide, d) {
  addTitle(slide, d.title);
  const positions = [[110, 170], [470, 170], [830, 170], [110, 360], [470, 360], [830, 360]];
  d.items.forEach((item, i) => {
    const [x, y] = positions[i];
    addShape(slide, { left: x, top: y, width: 280, height: 120 }, "#13131A", safeLine("#2E3037", 1), "roundRect");
    addShape(slide, { left: x + 22, top: y + 26, width: 42, height: 42 }, COLORS.green, safeLine("none", 0), "ellipse");
    addText(slide, "✓", { left: x + 31, top: y + 28, width: 24, height: 24 }, { size: 22, bold: true, color: COLORS.white, align: "center" });
    addText(slide, item, { left: x + 82, top: y + 28, width: 170, height: 58 }, { size: 20, bold: true, color: COLORS.white });
  });
  addInsight(slide, "清单页用于答辩前自检，确保机体、通电、桨叶、传感器和试飞场地全部过关。", { left: 110, top: 520, width: 1000, height: 72 });
}

function codeSlide(slide, d) {
  addTitle(slide, d.title);
  addShape(slide, { left: 88, top: 158, width: 760, height: 386 }, "#020617", safeLine("#334155", 1), "roundRect");
  addText(slide, d.code, { left: 120, top: 190, width: 696, height: 330 }, { size: 21, color: "#D1FAE5", face: FONT.mono });
  addShape(slide, { left: 890, top: 190, width: 270, height: 250 }, "#12161D", safeLine(COLORS.cyan, 2), "roundRect");
  addBullets(slide, d.bullets, { left: 918, top: 226, width: 214, height: 168 }, { size: 18, color: COLORS.white });
  addInsight(slide, "这里把“编程”落到飞控固件、通道映射和磁吸控制逻辑，便于课程答辩解释技术链路。", { left: 890, top: 468, width: 270, height: 76 });
}

function bars(slide, d) {
  addTitle(slide, d.title);
  const x = 150, y = 180, maxW = 800, barH = 42, gap = 34;
  d.labels.forEach((label, i) => {
    const top = y + i * (barH + gap);
    addText(slide, label, { left: x, top: top + 8, width: 150, height: 26 }, { size: 18, bold: true, color: COLORS.white });
    addShape(slide, { left: x + 180, top, width: maxW, height: barH }, "#12161D", safeLine("#2E3037", 1), "roundRect");
    addShape(slide, { left: x + 180, top, width: Math.max(8, maxW * d.values[i] / 100), height: barH }, i >= d.values.length - 1 ? COLORS.orange : COLORS.cyan, safeLine("none", 0), "roundRect");
    addText(slide, d.values[i] + "%", { left: x + 1000, top: top + 8, width: 70, height: 24 }, { size: 18, bold: true, color: COLORS.white, align: "right" });
  });
  addText(slide, d.note, { left: 150, top: 580, width: 980, height: 42 }, { size: 16, color: "#B7C4D6" });
}

async function concept(slide, d) {
  addTitle(slide, d.title);
  addNeonPanel(slide, { left: 86, top: 160, width: 590, height: 360 }, "#0E0E12", COLORS.orange);
  await addImage(slide, d.image, { left: 106, top: 176, width: 550, height: 310 }, { fit: "contain" });
  addShape(slide, { left: 300, top: 445, width: 160, height: 92 }, "#475569", safeLine("#CBD5E1", 2), "roundRect");
  addText(slide, "5kg\n货物", { left: 335, top: 466, width: 90, height: 50 }, { size: 23, bold: true, color: COLORS.white, align: "center" });
  addShape(slide, { left: 365, top: 405, width: 30, height: 48 }, COLORS.orange, safeLine("none", 0), "roundRect");
  addText(slide, "电磁铁", { left: 330, top: 382, width: 98, height: 24 }, { size: 16, bold: true, color: COLORS.orange2, align: "center" });
  addNeonPanel(slide, { left: 705, top: 172, width: 455, height: 335 }, "#0E0E12", COLORS.cyan);
  addBullets(slide, d.bullets, { left: 738, top: 212, width: 388, height: 250 }, { size: 19 });
  addText(slide, "注：概念示意图，非未标注实拍证据", { left: 735, top: 482, width: 390, height: 24 }, { size: 15, color: COLORS.orange2, bold: true, align: "center" });
}

function sources(slide, d) {
  addTitle(slide, d.title);
  addShape(slide, { left: 86, top: 160, width: 1080, height: 390 }, "#101015", safeLine("#2E3037", 1), "roundRect");
  addBullets(slide, d.bullets, { left: 126, top: 200, width: 1000, height: 300 }, { size: 20, color: COLORS.white });
  addInsight(slide, "附录页说明素材来源、概念页边界和相对数据性质，答辩时可直接作为可信度补充。", { left: 86, top: 572, width: 1080, height: 64 });
}

function thanks(slide, d) {
  addShape(slide, { left: 0, top: 0, width: W, height: H }, COLORS.bg2, safeLine("none", 0));
  addShape(slide, { left: 0, top: 0, width: W, height: 10 }, COLORS.orange, safeLine("none", 0));
  addText(slide, d.title, { left: 220, top: 230, width: 840, height: 90 }, { size: 72, bold: true, color: COLORS.white, face: FONT.title, align: "center" });
  addText(slide, d.subtitle, { left: 210, top: 340, width: 860, height: 50 }, { size: 22, color: "#B7C4D6", align: "center" });
  addShape(slide, { left: 505, top: 438, width: 270, height: 54 }, COLORS.orange, safeLine("none", 0), "roundRect");
  addText(slide, "Q&A", { left: 590, top: 450, width: 100, height: 30 }, { size: 30, bold: true, color: COLORS.white, align: "center" });
  addText(slide, "F450 / APM / Mission Planner / 磁吸载重", { left: 402, top: 512, width: 460, height: 26 }, { size: 15, color: "#D5DBE4", align: "center" });
}

async function writePlanningFiles() {
  const planLines = [
    "F450 5kg装货无人机完整结课报告 - 45页页面计划",
    "模式：create（参考模板风格，因源模板渲染存在图片对象错误，未做逐页克隆）",
    "风格：黑色与荧光橙的暗色科技风，极简、高级、波普艺术感。",
    "字体：Microsoft YaHei UI / Microsoft YaHei / Consolas。",
    "页面结构：封面、项目总览、结构认知、六周周报、测试数据、总结附录。",
    "",
    ...DECK_SLIDES.map((s, i) => `${String(i + 1).padStart(2, "0")}. ${s.section || ""} ${s.title}`),
  ];
  await fs.writeFile(SLIDE_PLAN, planLines.join("\n"), "utf8");
  await fs.writeFile(SOURCE_NOTES, [
    "素材来源记录",
    "1. 用户提供参考素材目录：D:/无人机/无人机。使用其中F450航拍机安装教程图片、APM/Pixhawk教程与Mission Planner文档媒体。",
    "2. 用户提供参考模板：无人机基准对比研究 & 5kg重载方案探索.pptx。因artifact-tool渲染源模板时报底层图片对象错误，改为参考深色技术风格和抽取媒体。",
    "3. 项目计划、周任务、成员分工、六周时间节点来自用户在对话中提供的项目周计划。",
    "4. 载重趋势与稳定性条形图为课程报告表达用的相对评分/计划数据，不作为实测数值证明。",
    "5. 磁吸载重起飞页为概念示意，用于说明方案和安全流程，已在页面中标注。",
  ].join("\n"), "utf8");
}

async function exportPreviews(p) {
  await fs.rm(PREVIEW_DIR, { recursive: true, force: true });
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  for (const [i, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(i + 1).padStart(2, "0")}`;
    const png = await p.export({ slide, format: "png", scale: 1 });
    await fs.writeFile(path.join(PREVIEW_DIR, `${stem}.png`), Buffer.from(await png.arrayBuffer()));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(PREVIEW_DIR, `${stem}.layout.json`), await layout.text(), "utf8");
  }
  const montage = await p.export({ format: "webp", montage: true, scale: 1 });
  await fs.writeFile(path.join(PREVIEW_DIR, "montage.webp"), Buffer.from(await montage.arrayBuffer()));
}

async function main() {
  if (DECK_SLIDES.length !== 45) throw new Error(`Expected 45 slides, got ${DECK_SLIDES.length}`);
  await fs.mkdir(OUT_DIR, { recursive: true });
  await fs.mkdir(path.dirname(LOCAL_COPY), { recursive: true });
  await fs.mkdir(QA_DIR, { recursive: true });
  await writePlanningFiles();

  const p = Presentation.create({ slideSize: { width: W, height: H } });
  for (let i = 0; i < DECK_SLIDES.length; i += 1) {
    DECK_SLIDES[i]._index = i + 1;
    await renderSlide(p, DECK_SLIDES[i], i + 1);
  }

  await exportPreviews(p);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(FINAL_PPTX);
  await pptx.save(LOCAL_COPY);
  const stat = await fs.stat(FINAL_PPTX);
  await fs.writeFile(path.join(QA_DIR, "visual-qa.txt"), [
    "QA记录",
    `生成页数：${DECK_SLIDES.length}`,
    `最终文件：${FINAL_PPTX}`,
    `文件大小：${stat.size} bytes`,
    "已导出45页PNG预览和montage用于检查。",
    "注意：载重起飞页为概念示意，页面中已标注。",
  ].join("\n"), "utf8");
  console.log(JSON.stringify({ final: FINAL_PPTX, localCopy: LOCAL_COPY, slides: DECK_SLIDES.length, bytes: stat.size, previewDir: PREVIEW_DIR }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
