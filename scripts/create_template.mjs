import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPaths = {
  template: path.join(rootDir, "templates", "秋招个人资料库模板.xlsx"),
  example: path.join(rootDir, "examples", "示例资料库.xlsx"),
};

const personalHeaders = [
  "姓名", "手机号码", "邮箱", "出生日期", "籍贯", "现居城市", "学校名称", "专业名称",
  "学历", "学位", "入学时间", "预计毕业时间", "期望岗位方向", "期望工作城市",
];
const experienceHeaders = [
  "类别", "项目/单位", "角色/级别", "地点", "开始/考试时间", "结束/颁发时间",
  "成绩/编号", "详细内容", "备注",
];
const attachmentHeaders = ["文件类型", "本地文件路径", "说明"];
const experienceCategories = [
  "教育经历", "实习经历", "工作经历", "项目经历", "语言能力", "专业证书",
  "竞赛获奖", "荣誉称号", "校园活动", "科研成果",
];

const headerFormat = {
  fill: "#1F4E78",
  font: { name: "Aptos", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
};
const inputFormat = {
  fill: "#FFF2CC",
  font: { name: "Aptos", size: 10, color: "#222222" },
  verticalAlignment: "center",
};

function formatHeader(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.format = headerFormat;
  range.format.rowHeight = 32;
}

function setSheetDefaults(sheet) {
  sheet.showGridLines = false;
  sheet.tabColor = "#5B9BD5";
}

function setWidths(sheet, widths) {
  widths.forEach((width, column) => {
    sheet.getRangeByIndexes(0, column, 1, 1).format.columnWidth = width;
  });
}

function addCategoryValidation(sheet) {
  sheet.getRange("A2:A501").dataValidation = {
    rule: { type: "list", values: experienceCategories },
  };
}

function addGuidance(sheet) {
  setSheetDefaults(sheet);
  sheet.tabColor = "#7F8C8D";
  sheet.getRange("A1").values = [["秋招个人资料库使用说明"]];
  sheet.getRange("A1").format = {
    font: { name: "Aptos", size: 14, bold: true, color: "#1F1F1F" },
  };
  sheet.getRange("A3:B8").values = [
    ["步骤", "说明"],
    ["1", "在“个人资料”第 2 行填写基本信息。"],
    ["2", "在“经历清单”中逐行记录教育、实习、项目和证书等经历。"],
    ["3", "“类别”列请从下拉列表中选择，以便后续表单自动匹配。"],
    ["4", "在“附件清单”记录需要上传的文件及其本地位置。"],
    ["提示", "请仅填写本人确认且可用于招聘申请的信息。"],
  ];
  formatHeader(sheet, "A3:B3");
  sheet.getRange("A4:B8").format = {
    font: { name: "Aptos", size: 10, color: "#222222" },
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.getRange("A4:B8").format.borders = {
    preset: "outside", style: "thin", color: "#D9E2F3",
  };
  sheet.getRange("B4:B8").format.rowHeight = 28;
  setWidths(sheet, [12, 72]);
}

function addPersonalProfile(sheet, values) {
  setSheetDefaults(sheet);
  sheet.getRange("A1:N1").values = [personalHeaders];
  sheet.getRange("A2:N2").values = [values];
  formatHeader(sheet, "A1:N1");
  sheet.getRange("A2:N2").format = inputFormat;
  sheet.getRange("D2").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("K2:L2").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("A1:N2").format.borders = {
    preset: "outside", style: "thin", color: "#9EADBA",
  };
  sheet.freezePanes.freezeRows(1);
  setWidths(sheet, [16, 16, 28, 14, 14, 14, 22, 22, 12, 16, 14, 16, 20, 18]);
}

function addExperienceList(sheet, rows) {
  setSheetDefaults(sheet);
  sheet.getRange("A1:I1").values = [experienceHeaders];
  formatHeader(sheet, "A1:I1");
  if (rows.length > 0) {
    sheet.getRangeByIndexes(1, 0, rows.length, experienceHeaders.length).values = rows;
    sheet.getRange(`E2:F${rows.length + 1}`).setNumberFormat("yyyy-mm-dd");
  }
  sheet.getRange("A2:I501").format = inputFormat;
  sheet.getRange("H2:H501").format.wrapText = true;
  addCategoryValidation(sheet);
  sheet.freezePanes.freezeRows(1);
  setWidths(sheet, [16, 28, 18, 16, 16, 16, 18, 52, 28]);
}

function addAttachmentList(sheet, rows) {
  setSheetDefaults(sheet);
  sheet.getRange("A1:C1").values = [attachmentHeaders];
  formatHeader(sheet, "A1:C1");
  if (rows.length > 0) {
    sheet.getRangeByIndexes(1, 0, rows.length, attachmentHeaders.length).values = rows;
  }
  sheet.getRange("A2:C501").format = inputFormat;
  sheet.getRange("C2:C501").format.wrapText = true;
  sheet.freezePanes.freezeRows(1);
  setWidths(sheet, [18, 48, 40]);
}

function createWorkbook({ personalValues, experienceRows, attachmentRows }) {
  const workbook = Workbook.create();
  const guidance = workbook.worksheets.add("使用说明");
  const personal = workbook.worksheets.add("个人资料");
  const experience = workbook.worksheets.add("经历清单");
  const attachments = workbook.worksheets.add("附件清单");

  addGuidance(guidance);
  addPersonalProfile(personal, personalValues);
  addExperienceList(experience, experienceRows);
  addAttachmentList(attachments, attachmentRows);
  return workbook;
}

async function saveWorkbook(workbook, outputPath) {
  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await output.save(outputPath);
}

const blankValues = Array(personalHeaders.length).fill(null);
const fictionalPersonalValues = [
  "林知夏（虚构）", "13800000000", "lin.zhixia@example.invalid", new Date("2003-04-12"),
  "星海市", "云岚市", "海湾虚构大学", "信息管理与信息系统", "本科", "管理学学士",
  new Date("2021-09-01"), new Date("2025-06-30"), "产品运营", "云岚市",
];
const fictionalExperienceRows = [
  ["教育经历", "海湾虚构大学", "信息管理与信息系统本科", "云岚市", new Date("2021-09-01"), new Date("2025-06-30"), "GPA 3.7/4.0", "完成信息系统、数据分析与用户研究课程。", "虚构示例"],
  ["实习经历", "星环数据工作室", "产品运营实习生", "云岚市", new Date("2024-01-08"), new Date("2024-04-30"), "项目复盘 4 份", "协助整理用户反馈并维护活动内容排期。", "虚构示例"],
  ["项目经历", "校园服务地图", "项目负责人", "云岚市", new Date("2023-10-01"), new Date("2023-12-20"), "校内课程项目", "组织四人小组完成需求访谈、原型和测试总结。", "虚构示例"],
  ["语言能力", "CET-4", "大学英语四级", "云岚市", new Date("2022-06-11"), null, "560", "虚构语言成绩示例。", "虚构示例"],
  ["语言能力", "CET-6", "大学英语六级", "云岚市", new Date("2023-06-17"), null, "530", "虚构语言成绩示例。", "虚构示例"],
  ["语言能力", "IELTS", "国际英语语言测试", "云岚市", new Date("2024-08-03"), null, "6.5", "虚构语言成绩示例。", "虚构示例"],
  ["专业证书", "数据分析基础证书", "课程结业", "云岚市", new Date("2024-05-10"), new Date("2024-05-10"), "DEMO-2024-0510", "虚构证书记录示例。", "虚构示例"],
];
const fictionalAttachmentRows = [
  ["简历", "示例文件/虚构简历.pdf", "仅用于演示字段格式，不对应真实文件。"],
  ["成绩单", "示例文件/虚构成绩单.pdf", "仅用于演示字段格式，不对应真实文件。"],
];

await saveWorkbook(
  createWorkbook({ personalValues: blankValues, experienceRows: [], attachmentRows: [] }),
  outputPaths.template,
);
await saveWorkbook(
  createWorkbook({
    personalValues: fictionalPersonalValues,
    experienceRows: fictionalExperienceRows,
    attachmentRows: fictionalAttachmentRows,
  }),
  outputPaths.example,
);
