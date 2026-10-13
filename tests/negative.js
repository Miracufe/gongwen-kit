#!/usr/bin/env node
/**
 * negative.js —— 负对照：每道机检造一份「必错」的稿，断言它报出且非零退出。
 *
 * 为什么必须有它：检查器在空输入、坏输入上会逐条通过（「检查器在空输入上
 * 会逐条通过」是交付验证的经典静默故障）。断言没反向测过不算数——这份
 * 测试就是给每条规则一具「必错的标本」，报不出来就算检查失效。
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const docx = require("docx");
const G = require("../lib/gongwen-core").form("wen");

const ROOT = path.join(__dirname, "..");
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "gwneg-"));
let failed = 0;

async function build(name, children, props) {
  const doc = new docx.Document({
    externalStyles: G.externalStyles(),
    sections: [{
      properties: props || {
        page: { size: { width: 11906, height: 16838 }, margin: G.MARGIN },
        grid: G.GRID,
      },
      children,
    }],
  });
  const out = path.join(tmp, name);
  fs.writeFileSync(out, await docx.Packer.toBuffer(doc));
  return out;
}

const P = (s, o = {}) => new docx.Paragraph({ children: [new docx.TextRun(s)], ...o });

function expectFail(checker, file, note) {
  const r = spawnSync(process.execPath, [path.join(ROOT, "checks", checker), file, "--strict"], { encoding: "utf8" });
  const out = (r.stdout || "") + (r.stderr || "");
  if (r.status === 0 || r.status === null) {
    failed += 1;
    console.log(`  ✗ ${checker} 没报出「${note}」（负对照失败）\n${out.trim()}`);
    return;
  }
  const first = out.split("\n").find((l) => l.includes("·")) || out.split("\n")[0];
  console.log(`  ✓ ${checker} 报出「${note}」：${String(first).trim().slice(0, 64)}`);
}

(async () => {
  // 版心被改（GB/T 9704 8.1）
  const wrongMargin = {
    page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1587, right: 1474 } },
    grid: G.GRID,
  };
  expectFail("check-doc.js", await build("版心错.docx", [P("测试正文文字。")], wrongMargin), "版心不合");

  // 层次序数用 1.1 式
  expectFail("check-doc.js", await build("序数错.docx", [P("1.1 这应写成 一、")]), "1.1 式编号");

  // 中文数字日期
  expectFail("check-doc.js", await build("日期错.docx", [P("二〇二六年十月十二日")]), "中文数字日期");

  // 月日编虚位
  expectFail("check-doc.js", await build("虚位错.docx", [P("会议定于 2026 年 05 月 06 日召开。")]), "月日编虚位");

  // 发文字号括号与虚位
  expectFail("check-doc.js", await build("文号错.docx", [P("×政发【2026】012号")]), "发文字号格式");

  // 半角括号
  expectFail("check-doc.js", await build("括号错.docx", [P("这是(半角)括号。")]), "半角括号");

  // 异形词（规范层，GB 1001 来自 data/yixingci.tsv）
  expectFail("check-style.js", await build("异形词.docx", [P("他的笔划工整。")]), "异形词");

  // 中文里的半角标点（规范层）
  expectFail("check-style.js", await build("半角标点.docx", [P("句子中间,夹了半角逗号。")]), "半角标点");

  // 表格合计不符
  const t = G.table({ head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "10"], ["乙", "20"], ["合计", "99"]] });
  expectFail("check-num.js", await build("合计错.docx", [...t]), "合计不符");

  console.log(failed ? `\n负对照失败 ${failed} 项` : "\n负对照全部报出");
  process.exit(failed ? 1 : 0);
})().catch((e) => {
  console.error("负对照跑不起来：", e);
  process.exit(2);
});
