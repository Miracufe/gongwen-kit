#!/usr/bin/env node
/**
 * check-num.js —— 数据表里的数字自洽机检。
 *
 * 用法：node check-num.js <文件.docx> [--strict]
 *       稿源脚本由 writeDoc 写盘后自动调用；--strict 时非零退出。
 *
 * 为什么有它：check-doc 只查格式（单位空格、小数 0、日期写法），不查数字本身
 * 对不对。表内合计、百分比构成、同一指标在两张表里的取值——全靠手算，改一个
 * 参数要人工复算全部关联表。这条把它机检化。
 *
 * **只收无歧义的三类，宁漏勿滥**（检查项自己制造假阳性比没有检查更糟）：
 *   1. 合计行／合计列：标了「合计／总计／小计」的行列，等于其余数值格之和
 *   2. 百分比构成：标了百分号的合计，且该行／列是占比，和应为 100
 *   3. 同指标跨表同值：同一文字标签在两张数据表里出现不同数值
 *
 * 一格不是干净数字（空格、中文单位、括号负数、破折号）就不参与这一列，不当 0，
 * 也不报错——那是 check-doc 与人工的活。数值容差见 TOL。
 *
 * 合并格（w:gridSpan）按真实列位展开——读取在 table-read.js，与 check-doc 共用。
 * 读取走 jszip，不调 unzip（Windows 上没有后者）。
 */
"use strict";
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { textOf, rowCells } = require("../lib/table-read");

// 相对容差 0.5%，绝对下限 0.02——印刷稿的舍入误差不会比这更大。
const TOL_REL = 0.005;
const TOL_ABS = 0.02;
const PCT_TOL = 0.5;          // 百分比构成和与 100 的容差
const TOTAL_LABEL = /^(合计|总计|小计|合\s*计|总\s*计|小\s*计)$/;

async function readTables(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file("word/document.xml").async("string");
  return (xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) || [])
    .filter((tb) => /<w:tblHeader/.test(tb))
    .map((tb) => {
      const trs = tb.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
      const rows = trs.map((tr) => rowCells(tr, textOf));
      // 表尾的无边框定高空行不是数据行（展开后是一串 null／空串）
      if (rows.length && rows[rows.length - 1].every((c) => !c)) rows.pop();
      return rows;
    })
    .filter((rows) => rows.length >= 2);
}

// 干净数字才返回对象：可带千分位、正负号、百分号。带中文单位、括号、范围一律 null。
function num(s) {
  if (typeof s !== "string") return null;
  const t = s.replace(/[,\s]/g, "");
  const m = t.match(/^(-?\d+(?:\.\d+)?)([%％]?)$/);
  if (!m) return null;
  const decimals = m[1].includes(".") ? m[1].split(".")[1].length : 0;
  return { v: Number(m[1]), pct: !!m[2], decimals, raw: s };
}

const round = (n) => Math.round(n * 1000) / 1000;
const close = (a, b) => Math.abs(a - b) <= Math.max(TOL_ABS, TOL_REL * Math.abs(b));

async function check(file) {
  const tables = await readTables(file);
  const hits = [];
  const add = (rule, detail) => hits.push({ rule, detail });

  tables.forEach((rows, ti) => {
    const head = rows[0];

    // ── 1 合计行 ──────────────────────────────────────────
    rows.forEach((cells, ri) => {
      if (ri === 0) return;
      const labelIdx = cells.findIndex((c) => TOTAL_LABEL.test(c));
      if (labelIdx < 0) return;
      for (let ci = 0; ci < cells.length; ci++) {
        if (ci === labelIdx) continue;
        const total = num(cells[ci]);
        if (!total) continue;
        // 该列其余数据行都必须是干净数字，才判得了
        const vals = [];
        let clean = true;
        for (let r = 1; r < rows.length; r++) {
          if (r === ri) continue;
          const n = num(rows[r][ci]);
          if (!n) { clean = false; break; }
          vals.push(n.v);
        }
        if (!clean || !vals.length) continue;
        const sum = vals.reduce((a, b) => a + b, 0);
        if (!close(sum, total.v)) {
          add("合计不符", `第 ${ti + 1} 表「${cells[labelIdx]}」行第 ${ci + 1} 列：列内 ${vals.length} 项之和 ${round(sum)} ≠ 表内合计 ${total.raw}`);
        } else if (total.pct && Math.abs(sum - 100) > PCT_TOL) {
          add("百分比构成", `第 ${ti + 1} 表「${cells[labelIdx]}」行第 ${ci + 1} 列：占比合计 ${round(sum)}%，应为 100%`);
        }
      }
    });

    // ── 2 合计列（表头标了合计的那一列，等于本行其余列之和）────
    head.forEach((h, ci) => {
      if (!TOTAL_LABEL.test(h)) return;
      for (let r = 1; r < rows.length; r++) {
        // 合计行已在第 1 段查过，别从列的方向再报一遍
        if (TOTAL_LABEL.test(rows[r][0] || "")) continue;
        const total = num(rows[r][ci]);
        if (!total) continue;
        const vals = [];
        let clean = true;
        for (let k = 1; k < head.length; k++) {
          if (k === ci) continue;
          const n = num(rows[r][k]);
          if (!n) { clean = false; break; }
          vals.push(n.v);
        }
        if (!clean || !vals.length) continue;
        const sum = vals.reduce((a, b) => a + b, 0);
        if (!close(sum, total.v)) {
          add("合计不符", `第 ${ti + 1} 表第 ${r + 1} 行「${head[ci]}」列：其余列之和 ${round(sum)} ≠ ${total.raw}`);
        }
      }
    });
  });

  // ── 3 同指标跨表同值 ────────────────────────────────────
  // **只在表头完全相同时才比**：历史表与预测表都有「营业收入」，但年份区间不同，
  // 数值本就不同——拿整行比会满屏误报。表头一致还出现同标签不同值，才是真的对不上。
  const seen = new Map();
  tables.forEach((rows, ti) => {
    const head = rows[0];
    const schema = head.join("｜");
    rows.slice(1).forEach((cells) => {
      const label = (cells[0] || "").trim();
      if (!label || label.length < 2 || TOTAL_LABEL.test(label)) return;
      if (!/[\u4e00-\u9fa5]/.test(label)) return;
      const vals = cells.slice(1).map(num).filter(Boolean).map((n) => n.v);
      if (!vals.length) return;
      const key = schema + "::" + label;
      const sig = vals.map(round).join("｜");
      if (seen.has(key)) {
        const prev = seen.get(key);
        if (prev.sig !== sig) {
          add("跨表同值", `「${label}」（表头 ${head.join("/")}）在第 ${prev.ti + 1} 表是 ${prev.sig}，第 ${ti + 1} 表是 ${sig}`);
        }
      } else {
        seen.set(key, { ti, sig });
      }
    });
  });

  return hits;
}

// ── 入口 ─────────────────────────────────────────────────
(async () => {
  const file = process.argv[2];
  if (!file) {
    console.error("用法：node check-num.js <文件.docx> [--strict]");
    process.exit(2);
  }
  const strict = process.argv.includes("--strict");
  let hits;
  try {
    hits = await check(file);
  } catch (e) {
    console.error("机检读不进这份文件（工具故障，不是稿子的问题）：", e.message);
    process.exit(2);
  }
  if (!hits.length) {
    console.log(`数字自洽通过　${path.basename(file)}`);
  } else {
    console.log(`数字自检 ${hits.length} 处　${path.basename(file)}`);
    hits.forEach((h) => console.log(`  · ${h.rule}：${h.detail}`));
    if (strict) process.exit(1);
  }
})();
