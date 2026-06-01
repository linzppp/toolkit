'use strict';
 
const fs   = require('fs');
const path = require('path');
const readline = require('readline');
 
// ──────────────────────────────────────────
// 参数解析
// ──────────────────────────────────────────
const args = process.argv.slice(2);
if (args.length === 0 || args[0] === '--help' || args[0] === '-h') {
  console.log('用法: node tsv_to_csv.js <input.tsv> [output.csv]');
  process.exit(0);
}
 
const inputFile  = args[0];
const outputFile = args[1] || inputFile.replace(/\.tsv$/i, '') + '.csv';
 
if (!fs.existsSync(inputFile)) {
  console.error(`错误: 找不到输入文件 "${inputFile}"`);
  process.exit(1);
}
 
if (path.resolve(inputFile) === path.resolve(outputFile)) {
  console.error('错误: 输入和输出文件不能相同');
  process.exit(1);
}
 
// ──────────────────────────────────────────
// CSV 单元格转义
// ──────────────────────────────────────────
/**
 * 将单个字段值转义为合法的 CSV 字段。
 * 规则（RFC 4180）：
 *   - 若字段包含 逗号、双引号、换行符 → 用双引号包裹
 *   - 字段内的双引号 → 转义为 ""
 */
function escapeCSVField(value) {
  // 包含特殊字符才需要处理
  if (value.includes('"') || value.includes(',') || value.includes('\n') || value.includes('\r')) {
    return '"' + value.replace(/"/g, '""') + '"';
  }
  return value;
}
 
/**
 * 将一行 TSV（已按 \t 分割的字段数组）转换为 CSV 行字符串。
 */
function tsvFieldsToCSVLine(fields) {
  return fields.map(escapeCSVField).join(',');
}
 
// ──────────────────────────────────────────
// 主流程
// ──────────────────────────────────────────
const startTime = Date.now();
let lineCount   = 0;
let errorCount  = 0;
 
console.log(`输入文件: ${inputFile}`);
console.log(`输出文件: ${outputFile}`);
 
// 获取文件大小，用于估算进度
const fileStat   = fs.statSync(inputFile);
const fileSizeMB = (fileStat.size / 1024 / 1024).toFixed(1);
console.log(`文件大小: ${fileSizeMB} MB`);
console.log('开始转换...\n');
 
const readStream  = fs.createReadStream(inputFile, { encoding: 'utf8' });
const writeStream = fs.createWriteStream(outputFile, { encoding: 'utf8' });
 
// 监听写入错误
writeStream.on('error', (err) => {
  console.error(`写入错误: ${err.message}`);
  process.exit(1);
});
 
// readline 逐行读取（自动处理 \r\n 和 \n）
const rl = readline.createInterface({
  input:      readStream,
  crlfDelay: Infinity,   // 正确处理 Windows 换行符
});
 
rl.on('line', (line) => {
  lineCount++;
 
  try {
    // TSV 按制表符分割
    const fields  = line.split('\t');
    const csvLine = tsvFieldsToCSVLine(fields);
    writeStream.write(csvLine + '\n');
  } catch (err) {
    errorCount++;
    console.error(`第 ${lineCount} 行处理出错: ${err.message}`);
  }
 
  // 进度提示
  if (lineCount % 100_000 === 0) {
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`已处理 ${lineCount.toLocaleString()} 行 (${elapsed}s)`);
  }
});
 
rl.on('close', () => {
  // 等待写入流完全刷新后再退出
  writeStream.end(() => {
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    const outStat = fs.statSync(outputFile);
    const outMB   = (outStat.size / 1024 / 1024).toFixed(1);
 
    console.log(`\n✅ 转换完成`);
    console.log(`   总行数   : ${lineCount.toLocaleString()}`);
    console.log(`   错误行数 : ${errorCount}`);
    console.log(`   耗时     : ${elapsed}s`);
    console.log(`   输出大小 : ${outMB} MB → ${outputFile}`);
  });
});
 
rl.on('error', (err) => {
  console.error(`读取错误: ${err.message}`);
  process.exit(1);
});
 