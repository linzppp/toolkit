'use strict';
 
const fs      = require('fs');
const mysql   = require('mysql2/promise');
const { parse } = require('csv-parse');
 
// ═══════════════════════════════════════════════════════════
//  ✏️  CONFIG — 按需修改这里
// ═══════════════════════════════════════════════════════════
 
const CONFIG = {
  // ── CSV 文件 ──────────────────────────────────────────────
  csvFile: process.argv[2] || 'data.csv',
 
  // ── MySQL 连接 ────────────────────────────────────────────
  db: {
    host:     'localhost',
    port:      3306,
    user:     'root',
    password: 'root123456',
    database: 'roc',
  },
 
  // ── 目标表名 ──────────────────────────────────────────────
  tableName: 'title_akas',
 
  // ── 列映射: CSV列名(表头) → MySQL列名
  //    设为 null 则直接用 CSV 表头作为列名（需表头与列名完全一致）
  //    设为对象则按映射关系导入（可过滤不需要的列）
//   columnMap: null,
  // columnMap 示例:
  columnMap: {
    'titleId': 'titleId',
    'ordering': 'ordering',
    'title': 'title',
    'region': 'region',
    'language': 'language',
    'types': 'types',
    'attributes': 'attributes',
    'isOriginalTitle': 'isOriginalTitle',
},
 
  // ── 批量写入大小（每批插入多少行）────────────────────────
  //    推荐 500~2000，越大速度越快但内存占用越高
  batchSize: 3000,
 
  // ── 若遇到重复主键: 'ignore' | 'replace' | 'error'
  onDuplicate: 'ignore',
 
  // ── 是否在导入前自动建表（根据第一行数据推断类型）────────
  //    生产环境建议关闭，手动建表以控制字段类型
  autoCreateTable: false,
};
 
// ═══════════════════════════════════════════════════════════
//  内部实现（通常不需要改动）
// ═══════════════════════════════════════════════════════════
 
if (!fs.existsSync(CONFIG.csvFile)) {
  console.error(`❌ 找不到文件: ${CONFIG.csvFile}`);
  process.exit(1);
}
 
/** 根据数据样本猜测 MySQL 类型（仅 autoCreateTable 时使用） */
function guessType(value) {
  if (value === null || value === '') return 'TEXT';
  if (/^\d+$/.test(value) && value.length < 12) return 'BIGINT';
  if (/^\d+\.\d+$/.test(value)) return 'DOUBLE';
  if (/^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?$/.test(value)) return 'DATETIME';
  if (value.length <= 255) return 'VARCHAR(255)';
  return 'TEXT';
}
 
/** 构造 INSERT 语句前缀 */
function buildInsertPrefix(tableName, columns, onDuplicate) {
  const keyword = onDuplicate === 'ignore'  ? 'INSERT IGNORE' :
                  onDuplicate === 'replace' ? 'REPLACE' : 'INSERT';
  const cols = columns.map(c => `\`${c}\``).join(', ');
  return `${keyword} INTO \`${tableName}\` (${cols}) VALUES `;
}
 
/** 将一批行批量写入 MySQL */
async function flushBatch(conn, prefix, batch) {
  if (batch.length === 0) return;
  const placeholders = batch.map(row => `(${row.map(() => '?').join(', ')})`).join(', ');
  const values = batch.flat();
  await conn.execute(prefix + placeholders, values);
}
 
async function main() {
  const fileSizeMB = (fs.statSync(CONFIG.csvFile).size / 1024 / 1024).toFixed(1);
  console.log(`📂 文件: ${CONFIG.csvFile} (${fileSizeMB} MB)`);
  console.log(`🗄️  目标: ${CONFIG.db.database}.${CONFIG.tableName}`);
  console.log(`📦 批次大小: ${CONFIG.batchSize} 行\n`);
 
  const conn = await mysql.createConnection(CONFIG.db);
  console.log('✅ 数据库连接成功');
 
  // 关闭 autocommit，使用手动事务提升性能
  await conn.execute('SET autocommit = 0');
  await conn.execute('SET unique_checks = 0');
  await conn.execute('SET foreign_key_checks = 0');
 
  let csvColumns  = null;   // CSV 表头
  let dbColumns   = null;   // 对应的 MySQL 列名
  let insertPrefix = null;
  let batch        = [];
  let totalRows    = 0;
  let errorRows    = 0;
  const startTime  = Date.now();
 
  const parser = fs.createReadStream(CONFIG.csvFile).pipe(
    parse({
      columns:          true,   // 第一行作为列名
      skip_empty_lines: true,
      trim:             true,
      relax_column_count: true, // 宽容模式，允许行列数不一致
      bom:              true,   // 自动去除 UTF-8 BOM
    })
  );
 
  for await (const record of parser) {
    // ── 首行：初始化列信息 ───────────────────────────────
    if (csvColumns === null) {
      csvColumns = Object.keys(record);
 
      if (CONFIG.columnMap) {
        // 只取 columnMap 中指定的列
        dbColumns = [];
        for (const [csvCol, dbCol] of Object.entries(CONFIG.columnMap)) {
          if (!csvColumns.includes(csvCol)) {
            console.warn(`⚠️  columnMap 中的列 "${csvCol}" 不存在于 CSV 表头，已跳过`);
          } else {
            dbColumns.push({ csvCol, dbCol });
          }
        }
      } else {
        // 直接使用 CSV 表头
        dbColumns = csvColumns.map(c => ({ csvCol: c, dbCol: c }));
      }
 
      const dbColNames = dbColumns.map(c => c.dbCol);
      insertPrefix = buildInsertPrefix(CONFIG.tableName, dbColNames, CONFIG.onDuplicate);
 
      // ── 自动建表 ─────────────────────────────────────
      if (CONFIG.autoCreateTable) {
        const colDefs = dbColumns.map(({ csvCol, dbCol }) =>
          `\`${dbCol}\` ${guessType(record[csvCol])}`
        ).join(', ');
        const ddl = `CREATE TABLE IF NOT EXISTS \`${CONFIG.tableName}\` (${colDefs})`;
        await conn.execute(ddl);
        console.log(`🔨 自动建表: ${ddl}\n`);
      }
 
      console.log(`📋 CSV 列: ${csvColumns.join(', ')}`);
      console.log(`📋 导入列: ${dbColumns.map(c => c.dbCol).join(', ')}\n`);
      console.log('⏳ 开始导入...');
    }
 
    // ── 组装当前行的值 ───────────────────────────────────
    const rowValues = dbColumns.map(({ csvCol }) => {
      const v = record[csvCol];
      // 空字符串转 NULL（可按需调整）
      return (v === '' || v === undefined) ? null : v;
    });
    batch.push(rowValues);
    totalRows++;
 
    // ── 达到批次大小：写入并提交 ─────────────────────────
    if (batch.length >= CONFIG.batchSize) {
      try {
        await flushBatch(conn, insertPrefix, batch);
        await conn.execute('COMMIT');
      } catch (err) {
        errorRows += batch.length;
        console.error(`\n❌ 批次写入失败 (第 ${totalRows} 行附近): ${err.message}`);
        await conn.execute('ROLLBACK');
      }
      batch = [];
 
      // 进度输出
      if (totalRows % (CONFIG.batchSize * 10) === 0) {
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
        const speed   = (totalRows / elapsed).toFixed(0);
        process.stdout.write(`\r   已处理: ${totalRows.toLocaleString()} 行 | ${elapsed}s | ${speed} 行/s   `);
      }
    }
  }
 
  // ── 写入剩余行 ────────────────────────────────────────────
  if (batch.length > 0) {
    try {
      await flushBatch(conn, insertPrefix, batch);
      await conn.execute('COMMIT');
    } catch (err) {
      errorRows += batch.length;
      console.error(`\n❌ 最终批次写入失败: ${err.message}`);
      await conn.execute('ROLLBACK');
    }
  }
 
  // ── 恢复 MySQL 设置 ───────────────────────────────────────
  await conn.execute('SET unique_checks = 1');
  await conn.execute('SET foreign_key_checks = 1');
  await conn.execute('SET autocommit = 1');
  await conn.end();
 
  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  console.log(`\n\n✅ 导入完成`);
  console.log(`   总行数   : ${totalRows.toLocaleString()}`);
  console.log(`   成功行数 : ${(totalRows - errorRows).toLocaleString()}`);
  console.log(`   失败行数 : ${errorRows}`);
  console.log(`   耗时     : ${elapsed}s`);
  console.log(`   平均速度 : ${(totalRows / elapsed).toFixed(0)} 行/s`);
}
 
main().catch(err => {
  console.error('❌ 致命错误:', err.message);
  process.exit(1);
});
 