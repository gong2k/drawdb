import * as XLSX from "xlsx";
import { nanoid } from "nanoid";

/**
 * drawdb ↔ XLSX 双向转换工具
 *
 * XLSX 结构设计：
 *   - 工作簿中第一张 sheet 固定为 "_index"，列出所有表名/字段数/注释，作为目录
 *   - 其余每张 sheet 对应一张数据表，sheet 名 = 表名（最多 31 字符，已转义非法字符）
 *   - 字段定义一行一列，前 9 列固定：
 *       名称 | 类型 | 大小 | 默认值 | 主键 | 非空 | 唯一 | 自增 | 注释
 *   - 主键 / 非空 / 唯一 / 自增 用 ✓ 标记导出，导入时识别 ✓ / true
 */

const INDEX_SHEET_NAME = "_index";
const MAX_SHEET_NAME_LEN = 31;

const FIELD_COLUMNS = [
  "名称",
  "类型",
  "大小",
  "默认值",
  "主键",
  "非空",
  "唯一",
  "自增",
  "注释",
];

const FIELD_KEYS = [
  "name",
  "type",
  "size",
  "default",
  "primary",
  "notNull",
  "unique",
  "increment",
  "comment",
];

/**
 * 把表名转成 XLSX 合法的 sheet 名
 *  Excel 规则：≤31 字符；不能含 \ / * ? : [ ]
 */
function safeSheetName(name, fallback) {
  const raw = (name || fallback || "Untitled").toString().slice(0, MAX_SHEET_NAME_LEN);
  // eslint-disable-next-line no-useless-escape
  return raw.replace(/[\\/*?:\[\]]/g, "_");
}

/**
 * 在向工作簿加 sheet 时，若 sheet 名重复则在末尾追加 _2, _3 ...
 */
function uniqueSheetName(wb, desired) {
  let name = desired;
  let i = 2;
  while (wb.SheetNames.includes(name)) {
    const suffix = "_" + i;
    name = desired.slice(0, MAX_SHEET_NAME_LEN - suffix.length) + suffix;
    i += 1;
  }
  return name;
}

function checkmark(v) {
  return v ? "✓" : "";
}

function parseCheckmark(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  const s = (v ?? "").toString().trim().toLowerCase();
  return s === "✓" || s === "yes" || s === "y" || s === "1" || s === "true";
}

function fieldToRow(field) {
  return {
    名称: field.name ?? "",
    类型: field.type ?? "",
    大小: field.size ?? "",
    默认值: field.default ?? "",
    主键: checkmark(field.primary),
    非空: checkmark(field.notNull),
    唯一: checkmark(field.unique),
    自增: checkmark(field.increment),
    注释: field.comment ?? "",
  };
}

function rowToField(row, idx) {
  const get = (key) => (row[key] !== undefined && row[key] !== null ? row[key] : "");
  return {
    id: idx,
    name: (get("名称") || "").toString().trim(),
    type: (get("类型") || "").toString().trim(),
    size: get("大小") === "" ? undefined : Number(get("大小")) || undefined,
    default: get("默认值") !== "" ? (get("默认值") || "").toString() : undefined,
    primary: parseCheckmark(get("主键")),
    notNull: parseCheckmark(get("非空")),
    unique: parseCheckmark(get("唯一")),
    increment: parseCheckmark(get("自增")),
    comment: get("注释") !== "" ? (get("注释") || "").toString() : "",
  };
}

/**
 * 将整个 diagram 导出为 .xlsx 文件。浏览器端下载。
 * @param {object} diagram drawdb 的 diagram 对象
 * @param {string} [filename] 缺省自动用标题 + 时间戳
 */
export function exportDiagramToXlsx(diagram, filename) {
  const wb = XLSX.utils.book_new();

  // 1) _index 目录页
  const indexRows = (diagram.tables || []).map((t) => ({
    表名: t.name ?? "",
    字段数: (t.fields || []).length,
    主键字段: (t.fields || []).find((f) => f.primary)?.name ?? "",
    注释: t.comment ?? "",
    颜色: t.color ?? "",
  }));
  if (indexRows.length > 0) {
    const ws = XLSX.utils.json_to_sheet(indexRows);
    XLSX.utils.book_append_sheet(wb, ws, INDEX_SHEET_NAME);
  } else {
    // 即使没有表也给一个空的工作表，免得用户打开报错
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), INDEX_SHEET_NAME);
  }

  // 2) 每张表一张 sheet
  for (const table of diagram.tables || []) {
    const safeName = safeSheetName(table.name, "untitled");
    const sheetName = uniqueSheetName(wb, safeName);

    const rows = (table.fields || []).map(fieldToRow);
    const aoa = [FIELD_COLUMNS, ...rows.map((r) => FIELD_COLUMNS.map((c) => r[c]))];
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // 简单列宽设置（用英文字符宽度估算）
    ws["!cols"] = [
      { wch: 22 }, // 名称
      { wch: 18 }, // 类型
      { wch: 8 },  // 大小
      { wch: 14 }, // 默认值
      { wch: 6 },  // 主键
      { wch: 6 },  // 非空
      { wch: 6 },  // 唯一
      { wch: 6 },  // 自增
      { wch: 28 }, // 注释
    ];

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  const fname =
    filename || `drawdb_${(diagram.title || "diagram").toString().replace(/[^\w-]/g, "_")}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  XLSX.writeFile(wb, fname);
  return fname;
}

/**
 * 解析 xlsx 字节流，返回 Table[] 数组（不含坐标系，调用方负责摆放）
 * @param {ArrayBuffer} arrayBuffer xlsx 文件的字节内容
 * @returns {{ tables: Array, errors: string[] }}
 */
export function importXlsxToTables(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const errors = [];

  if (!wb.SheetNames || wb.SheetNames.length === 0) {
    return { tables: [], errors: ["工作簿为空"] };
  }

  const tables = [];
  let fieldIdCounter = 0;

  for (const sheetName of wb.SheetNames) {
    if (sheetName === INDEX_SHEET_NAME) continue;

    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    const rows = XLSX.utils.sheet_to_json(ws, {
      header: FIELD_COLUMNS,
      defval: "",
      blankrows: false,
      raw: false,
    });

    if (!rows || rows.length === 0) continue;

    // 过滤掉完全空行
    const validRows = rows
      .map((r) => rowToField(r, fieldIdCounter++))
      .filter((f) => f.name);

    if (validRows.length === 0) continue;

    // 主键字段不能没有
    if (!validRows.some((f) => f.primary)) {
      errors.push(`表 "${sheetName}" 没有主键字段（PK 列），已忽略`);
      continue;
    }

    tables.push({
      id: nanoid(),
      name: sheetName,
      comment: "",
      x: 0,
      y: 0,
      color: "",
      fields: validRows,
      indices: [],
    });
  }

  return { tables, errors };
}

/**
 * xlsx 模块版本号，方便排错
 */
export const XLSX_FORMAT_VERSION = "1.0";

export const XLSX_FIELD_COLUMNS = FIELD_COLUMNS;
export const XLSX_FIELD_KEYS = FIELD_KEYS;
export const XLSX_INDEX_SHEET_NAME = INDEX_SHEET_NAME;
