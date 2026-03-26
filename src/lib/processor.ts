import * as XLSX from 'xlsx';
import { similarity } from './levenshtein';
import { GoogleGenAI, Type } from "@google/genai";

export interface ProcessResult {
  data: any[];
  report?: {
    total: number;
    exactMatch: number;
    fuzzyMatch: number;
    unmatched: number;
    unmatchedList: string[];
  };
  step: 1 | 2;
}

interface Block {
  cat1: string;
  cat2: string;
  rows: any[][];
}

export async function organizeIndicators(pendingFiles: File[]): Promise<ProcessResult> {
  let rawPendingData: any[] = [];
  
  // Initialize Gemini API
  const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
  
  for (const file of pendingFiles) {
    const sheets = await readAllSheets(file, true); 
    
    for (const sheet of sheets) {
      const rows = sheet.data as any[][];
      if (rows.length < 2) continue; // Skip empty or virtually empty sheets
      
      console.log(`[Processor] Scanning Sheet: ${sheet.sheetName} (${rows.length} rows)`);
      
      // Clean empty columns and rows to reduce noise and token usage
      const cleanedRows = cleanEmptyColumnsAndRows(rows);
      if (cleanedRows.length === 0) continue;

      // Chunk the sheet into 200-row chunks to avoid exceeding output token limits
      const CHUNK_SIZE = 200;
      for (let i = 0; i < cleanedRows.length; i += CHUNK_SIZE) {
        const chunk = cleanedRows.slice(i, i + CHUNK_SIZE);
        const csvContent = chunk.map(row => row.join(',')).join('\n');
        
        console.log(`[Processor] Sending chunk to AI -> Sheet: ${sheet.sheetName} | Rows: ${i} to ${i + chunk.length}`);
        
        const prompt = `你是一个专业的数据分析师。我将提供一个Excel看板的CSV格式数据（部分或全部）。

请从以下CSV数据中提取出所有的业务指标名称，并推断其数据类型，以及它所属的分类。
注意处理以下复杂排版：
1. 键值对平铺：左边是指标名，右侧或下方是空白或数值。提取文本作为指标名。
2. 隐式语义：如果遇到类似"累计认购"下方跟着"11亿"和"xx万m2"，请结合单位自动推断并补全指标名称为"累计认购金额"和"累计认购面积"。
3. 二维交叉表：如果存在行表头和列表头，请将它们组合成完整的指标名称（例如行头"已开工" + 列头"面积" = "已开工面积"）。
4. 并排表格：数据可能存在左右并排的多个表格（例如左边是"二、库存分析"，右边是"三、债务分析"），请务必横向扫描每一行，不要遗漏右侧的表格内容。

严格要求：
- 仅提取存在的指标，绝不能凭空捏造（幻觉）不存在的指标！
- 排除纯数据值（如"11亿", "10%", "2023-01-01"），只提取"指标名称"。
- 提取每个指标所属的"大类"（如"一、销售分析"、"二、库存分析"）和"小类"（如果有，如"1.1 签约"）。如果没有明确分类，可以填"未分类"。
- 返回JSON数组，每个对象包含 "大类" (string), "小类" (string), "指标名称" (string) 和 "数据类型" (string，仅限"数值"、"日期"、"文本")。

CSV数据：
${csvContent}
`;

        try {
          const response = await ai.models.generateContent({
            model: "gemini-3.1-pro-preview",
            contents: prompt,
            config: {
              responseMimeType: "application/json",
              responseSchema: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    大类: { type: Type.STRING },
                    小类: { type: Type.STRING },
                    指标名称: { type: Type.STRING },
                    数据类型: { type: Type.STRING }
                  },
                  required: ["大类", "小类", "指标名称", "数据类型"]
                }
              }
            }
          });

          let jsonStr = response.text;
          if (jsonStr) {
            jsonStr = jsonStr.replace(/```json/gi, '').replace(/```/g, '').trim();
            try {
              const parsed = JSON.parse(jsonStr);
              for (const item of parsed) {
                if (!item['指标名称']) continue;
                rawPendingData.push({
                  '所属文件': file.name,
                  '指标类别': sheet.sheetName,
                  '分类1': item['大类'] || '未分类',
                  '分类2': item['小类'] || '',
                  '指标名称': item['指标名称'],
                  '数据类型': item['数据类型'] || '文本'
                });
              }
            } catch (parseError) {
              console.error(`[Processor] JSON Parse Error in chunk:`, parseError, jsonStr);
            }
          }
        } catch (error) {
          console.error(`[Processor] AI Request Failed:`, error);
        }
      }
    }
  }

  return { data: rawPendingData, step: 1 };
}

function cleanEmptyColumnsAndRows(rows: any[][]): any[][] {
  if (rows.length === 0) return [];
  
  // First, filter out completely empty rows
  const nonEmptyRows = rows.filter(row => {
    return row.some(val => val !== undefined && val !== null && String(val).trim() !== '');
  });

  if (nonEmptyRows.length === 0) return [];

  const maxCols = Math.max(...nonEmptyRows.map(r => r.length));
  const colHasData = new Array(maxCols).fill(false);

  for (let r = 0; r < nonEmptyRows.length; r++) {
    for (let c = 0; c < nonEmptyRows[r].length; c++) {
      const val = nonEmptyRows[r][c];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        colHasData[c] = true;
      }
    }
  }

  return nonEmptyRows.map(row => {
    return row.filter((_, c) => colHasData[c]);
  });
}

export interface Alias {
  source: string;
  target: string;
}

export async function mapToDictionary(
  organizedData: any[],
  dictFile: File,
  similarityThreshold: number = 0.85,
  aliases: Alias[] = []
): Promise<ProcessResult> {
  // Read dict file (all sheets)
  const dictSheets = await readAllSheets(dictFile, false);
  const rawDictData = dictSheets.flatMap(s => s.data);
  const cleanedDict = cleanData(rawDictData);
  
  const resultData = [];
  const report = {
    total: organizedData.length,
    exactMatch: 0,
    fuzzyMatch: 0,
    unmatched: 0,
    unmatchedList: [] as string[]
  };

  for (const row of organizedData) {
    // Find key in pending row
    const pendingKey = getPendingKey(row);
    const sourceInfo = `[${row['所属文件']} - ${row['指标类别']}]`;
    
    if (!pendingKey) {
      resultData.push({
        ...row,
        '业务口径': '',
        '取数口径': '',
        '数据来源表': '',
        '系统查询位置': '',
        '对应查询报表': '',
        '映射核对状态': '未匹配-需人工接入'
      });
      report.unmatched++;
      report.unmatchedList.push(`${sourceInfo} 未知指标 (无名称)`);
      continue;
    }

    const pendingKeys = generateAliasKeys(pendingKey, aliases);

    // Find best match across all sheets
    const matchResult = findBestMatch(pendingKeys, cleanedDict, similarityThreshold);
    
    if (matchResult.match) {
      const match = matchResult.match;
      const status = matchResult.isExact ? '匹配成功' : '模糊匹配-需复核';
      
      resultData.push({
        ...row,
        '业务口径': getFieldValue(match, ['业务口径', '业务定义', '指标定义']),
        '取数口径': getFieldValue(match, ['取数口径', '计算逻辑']),
        '数据来源表': getFieldValue(match, ['数据来源表', '底层表', '来源表']),
        '系统查询位置': getFieldValue(match, ['系统查询位置', '查询位置']),
        '对应查询报表': getFieldValue(match, ['对应查询报表', '查询报表']),
        '映射核对状态': status
      });
      
      if (matchResult.isExact) {
        report.exactMatch++;
      } else {
        report.fuzzyMatch++;
      }
    } else {
      resultData.push({
        ...row,
        '业务口径': '',
        '取数口径': '',
        '数据来源表': '',
        '系统查询位置': '',
        '对应查询报表': '',
        '映射核对状态': '未匹配-需人工接入'
      });
      report.unmatched++;
      report.unmatchedList.push(`${sourceInfo} ${pendingKey}`);
    }
  }

  return { data: resultData, report, step: 2 };
}

function getFieldValue(row: any, possibleNames: string[]): string {
  if (!row) return '';
  for (const key of Object.keys(row)) {
    for (const name of possibleNames) {
      if (key.includes(name) && row[key] !== undefined && row[key] !== null) {
        return String(row[key]).trim();
      }
    }
  }
  return '';
}

function generateAliasKeys(key: string, aliases: Alias[]): string[] {
  const keys = new Set<string>([key]);
  for (const alias of aliases) {
    if (!alias.source || !alias.target) continue;
    if (key.includes(alias.source)) {
      keys.add(key.replace(alias.source, alias.target));
    }
    if (key.includes(alias.target)) {
      keys.add(key.replace(alias.target, alias.source));
    }
  }
  return Array.from(keys);
}

function findBestMatch(pendingKeys: string[], dict: any[], threshold: number): { match: any | null, isExact: boolean } {
  let bestMatch = null;
  let highestScore = -1;
  let isBestMatchExact = false;
  
  for (const row of dict) {
    const keys = Object.keys(row);
    const nameKey = keys.find(k => k.includes('指标名称') || k.includes('指标名') || k === '名称');
    const engKey = keys.find(k => k.includes('字段英文名') || k.includes('英文名') || k.includes('字段名'));
    
    let isExact = false;
    let sim = 0;
    
    for (const pk of pendingKeys) {
      const pkLower = pk.toLowerCase();
      if (nameKey && String(row[nameKey]).toLowerCase() === pkLower) isExact = true;
      if (engKey && String(row[engKey]).toLowerCase() === pkLower) isExact = true;
      
      if (!isExact) {
        const sim1 = nameKey ? similarity(pk, String(row[nameKey])) : 0;
        const sim2 = engKey ? similarity(pk, String(row[engKey])) : 0;
        sim = Math.max(sim, sim1, sim2);
      }
      
      if (isExact) break; // No need to check other aliases if we found an exact match
    }
    
    if (isExact || sim > threshold) {
      // Calculate data richness score to prefer rows with actual definitions
      let dataCount = 0;
      if (getFieldValue(row, ['业务口径', '业务定义', '指标定义'])) dataCount++;
      if (getFieldValue(row, ['取数口径', '计算逻辑'])) dataCount++;
      if (getFieldValue(row, ['数据来源表', '底层表', '来源表'])) dataCount++;
      if (getFieldValue(row, ['系统查询位置', '查询位置'])) dataCount++;
      if (getFieldValue(row, ['对应查询报表', '查询报表'])) dataCount++;
      
      // Exact matches always score higher than fuzzy matches
      const currentScore = isExact ? 1000 + dataCount : sim * 100 + dataCount;
      
      if (currentScore > highestScore) {
        highestScore = currentScore;
        bestMatch = row;
        isBestMatchExact = isExact;
      }
    }
  }
  
  return { match: bestMatch, isExact: isBestMatchExact };
}

function readAllSheets(file: File, as2DArray: boolean = false): Promise<{ sheetName: string, data: any[] }[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheets = workbook.SheetNames.map(name => ({
          sheetName: name,
          data: XLSX.utils.sheet_to_json(workbook.Sheets[name], { 
            header: as2DArray ? 1 : undefined,
            defval: as2DArray ? undefined : '' 
          })
        }));
        resolve(sheets);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function cleanData(data: any[]): any[] {
  return data.map(row => {
    const cleanedRow: any = {};
    for (const key in row) {
      const cleanedKey = key.trim();
      const value = row[key];
      cleanedRow[cleanedKey] = typeof value === 'string' ? value.trim() : value;
    }
    return cleanedRow;
  });
}

function getPendingKey(row: any): string | null {
  const possibleKeys = ['指标名称', '指标名', '名称', '字段英文名', '英文名', '字段名'];
  for (const key of possibleKeys) {
    if (row[key]) return String(row[key]);
  }
  // Fallback to first column if it looks like a name (ignoring our injected metadata)
  const keys = Object.keys(row).filter(k => k !== '所属文件' && k !== '所属Sheet');
  if (keys.length > 0 && row[keys[0]]) {
    return String(row[keys[0]]);
  }
  return null;
}
