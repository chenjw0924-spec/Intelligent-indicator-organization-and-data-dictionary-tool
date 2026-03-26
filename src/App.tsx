/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import { Upload, FileText, CheckCircle, AlertCircle, Download, Play, RefreshCw, FileSpreadsheet, Info, Layers, Settings, Plus, Trash2 } from 'lucide-react';
import { organizeIndicators, mapToDictionary, ProcessResult } from './lib/processor';
import * as XLSX from 'xlsx';

export default function App() {
  const [activeTab, setActiveTab] = useState<'main' | 'aliases'>('main');
  const [aliases, setAliases] = useState<{source: string, target: string}[]>([
    // 通用别名
    { source: '日期', target: '时间' },
    { source: '比例', target: '占比' },
    { source: '金额', target: '额' },
    { source: '数量', target: '数' },
    { source: '单价', target: '均价' },
    { source: '成本', target: '费用' },
    { source: '利润', target: '收益' },
    // 房地产经营分析常用别名
    { source: '面积', target: '建面' },
    { source: '销售', target: '签约' },
    { source: '认购', target: '认筹' },
    { source: '货值', target: '存货' },
    { source: '去化', target: '销售' },
    { source: '拿地', target: '获取' },
    { source: '交付', target: '交房' },
    { source: '可售', target: '库存' },
    { source: '回款', target: '收款' },
    { source: '楼面价', target: '地价' },
  ]);
  const [pendingFiles, setPendingFiles] = useState<File[]>([]);
  const [dictFile, setDictFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [organizedData, setOrganizedData] = useState<any[] | null>(null);
  const [result, setResult] = useState<ProcessResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [logs, setLogs] = useState<string[]>([]);

  const pendingInputRef = useRef<HTMLInputElement>(null);
  const dictInputRef = useRef<HTMLInputElement>(null);

  const addLog = (msg: string) => {
    setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);
  };

  const handleOrganize = async () => {
    if (pendingFiles.length === 0) return;

    setIsProcessing(true);
    setError(null);
    setResult(null);
    setOrganizedData(null);
    setLogs([]);

    try {
      addLog(`[步骤一] 开始读取并探查 ${pendingFiles.length} 个待处理文件...`);
      addLog('自动解析所有 Sheet 页，并清理格式异常（去除前后空格、处理空值）...');
      
      // Simulate some delay for UI feedback
      await new Promise(resolve => setTimeout(resolve, 500));
      
      const res = await organizeIndicators(pendingFiles);
      
      setOrganizedData(res.data);
      setResult(res);
      addLog(`[步骤一] 整理完成！共提取 ${res.data.length} 条指标数据。您可以直接导出，或继续进行字典映射。`);
    } catch (err: any) {
      setError(err.message || '整理过程中发生错误，请检查文件格式。');
      addLog(`错误: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleMap = async () => {
    if (!organizedData || !dictFile) {
      setError('请先完成指标整理并上传标准数据字典文件。');
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      addLog(`[步骤二] 开始读取标准数据字典文件...`);
      
      // Simulate some delay for UI feedback
      await new Promise(resolve => setTimeout(resolve, 500));
      
      addLog('执行智能 Join 与映射...');
      addLog('以“指标名称”或“字段英文名”为主键进行精确匹配...');
      addLog('启用模糊匹配回退机制（相似度阈值 > 0.85）...');
      addLog(`应用了 ${aliases.filter(a => a.source && a.target).length} 条别名配置...`);
      
      const res = await mapToDictionary(organizedData, dictFile, 0.85, aliases);
      
      addLog('字段补全与状态标记完成。');
      addLog('生成执行报告与输出交付物...');
      
      setResult(res);
      addLog(`[步骤二] 字典映射完成！`);
    } catch (err: any) {
      setError(err.message || '处理过程中发生错误，请检查文件格式。');
      addLog(`错误: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownload = () => {
    if (!result) return;
    
    const worksheet = XLSX.utils.json_to_sheet(result.data);
    const workbook = XLSX.utils.book_new();
    const sheetName = result.step === 1 ? "整理后指标" : "映射结果";
    const fileName = result.step === 1 ? "整理后指标.xlsx" : "指标映射结果.xlsx";
    
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, fileName);
  };

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans">
      <header className="bg-white border-b border-gray-200 px-6 py-4 shadow-sm">
        <div className="max-w-6xl mx-auto flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white">
            <FileSpreadsheet size={24} />
          </div>
          <div>
            <h1 className="text-xl font-bold text-gray-900">自动化指标整理与字典映射工具</h1>
            <p className="text-sm text-gray-500">数据治理、ETL逻辑映射与自动化处理</p>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto p-6 space-y-6">
        <div className="flex border-b border-gray-200 mb-6">
          <button
            className={`py-3 px-6 font-medium text-sm border-b-2 transition-colors ${activeTab === 'main' ? 'border-blue-600 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'}`}
            onClick={() => setActiveTab('main')}
          >
            处理中心
          </button>
          <button
            className={`py-3 px-6 font-medium text-sm border-b-2 transition-colors ${activeTab === 'aliases' ? 'border-blue-600 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'}`}
            onClick={() => setActiveTab('aliases')}
          >
            别名配置
          </button>
        </div>

        {activeTab === 'main' ? (
          <>
            {/* Upload Section */}
            <div className="grid md:grid-cols-2 gap-6">
          {/* Pending File */}
          <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
            <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
              <span className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-sm">1</span>
              待整理指标文件
            </h2>
            <p className="text-sm text-gray-500 mb-4">包含非标准化的初始指标需求（Excel/CSV）</p>
            
            <div 
              className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${pendingFiles.length > 0 ? 'border-blue-400 bg-blue-50' : 'border-gray-300 hover:border-blue-400 hover:bg-gray-50'}`}
              onClick={() => pendingInputRef.current?.click()}
            >
              <input 
                type="file" 
                multiple
                className="hidden" 
                ref={pendingInputRef}
                accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
                onChange={(e) => {
                  if (e.target.files?.length) {
                    setPendingFiles(Array.from(e.target.files));
                    setOrganizedData(null);
                    setResult(null);
                    setLogs([]);
                  }
                }}
              />
              {pendingFiles.length > 0 ? (
                <div className="flex flex-col items-center text-blue-600">
                  <FileText size={32} className="mb-2" />
                  <span className="font-medium">已选择 {pendingFiles.length} 个文件</span>
                  <div className="text-xs text-blue-400 mt-2 max-h-24 overflow-y-auto w-full px-4 space-y-1 text-left">
                    {pendingFiles.map((f, i) => (
                      <div key={i} className="truncate bg-blue-100 px-2 py-1 rounded">
                        {f.name} ({(f.size / 1024).toFixed(1)} KB)
                      </div>
                    ))}
                  </div>
                </div>
              ) : (
                <div className="flex flex-col items-center text-gray-500">
                  <Upload size={32} className="mb-2 text-gray-400" />
                  <span className="font-medium text-gray-700">点击批量上传文件</span>
                  <span className="text-xs mt-1">支持多文件、多Sheet页 (.xlsx, .csv)</span>
                </div>
              )}
            </div>
          </div>

          {/* Dict File */}
          <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
            <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
              <span className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-sm">2</span>
              标准数据字典文件
            </h2>
            <p className="text-sm text-gray-500 mb-4">包含规范定义、业务口径、取数口径等（Excel/CSV）</p>
            
            <div 
              className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${dictFile ? 'border-blue-400 bg-blue-50' : 'border-gray-300 hover:border-blue-400 hover:bg-gray-50'}`}
              onClick={() => dictInputRef.current?.click()}
            >
              <input 
                type="file" 
                className="hidden" 
                ref={dictInputRef}
                accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
                onChange={(e) => {
                  if (e.target.files?.[0]) setDictFile(e.target.files[0]);
                }}
              />
              {dictFile ? (
                <div className="flex flex-col items-center text-blue-600">
                  <FileText size={32} className="mb-2" />
                  <span className="font-medium">{dictFile.name}</span>
                  <span className="text-xs text-blue-400 mt-1">{(dictFile.size / 1024).toFixed(1)} KB</span>
                </div>
              ) : (
                <div className="flex flex-col items-center text-gray-500">
                  <Upload size={32} className="mb-2 text-gray-400" />
                  <span className="font-medium text-gray-700">点击上传文件</span>
                  <span className="text-xs mt-1">支持 .xlsx, .xls, .csv</span>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Action Bar */}
        <div className="flex items-center justify-between bg-white p-4 rounded-xl border border-gray-200 shadow-sm">
          <div className="flex items-center gap-2 text-sm text-gray-600">
            <Info size={18} className="text-blue-500" />
            <span>支持分步操作：您可以仅整理指标并导出，也可以继续上传字典进行映射。</span>
          </div>
          <div className="flex gap-3">
            <button
              onClick={handleOrganize}
              disabled={pendingFiles.length === 0 || isProcessing}
              className={`flex items-center gap-2 px-4 py-2.5 rounded-lg font-medium transition-all ${
                pendingFiles.length === 0 || isProcessing 
                  ? 'bg-gray-100 text-gray-400 cursor-not-allowed' 
                  : 'bg-white border border-blue-600 text-blue-600 hover:bg-blue-50'
              }`}
            >
              {isProcessing && !organizedData ? <RefreshCw size={18} className="animate-spin" /> : <Layers size={18} />}
              1. 仅整理指标
            </button>
            <button
              onClick={handleMap}
              disabled={!organizedData || !dictFile || isProcessing}
              className={`flex items-center gap-2 px-6 py-2.5 rounded-lg font-medium text-white transition-all ${
                !organizedData || !dictFile || isProcessing 
                  ? 'bg-gray-300 cursor-not-allowed' 
                  : 'bg-blue-600 hover:bg-blue-700 shadow-md hover:shadow-lg'
              }`}
            >
              {isProcessing && organizedData ? <RefreshCw size={18} className="animate-spin" /> : <Play size={18} />}
              2. 执行字典映射
            </button>
          </div>
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 p-4 rounded-lg flex items-start gap-3">
            <AlertCircle className="shrink-0 mt-0.5" />
            <div>
              <h3 className="font-medium">处理失败</h3>
              <p className="text-sm mt-1">{error}</p>
            </div>
          </div>
        )}

        {/* Logs & Results */}
        {(logs.length > 0 || result) && (
          <div className="grid lg:grid-cols-3 gap-6">
            {/* Execution Report */}
            <div className="lg:col-span-1 space-y-6">
              <div className="bg-gray-900 text-gray-300 p-4 rounded-xl font-mono text-xs shadow-inner h-64 overflow-y-auto">
                <div className="text-green-400 mb-2"># 运行日志</div>
                {logs.map((log, i) => (
                  <div key={i} className="mb-1">{log}</div>
                ))}
                {isProcessing && <div className="animate-pulse mt-2">_</div>}
              </div>

              {result && result.step === 2 && result.report && (
                <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                  <h3 className="text-lg font-semibold mb-4 border-b pb-2">执行报告</h3>
                  <div className="space-y-3">
                    <div className="flex justify-between items-center">
                      <span className="text-gray-600">输入总行数</span>
                      <span className="font-bold text-gray-900">{result.report.total}</span>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="text-gray-600 flex items-center gap-1"><CheckCircle size={14} className="text-green-500"/> 精确匹配</span>
                      <span className="font-bold text-green-600">{result.report.exactMatch}</span>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="text-gray-600 flex items-center gap-1"><AlertCircle size={14} className="text-yellow-500"/> 模糊匹配</span>
                      <span className="font-bold text-yellow-600">{result.report.fuzzyMatch}</span>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="text-gray-600 flex items-center gap-1"><AlertCircle size={14} className="text-red-500"/> 未匹配</span>
                      <span className="font-bold text-red-600">{result.report.unmatched}</span>
                    </div>
                  </div>

                  {result.report.unmatched > 0 && (
                    <div className="mt-4 pt-4 border-t border-gray-100">
                      <span className="text-sm font-medium text-gray-700 mb-2 block">未匹配指标列表 (前5项):</span>
                      <ul className="text-xs text-red-500 space-y-1">
                        {result.report.unmatchedList.slice(0, 5).map((item, i) => (
                          <li key={i} className="truncate">• {item}</li>
                        ))}
                        {result.report.unmatchedList.length > 5 && (
                          <li className="text-gray-400 italic">... 及其他 {result.report.unmatchedList.length - 5} 项</li>
                        )}
                      </ul>
                    </div>
                  )}
                </div>
              )}

              {result && result.step === 1 && (
                <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                  <h3 className="text-lg font-semibold mb-4 border-b pb-2">整理摘要</h3>
                  <div className="space-y-3">
                    <div className="flex justify-between items-center">
                      <span className="text-gray-600">提取总行数</span>
                      <span className="font-bold text-blue-600">{result.data.length}</span>
                    </div>
                    <p className="text-sm text-gray-500 mt-4">
                      指标已成功合并并添加了文件与Sheet来源标记。您可以直接导出，或在右侧上传字典文件进行映射。
                    </p>
                  </div>
                </div>
              )}
            </div>

            {/* Data Preview */}
            {result && (
              <div className="lg:col-span-2 bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden flex flex-col">
                <div className="p-4 border-b border-gray-200 flex justify-between items-center bg-gray-50">
                  <h3 className="font-semibold text-gray-800">数据预览 (前 10 行)</h3>
                  <button
                    onClick={handleDownload}
                    className="flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 text-white text-sm font-medium rounded-lg transition-colors"
                  >
                    <Download size={16} /> 导出结果
                  </button>
                </div>
                <div className="overflow-x-auto p-0 flex-1">
                  <table className="w-full text-sm text-left">
                    <thead className="text-xs text-gray-700 uppercase bg-gray-100">
                      <tr>
                        {Object.keys(result.data[0] || {}).map(key => (
                          <th key={key} className="px-4 py-3 whitespace-nowrap">{key}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {result.data.slice(0, 10).map((row, i) => (
                        <tr key={i} className="border-b border-gray-100 hover:bg-gray-50">
                          {Object.keys(result.data[0] || {}).map(key => {
                            const val = row[key];
                            const isStatus = key === '映射核令人状态' || key === '映射核对状态';
                            let statusClass = '';
                            if (isStatus) {
                              if (val === '匹配成功') statusClass = 'text-green-600 font-medium bg-green-50 px-2 py-1 rounded';
                              else if (val === '模糊匹配-需复核') statusClass = 'text-yellow-600 font-medium bg-yellow-50 px-2 py-1 rounded';
                              else statusClass = 'text-red-600 font-medium bg-red-50 px-2 py-1 rounded';
                            }
                            
                            return (
                              <td key={key} className="px-4 py-3 whitespace-nowrap max-w-[200px] truncate">
                                {isStatus ? (
                                  <span className={statusClass}>{val}</span>
                                ) : (
                                  <span title={String(val)}>{String(val)}</span>
                                )}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {result.data.length === 0 && (
                    <div className="p-8 text-center text-gray-500">
                      没有数据可预览
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}
          </>
        ) : (
          <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
            <h2 className="text-lg font-semibold mb-2 flex items-center gap-2">
              <Settings size={20} className="text-blue-600" />
              指标别名配置
            </h2>
            <p className="text-sm text-gray-500 mb-6">配置常见的指标同义词，在进行字典映射时，系统会自动尝试替换这些词汇以提高匹配率（双向生效）。</p>
            
            <div className="space-y-3 max-w-xl">
              {aliases.map((alias, index) => (
                <div key={index} className="flex items-center gap-3 bg-gray-50 p-3 rounded-lg border border-gray-100">
                  <input 
                    type="text" 
                    value={alias.source}
                    onChange={(e) => {
                      const newAliases = [...aliases];
                      newAliases[index].source = e.target.value;
                      setAliases(newAliases);
                    }}
                    className="border border-gray-300 rounded-md px-3 py-2 text-sm flex-1 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
                    placeholder="词汇 A (如: 日期)"
                  />
                  <span className="text-gray-400 font-medium">↔</span>
                  <input 
                    type="text" 
                    value={alias.target}
                    onChange={(e) => {
                      const newAliases = [...aliases];
                      newAliases[index].target = e.target.value;
                      setAliases(newAliases);
                    }}
                    className="border border-gray-300 rounded-md px-3 py-2 text-sm flex-1 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
                    placeholder="词汇 B (如: 时间)"
                  />
                  <button 
                    onClick={() => {
                      const newAliases = aliases.filter((_, i) => i !== index);
                      setAliases(newAliases);
                    }}
                    className="p-2 text-red-500 hover:bg-red-50 rounded-md transition-colors"
                    title="删除别名"
                  >
                    <Trash2 size={18} />
                  </button>
                </div>
              ))}
            </div>
            
            <button 
              onClick={() => setAliases([...aliases, { source: '', target: '' }])}
              className="mt-4 flex items-center gap-2 text-sm text-blue-600 hover:text-blue-700 font-medium px-4 py-2 rounded-lg hover:bg-blue-50 transition-colors"
            >
              <Plus size={18} /> 添加别名
            </button>
          </div>
        )}
      </main>
    </div>
  );
}
