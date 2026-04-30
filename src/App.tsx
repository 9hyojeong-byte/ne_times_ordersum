/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Download, Trash2, Plus, AlertCircle, FileSpreadsheet, X } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { OrderData, OrderKey, OUTPUT_COLUMNS } from './types';
import { processExcelFile, processPastedText } from './utils/excelProcessor';

export default function App() {
  const [orders, setOrders] = useState<OrderData[]>([]);
  const [notices, setNotices] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [pastedValue, setPastedValue] = useState('');

  const handlePasteData = async (text: string) => {
    if (!text.trim()) return;
    setIsProcessing(true);
    try {
      const result = await processPastedText(text);
      if (result.errors.length > 0) {
        setNotices(prev => [...prev, ...result.errors]);
      }
      if (result.data.length > 0) {
        setOrders(prev => [...prev, ...result.data]);
        setPastedValue('');
      } else if (result.errors.length === 0) {
        setNotices(prev => [...prev, "데이터를 파싱하지 못했습니다. 형식을 확인해 주세요."]);
      }
    } catch (err) {
      setNotices(prev => [...prev, "데이터 처리 중 오류가 발생했습니다."]);
    } finally {
      setIsProcessing(false);
    }
  };

  const updateCell = (id: string, key: OrderKey, value: string) => {
    setOrders(prev => prev.map(order => 
      order.id === id ? { ...order, [key]: value } : order
    ));
  };

  const [showConfirm, setShowConfirm] = useState(false);

  const clearAll = () => {
    if (showConfirm) {
      setOrders([]);
      setNotices([]);
      setPastedValue('');
      setShowConfirm(false);
    } else {
      setShowConfirm(true);
      setTimeout(() => setShowConfirm(false), 3000); // 3초 후 초기화
    }
  };

  const deleteRow = (id: string) => {
    setOrders(prev => prev.filter(order => order.id !== id));
  };

  const downloadExcel = () => {
    const dataToExport = orders.map(({ id, ...rest }) => rest);
    const worksheet = XLSX.utils.json_to_sheet(dataToExport, { header: OUTPUT_COLUMNS });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    // Generate filename: 주문건합본_소스1_소스2_YYMMDD.xlsx
    const sources = Array.from(new Set(orders.map(o => o["이름(주문)"]))).filter(Boolean);
    const sourceStr = sources.length > 0 ? `_${sources.join('_')}` : '';

    const kst = new Date().toLocaleString("en-US", {timeZone: "Asia/Seoul"});
    const d = new Date(kst);
    const yy = d.getFullYear().toString().slice(-2);
    const mm = (d.getMonth() + 1).toString().padStart(2, '0');
    const dd = d.getDate().toString().padStart(2, '0');
    const dateStr = `${yy}${mm}${dd}`;

    XLSX.writeFile(workbook, `주문건합본${sourceStr}_${dateStr}.xlsx`);
  };

  const removeNotice = (index: number) => {
    setNotices(prev => prev.filter((_, i) => i !== index));
  };

  const sourceFilesCount = Array.from(new Set(orders.map(o => o["이름(주문)"]))).length;

  return (
    <div className="min-h-screen bg-[#f9fafb] text-[#111827] font-sans p-6">
      <div className="max-w-7xl mx-auto space-y-6 flex flex-col h-[calc(100vh-3rem)]">
        {/* Header */}
        <header className="flex justify-between items-center pb-4 border-b border-gray-200 relative z-50">
          <div className="flex items-center space-x-3">
            <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center text-white font-bold">EX</div>
            <h1 className="text-xl font-bold tracking-tight">
              NE times 외부주문내역 merger <span className="text-gray-400 font-normal ml-2 text-sm italic">v1.2.0 - 구효정</span>
            </h1>
          </div>
          <div className="flex items-center gap-2">
            {orders.length > 0 && (
              <button 
                onClick={clearAll}
                className={`px-4 py-2 text-sm border rounded transition-all font-medium relative ${
                  showConfirm 
                    ? 'bg-red-600 text-white border-red-600 animate-pulse' 
                    : 'bg-red-50/50 text-red-600 border-red-200 hover:bg-red-50'
                }`}
              >
                {showConfirm ? '정말 삭제할까요? (다시 클릭)' : '전체 삭제 (Clear All)'}
              </button>
            )}
            <button
              onClick={downloadExcel}
              disabled={orders.length === 0}
              className={`px-4 py-2 text-sm rounded font-medium transition-all shadow-sm flex items-center gap-2 ${
                orders.length === 0 
                  ? 'bg-slate-200 text-slate-400 cursor-not-allowed' 
                  : 'bg-blue-600 text-white hover:bg-blue-700 active:scale-95'
              }`}
            >
              <Download size={16} />
              엑셀로 다운받기 (.xlsx)
            </button>
          </div>
        </header>

        {/* Paste Section */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <section
            className={`col-span-1 md:col-span-2 border border-gray-200 bg-white rounded-xl flex flex-col p-4 shadow-sm transition-all duration-200 ${
              isProcessing ? 'opacity-50 pointer-events-none cursor-wait' : ''
            }`}
          >
            <div className="flex justify-between items-center mb-2">
              <h3 className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">엑셀 데이터를 붙여넣기 해주세요. - yes24(입금주문), ssg, 나이스북, 북콤파스, 북매거진 형식 지원 </h3>
              <button 
                onClick={() => handlePasteData(pastedValue)}
                disabled={!pastedValue.trim() || isProcessing}
                className="px-3 py-1 bg-blue-600 text-white rounded text-[11px] font-bold uppercase hover:bg-blue-700 disabled:bg-gray-200 disabled:text-gray-400"
              >
                {isProcessing ? 'Processing...' : 'Apply Data'}
              </button>
            </div>
            <textarea
              className="flex-grow w-full h-32 md:h-full p-3 text-xs font-mono bg-gray-50 border border-gray-200 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:bg-white transition-all placeholder:text-gray-300 resize-none"
              placeholder="엑셀에서 헤더(품목, 주문상품명, 정간물명 등)를 포함하여 데이터를 복사한 후 여기에 붙여넣으세요.&#10;예:&#10;품목	수령자	수령자휴대전화	...&#10;상품A	홍길동	010-1234-5678	..."
              value={pastedValue}
              onChange={(e) => setPastedValue(e.target.value)}
              onPaste={(e) => {
                // Optional: Auto-process on paste? 
                // Let's stick to manual button to prevent accidental clicks
              }}
            />
            <p className="mt-2 text-[10px] text-gray-400 italic">
              헤더를 포함한 탭(Tab) 구분 형식의 텍스트를 지원합니다.
            </p>
          </section>

          <aside className="bg-white border border-gray-200 rounded-xl p-5 flex flex-col justify-between shadow-sm relative overflow-hidden">
            <h3 className="text-[10px] font-bold text-gray-400 uppercase tracking-wider mb-4 border-b border-gray-50 pb-2">Processing Summary</h3>
            <div className="space-y-4">
              <div className="flex justify-between items-end">
                <span className="text-sm text-gray-500">인식된 주문 건 수</span>
                <span className="text-2xl font-mono leading-none font-medium">{orders.length.toLocaleString()}</span>
              </div>
              <div className="flex justify-between items-end">
                <span className="text-sm text-gray-500">인식된 파일 수</span>
                <span className="text-2xl font-mono leading-none font-medium">{sourceFilesCount.toString().padStart(2, '0')}</span>
              </div>
            </div>
            <div className="mt-6 flex gap-2">
              <span className="px-2 py-1 bg-green-100 text-green-700 text-[10px] font-bold rounded uppercase">Validation Pass</span>
              <span className="px-2 py-1 bg-blue-100 text-blue-700 text-[10px] font-bold rounded uppercase">Auto-Mapped</span>
            </div>
          </aside>
        </div>

        {/* Alerts Section (Floating or Top fixed for visibility) */}
        {notices.length > 0 && (
          <div className="space-y-1">
            {notices.map((notice, idx) => (
              <div key={idx} className="flex items-center justify-between px-3 py-2 bg-orange-50 border border-orange-100 rounded text-orange-700 text-[12px]">
                <div className="flex items-center gap-2">
                  <AlertCircle size={14} />
                  <span>{notice}</span>
                </div>
                <button onClick={() => removeNotice(idx)} className="text-orange-400 hover:text-orange-600">
                  <X size={14} />
                </button>
              </div>
            ))}
          </div>
        )}

        {/* Table Section */}
        <section className="flex-grow bg-white border border-gray-200 rounded-xl shadow-sm overflow-hidden flex flex-col min-h-0">
          <div className="overflow-x-auto overflow-y-auto custom-scrollbar flex-grow">
            <table className="w-full text-[13px] text-left border-collapse min-w-[2800px]">
              <thead className="sticky top-0 z-20">
                <tr className="bg-[#f3f4f6] border-b border-gray-200">
                  <th className="px-3 py-3 font-semibold text-[11px] uppercase tracking-wider text-gray-500 w-12 text-center border-r border-gray-200">#</th>
                  {OUTPUT_COLUMNS.map(col => (
                    <th key={col} className="px-3 py-3 font-semibold text-[11px] uppercase tracking-wider text-gray-500 min-w-[140px] border-r border-gray-200">
                      {col}
                    </th>
                  ))}
                  <th className="px-3 py-3 font-semibold text-[11px] uppercase tracking-wider text-gray-500 w-12 text-center sticky right-0 bg-[#f3f4f6] z-30 shadow-[-1px_0_0_#e5e7eb]">
                    -
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {orders.length === 0 ? (
                  <tr>
                    <td colSpan={OUTPUT_COLUMNS.length + 2} className="py-24 text-center">
                      <div className="flex flex-col items-center gap-3 grayscale opacity-30">
                        <FileSpreadsheet size={40} />
                        <p className="text-sm font-medium">No order data found.</p>
                      </div>
                    </td>
                  </tr>
                ) : (
                  orders.map((order, index) => (
                    <tr 
                      key={order.id} 
                      className={`hover:bg-gray-50 transition-colors group ${order["상담내용 입력"] ? 'bg-yellow-50/70' : ''}`}
                    >
                      <td className="px-3 py-2 text-center text-gray-400 font-mono text-[11px] border-r border-gray-100">{index + 1}</td>
                      {OUTPUT_COLUMNS.map(col => {
                        const isDate = col.includes('일');
                        const isItem = col === '상품번호';
                        const isOrderSite = col === '이름(주문)';
                        
                        let colorClass = "text-gray-900";
                        if (isOrderSite) {
                          if (order[col] === '북콤파스') colorClass = "text-blue-600 font-medium";
                          if (order[col] === '더매거진') colorClass = "text-orange-600 font-medium";
                          if (order[col] === '나이스북') colorClass = "text-purple-600 font-medium";
                          if (order[col] === 'yes24') colorClass = "text-emerald-600 font-medium";
                          if (order[col] === 'ssg') colorClass = "text-rose-600 font-medium";
                        }

                        return (
                          <td key={col} className="px-0 py-0 border-r border-gray-100 relative group/cell">
                            <input
                              type="text"
                              value={order[col]}
                              onChange={(e) => updateCell(order.id, col, e.target.value)}
                              className={`w-full px-3 py-3 focus:outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 transition-all bg-transparent editable-cell border-none ${colorClass} ${ (isDate || isItem) ? 'font-mono text-[12px]' : ''}`}
                            />
                          </td>
                        );
                      })}
                      <td className="px-3 py-2 text-center sticky right-0 bg-white group-hover:bg-gray-50 transition-colors shadow-[-1px_0_0_#e5e7eb] z-10">
                        <button
                          onClick={() => deleteRow(order.id)}
                          className="text-gray-300 hover:text-red-500 transition-colors text-lg font-light leading-none"
                        >
                          ×
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>

        {/* Footer */}
        <footer className="flex justify-between items-center text-[10px] text-gray-400 font-medium uppercase tracking-wider">
          <div className="flex space-x-4">
            <span>System: Online</span>
            <span>Last Merge: {new Date().toLocaleTimeString()}</span>
          </div>
          <div>All data is processed locally in the browser</div>
        </footer>
      </div>
    </div>
  );
}

