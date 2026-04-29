/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as XLSX from 'xlsx';
import { OrderData, ProcessingResult } from '../types';

/**
 * Helper to get the last day of a month
 */
function getLastDayOfMonth(year: number, month: number): string {
  const date = new Date(year, month, 0); // month is 1-indexed for the '0' trick
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * Normalizes "YYYY MM" to "YYYY-MM-DD"
 */
function formatYearMonth(val: any, type: 'start' | 'end'): string {
  if (!val) return '';
  const str = String(val).trim();
  const match = str.match(/^(\d{4})[ \.\-]?(\d{1,2})/);
  
  if (match) {
    const year = parseInt(match[1]);
    const month = parseInt(match[2]);
    if (type === 'start') {
      return `${year}-${String(month).padStart(2, '0')}-01`;
    } else {
      return getLastDayOfMonth(year, month);
    }
  }
  return '';
}

/**
 * Generic string cleaner
 */
function clean(val: any): string {
  return val === undefined || val === null ? '' : String(val).trim();
}

/**
 * Processes a single sheet and converts it to our format
 */
export async function processExcelFile(file: File): Promise<ProcessingResult> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    const result: ProcessingResult = { data: [], errors: [] };

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        let foundValidSheet = false;
        workbook.SheetNames.forEach((sheetName) => {
          const trimmedName = sheetName.trim();
          const worksheet = workbook.Sheets[sheetName];
          const rawData: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

          if (trimmedName === '북콤파스') {
            processBookCompass(rawData, result);
            foundValidSheet = true;
          } else if (trimmedName === '더매거진') {
            processTheMagazine(rawData, result);
            foundValidSheet = true;
          } else if (trimmedName === '나이스북') {
            processNiceBook(rawData, result);
            foundValidSheet = true;
          }
        });

        if (!foundValidSheet) {
          result.errors.push(`'${file.name}' 파일에서 지원하는 시트(북콤파스, 더매거진, 나이스북)를 찾을 수 없습니다.`);
        }
        resolve(result);
      } catch (err) {
        result.errors.push(`파일을 읽는 중 오류가 발생했습니다: ${file.name}`);
        resolve(result);
      }
    };

    reader.readAsArrayBuffer(file);
  });
}

function createEmptyRow(): OrderData {
  return {
    id: crypto.randomUUID(),
    "정산구분": "능률",
    "상품번호": "",
    "시작일": "",
    "종료일": "",
    "이름(주문)": "",
    "메일(주문)": "",
    "전번(주문)": "",
    "휴대폰(주문)": "",
    "우편번호(주문)": "",
    "주소1(주문)": "",
    "상세주소(주문)": "",
    "이름(배송)": "",
    "메일(배송)": "",
    "전번(배송)": "",
    "휴대폰(배송)": "",
    "우편번호(배송)": "",
    "주소1(배송)": "",
    "상세주소(배송)": "",
    "배송방법": "우편",
    "계산서": "미발행(개인)",
    "상담내용 입력": ""
  };
}

/**
 * Helper to find a value by flexible key matching
 */
function getValue(row: any, ...possibleKeys: string[]): any {
  const rowKeys = Object.keys(row);
  for (const k of possibleKeys) {
    const target = k.toLowerCase().replace(/\s/g, '');
    const foundKey = rowKeys.find(rk => rk.toLowerCase().replace(/\s/g, '') === target);
    if (foundKey) return row[foundKey];
  }
  return '';
}

function processBookCompass(data: any[], result: ProcessingResult) {
  // Check if it's actually BookCompass by looking for '품목'
  const firstRow = data[0] || {};
  if (getValue(firstRow, '품목') === '') {
    result.errors.push("북콤파스: '품목' 컬럼을 찾을 수 없습니다.");
    return;
  }

  data.forEach((row) => {
    const item = createEmptyRow();
    item["이름(주문)"] = "북콤파스";
    item["상품번호"] = clean(getValue(row, '품목'));
    item["시작일"] = formatYearMonth(getValue(row, '시작호수'), 'start');
    item["종료일"] = formatYearMonth(getValue(row, '만기호수'), 'end');
    
    // Delivery mapping
    const name = clean(getValue(row, '수령자', '수령자명')) || clean(getValue(row, '주문자', '주문자명'));
    item["이름(배송)"] = name;
    item["휴대폰(배송)"] = clean(getValue(row, '수령자휴대전화', '수령자휴대폰', '연락처', '전화번호', '휴대전화'));
    item["우편번호(배송)"] = clean(getValue(row, '우편번호', '수령자우편번호'));
    item["주소1(배송)"] = clean(getValue(row, '주소', '수령자주소'));
    item["상세주소(배송)"] = clean(getValue(row, '상세주소', '수령자상세주소'));

    result.data.push(item);
  });
}

function processTheMagazine(data: any[], result: ProcessingResult) {
  const firstRow = data[0] || {};
  if (getValue(firstRow, '주문상품명') === '') {
    result.errors.push("더매거진: '주문상품명' 컬럼을 찾을 수 없습니다.");
    return;
  }

  data.forEach((row) => {
    const item = createEmptyRow();
    item["이름(주문)"] = "더매거진";
    
    // Product Number
    const rawProd = clean(getValue(row, '주문상품명'));
    item["상품번호"] = rawProd.split('(')[0].trim();

    // Dates from 상품옵션
    const options = clean(getValue(row, '상품옵션'));
    const dateMatch = options.match(/(\d{4}[ \-]\d{1,2})\s*~\s*(\d{4}[ \-]\d{1,2})/);
    if (dateMatch) {
      item["시작일"] = formatYearMonth(dateMatch[1], 'start');
      item["종료일"] = formatYearMonth(dateMatch[2], 'end');
    }

    // Delivery
    item["이름(배송)"] = clean(getValue(row, '수령인', '수취인명', '구매자명'));
    item["휴대폰(배송)"] = clean(getValue(row, '수령인휴대폰', '수취인휴대폰', '구매자휴대폰'));
    item["우편번호(배송)"] = clean(getValue(row, '수령인우편번호', '수취인우편번호'));
    item["주소1(배송)"] = clean(getValue(row, '수령인주소', '수취인주소'));
    item["상세주소(배송)"] = clean(getValue(row, '수령인상세주소', '수취인상세주소'));

    result.data.push(item);
  });
}

function processNiceBook(data: any[], result: ProcessingResult) {
  const firstRow = data[0] || {};
  if (getValue(firstRow, '정간물명') === '') {
    result.errors.push("나이스북: '정간물명' 컬럼을 찾을 수 없습니다.");
    return;
  }

  data.forEach((row) => {
    const item = createEmptyRow();
    item["이름(주문)"] = "나이스북";
    item["상품번호"] = clean(getValue(row, '정간물명'));
    
    // Date normalization
    item["시작일"] = normalizeDate(getValue(row, '구독시작일', '시작일'));
    item["종료일"] = normalizeDate(getValue(row, '구독마감일', '종료일', '마감일'));

    // Delivery
    item["이름(배송)"] = clean(getValue(row, '수령자명', '수취인명', '주문자명'));
    item["휴대폰(배송)"] = clean(getValue(row, '수령자휴대폰', '수령자전화', '수취인휴대폰'));
    item["우편번호(배송)"] = clean(getValue(row, '수령자우편번호', '수취인우편번호'));
    item["주소1(배송)"] = clean(getValue(row, '수령자주소', '수취인주소'));
    item["상세주소(배송)"] = clean(getValue(row, '수령자상세주소', '수취인상세주소'));

    result.data.push(item);
  });
}

function normalizeDate(val: any): string {
  if (!val) return '';
  // SheetJS handles some dates as numbers or date objects
  if (val instanceof Date) {
    return val.toISOString().split('T')[0];
  }
  const str = String(val).trim();
  // Match YYYY-MM-DD or YYYY.MM.DD or YYYY/MM/DD
  const match = str.match(/^(\d{4})[ \.\-\/](\d{1,2})[ \.\-\/](\d{1,2})/);
  if (match) {
    return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  }
  return str;
}
