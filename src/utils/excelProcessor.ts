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
  // Match YYYY and MM with optional characters in between (like "년", " ", ".", "-", etc)
  const match = str.match(/(\d{4})\D*(\d{1,2})/);
  
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
 * Maps product names to specific product numbers
 */
function mapProductNumber(name: string): string {
  const cleanName = name.trim().replace(/\s+/g, ' ');
  const lowerName = cleanName.toLowerCase();

  // NE Times (2544)
  if (
    cleanName === '엔이 타임즈 NE times' ||
    cleanName === '엔이타임즈' ||
    lowerName === '엔이타임즈 ne times' ||
    lowerName.includes('ne times') && !lowerName.includes('junior') && !lowerName.includes('kids') && !lowerName.includes('kinder')
  ) {
    return '2544';
  }

  // NE Times Junior (2546)
  if (
    cleanName === '엔이타임즈주니어(월간)' ||
    lowerName === '엔이타임즈 주니어 ne times junior' ||
    lowerName === '엔이 타임즈 주니어 ne times junior' ||
    lowerName.includes('ne times junior')
  ) {
    return '2546';
  }

  // NE Times Kids (2548)
  if (lowerName === '엔이 타임즈 키즈 ne times kids' || lowerName.includes('ne times kids')) {
    return '2548';
  }

  return name;
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
          } else if (trimmedName === 'yes24' || trimmedName.toLowerCase().includes('yes24')) {
            processYes24(rawData, result);
            foundValidSheet = true;
          } else if (trimmedName === 'ssg' || trimmedName.toLowerCase().includes('ssg')) {
            processSSG(rawData, result);
            foundValidSheet = true;
          }
        });

        if (!foundValidSheet) {
          // Fallback detection by column names if sheet name doesn't match
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const rawData: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
          const firstRow = rawData[0] || {};
          const hStr = Object.keys(firstRow).join(' ');

          if (hStr.includes('주문번호') && hStr.includes('주문SEQ') && hStr.includes('상품명')) {
            processYes24(rawData, result);
            foundValidSheet = true;
          } else if (hStr.includes('순번') && hStr.includes('상품명') && hStr.includes('수취인도로명주소')) {
            processSSG(rawData, result);
            foundValidSheet = true;
          }
        }

        if (!foundValidSheet) {
          result.errors.push(`'${file.name}' 파일에서 지원하는 시트를 찾을 수 없습니다. (북콤파스, 더매거진, 나이스북, YES24, SSG)`);
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
    "정산구분": "GARA2",
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

function mapBookCompassProductNumber(name: string): string {
  const cleanName = name.trim();
  if (cleanName === '엔이타임즈주니어(월간)') return '2546';
  if (cleanName === '엔이타임즈') return '2544';
  if (cleanName === '엔이타임즈 주니어') return '2969';
  return name;
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
    item["상품번호"] = mapBookCompassProductNumber(clean(getValue(row, '품목')));
    item["시작일"] = formatYearMonth(getValue(row, '시작호수'), 'start');
    item["종료일"] = formatYearMonth(getValue(row, '만기호수'), 'end');
    
    // Delivery & Orderer Mapping
    item["이름(배송)"] = clean(getValue(row, '수령자', '수령자명')) || clean(getValue(row, '주문자', '주문자명'));
    
    // 전화번호 -> 전번(배송), 전번(주문)
    const phone = clean(getValue(row, '전화번호', '연락처', '수령자휴대전화', '수령자전화'));
    item["전번(배송)"] = phone;
    item["전번(주문)"] = phone;

    // 우편번호 -> 우편번호(배송), 우편번호(주문)
    const zip = clean(getValue(row, '우편번호', '수령자우편번호'));
    item["우편번호(배송)"] = zip;
    item["우편번호(주문)"] = zip;

    // 주    소 -> 주소1(배송), 주소1(주문)
    const address = clean(getValue(row, '주소', '수령자주소'));
    item["주소1(배송)"] = address;
    item["주소1(주문)"] = address;

    // 회 사 명 -> 상세주소(배송), 상세주소(주문)
    const company = clean(getValue(row, '회사명', '수령자상세주소', '상세주소'));
    item["상세주소(배송)"] = company;
    item["상세주소(주문)"] = company;

    result.data.push(item);
  });
}

function mapTheMagazineProductNumber(name: string): string {
  const cleanName = name.trim();
  if (cleanName.includes('엔이타임즈 NE TIMES (중고등용-주간지)')) return '2544';
  if (cleanName.includes('엔이타임즈 주니어 NE TIMES JUNIOR (월간-연12회)')) return '2546';
  if (cleanName.includes('엔이타임즈 키즈 NE TIMES KIDS (어린이-주간지)')) return '2548';
  if (cleanName.includes('엔이타임즈 킨더 NE TIMES KINDER (주간지)')) return '2538';
  if (cleanName.includes('엔이타임즈 주니어 NE TIMES JUNIOR (주간지)')) return '2969';
  return name;
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
    item["상품번호"] = mapTheMagazineProductNumber(rawProd);

    // Dates from 상품옵션
    const options = clean(getValue(row, '상품옵션'));
    const dateMatch = options.match(/(\d{4}[ \-]\d{1,2})\s*~\s*(\d{4}[ \-]\d{1,2})/);
    if (dateMatch) {
      item["시작일"] = formatYearMonth(dateMatch[1], 'start');
      item["종료일"] = formatYearMonth(dateMatch[2], 'end');
    }

    // 1. 이름(배송)
    item["이름(배송)"] = clean(getValue(row, '수령인', '수취인명', '구매자명'));

    // 2. 휴대폰: '수령인 휴대전화'를 주문/배송 모두에 입력
    const mobile = clean(getValue(row, '수령인 휴대전화', '수령인휴대폰', '구매자휴대폰'));
    item["휴대폰(주문)"] = mobile;
    item["휴대폰(배송)"] = mobile;

    // 3. 전번: '수령인 전화번호'를 주문/배송 모두에 입력
    const phone = clean(getValue(row, '수령인 전화번호', '수령인전화번호', '구매자전화번호'));
    item["전번(주문)"] = phone;
    item["전번(배송)"] = phone;

    // 4. 우편번호: '수령인 우편번호'를 주문/배송 모두에 입력
    const zip = clean(getValue(row, '수령인 우편번호', '수령인우편번호', '구매자우편번호'));
    item["우편번호(주문)"] = zip;
    item["우편번호(배송)"] = zip;

    // 5. 주소: '수령인 주소(전체)'를 4번째 띄어쓰기 기준으로 분리
    const addressFull = clean(getValue(row, '수령인 주소(전체)', '수령인주소', '구매자주소'));
    if (addressFull) {
      const parts = addressFull.split(' ');
      if (parts.length > 4) {
        const addr1 = parts.slice(0, 4).join(' ');
        const addr2 = parts.slice(4).join(' ');
        item["주소1(주문)"] = addr1;
        item["상세주소(주문)"] = addr2;
        item["주소1(배송)"] = addr1;
        item["상세주소(배송)"] = addr2;
      } else {
        item["주소1(주문)"] = addressFull;
        item["주소1(배송)"] = addressFull;
      }
    }

    // 6. 상담내용 입력: '배송메시지' 입력
    item["상담내용 입력"] = clean(getValue(row, '배송메시지', '배송옵션'));

    result.data.push(item);
  });
}

function mapNiceBookProductNumber(name: string): string {
  const cleanName = name.trim();
  if (cleanName === '엔이 타임즈 NE times') return '2544';
  if (cleanName === '엔이 타임즈 주니어 NE times JUNIOR') return '2546';
  if (cleanName === '엔이 타임즈 키즈 NE times KIDS') return '2548';
  if (cleanName === '엔이 타임즈 주니어 위클리 NE Times JUNIOR Weekly') return '2969';
  return name;
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
    item["상품번호"] = mapNiceBookProductNumber(clean(getValue(row, '정간물명')));
    
    // Date normalization
    item["시작일"] = normalizeDate(getValue(row, '구독시작일', '시작일'));
    item["종료일"] = normalizeDate(getValue(row, '구독마감일', '종료일', '마감일'));

    // 1. 이름(배송): 주문자명 + 구독자명
    const ordererName = clean(getValue(row, '주문자명'));
    const subscriberName = clean(getValue(row, '구독자명'));
    item["이름(배송)"] = `${ordererName} ${subscriberName}`.trim();

    // 2. 전번: '전화' 데이터를 주문/배송 모두에 입력
    const phone = clean(getValue(row, '전화'));
    item["전번(주문)"] = phone;
    item["전번(배송)"] = phone;

    // 3. 휴대폰: '휴대폰' 데이터를 주문/배송 모두에 입력
    const mobile = clean(getValue(row, '휴대폰'));
    item["휴대폰(주문)"] = mobile;
    item["휴대폰(배송)"] = mobile;

    // 4. 우편번호: '[' ']' 제거 후 주문/배송 모두에 입력
    const zipRaw = clean(getValue(row, '우편번호', '수령자우편번호'));
    const zipClean = zipRaw.replace(/[\[\]]/g, '');
    item["우편번호(주문)"] = zipClean;
    item["우편번호(배송)"] = zipClean;

    // 5. 주소: ')' 를 기준으로 주소1과 상세주소 분리
    const addressRaw = clean(getValue(row, '주소', '수령자주소'));
    const parenIndex = addressRaw.indexOf(')');
    if (parenIndex !== -1) {
      const addr1 = addressRaw.substring(0, parenIndex + 1).trim();
      const addr2 = addressRaw.substring(parenIndex + 1).trim();
      item["주소1(주문)"] = addr1;
      item["상세주소(주문)"] = addr2;
      item["주소1(배송)"] = addr1;
      item["상세주소(배송)"] = addr2;
    } else {
      item["주소1(주문)"] = addressRaw;
      item["주소1(배송)"] = addressRaw;
    }

    result.data.push(item);
  });
}

function mapSSGProductNumber(name: string): string {
  const lowerName = name.toLowerCase();

  if (lowerName.includes('kids') && lowerName.includes('주간')) return '2548';
  if (lowerName.includes('kinder') && lowerName.includes('주간')) return '2538';
  if (lowerName.includes('junior') && lowerName.includes('주간')) return '2969';
  if (lowerName.includes('junior w') && lowerName.includes('주간')) return '2969';
  if (lowerName.includes('times') && lowerName.includes('주간') && !lowerName.includes('junior') && !lowerName.includes('kids') && !lowerName.includes('kinder')) return '2544';
  if (lowerName.includes('junior') && lowerName.includes('월간')) return '2546';

  return mapProductNumber(name);
}

function processSSG(data: any[], result: ProcessingResult) {
  data.forEach((row) => {
    const productName = clean(getValue(row, '상품명'));
    
    // Filter: Only include items with "정기구독"
    if (!productName.includes('정기구독')) {
      return;
    }

    const item = createEmptyRow();
    
    // Identity mapping
    item["이름(주문)"] = "ssg"; 
    
    // Product Number Mapping
    item["상품번호"] = mapSSGProductNumber(productName);

    // Date Mapping: [출고예정일] 또는 [출고기준일] >> [시작일]
    // SSG format is often YYYYMMDD string
    let dateStr = clean(getValue(row, '출고예정일', '출고기준일'));
    if (dateStr && dateStr.length === 8) {
      dateStr = `${dateStr.substring(0, 4)}-${dateStr.substring(4, 6)}-${dateStr.substring(6, 8)}`;
    }
    
    const normalizedStart = normalizeDate(dateStr);
    item["시작일"] = normalizedStart;

    // End date = Start date + 1 year
    if (normalizedStart) {
      const d = new Date(normalizedStart);
      if (!isNaN(d.getTime())) {
        d.setFullYear(d.getFullYear() + 1);
        item["종료일"] = d.toISOString().split('T')[0];
      }
    }

    // Phone & Zip Mapping
    const tel = clean(getValue(row, '수취인전화번호')).replace(/--/g, '');
    const mobile = clean(getValue(row, '수취인휴대전화번호'));
    const zip = clean(getValue(row, '우편번호'));
    
    item["전번(배송)"] = tel;
    item["전번(주문)"] = tel;
    item["휴대폰(배송)"] = mobile;
    item["휴대폰(주문)"] = mobile;
    item["우편번호(배송)"] = zip;
    item["우편번호(주문)"] = zip;

    // Name Mapping
    item["이름(배송)"] = clean(getValue(row, '수취인'));

    // Address Mapping: 4th space split
    const addressFull = clean(getValue(row, '수취인도로명주소'));
    if (addressFull) {
      const parts = addressFull.split(' ');
      if (parts.length > 4) {
        const addr1 = parts.slice(0, 4).join(' ');
        const addr2 = parts.slice(4).join(' ');
        item["주소1(배송)"] = addr1;
        item["주소1(주문)"] = addr1;
        item["상세주소(배송)"] = addr2;
        item["상세주소(주문)"] = addr2;
      } else {
        item["주소1(배송)"] = addressFull;
        item["주소1(주문)"] = addressFull;
      }
    }

    // Memo
    item["상담내용 입력"] = clean(getValue(row, '고객배송메모', '배송업무메모'));

    result.data.push(item);
  });
}

function mapYes24ProductNumber(name: string): string {
  const lowerName = name.toLowerCase();

  // YES24 Specific Table Mapping
  if (lowerName.includes('kids') && lowerName.includes('주간')) return '2548';
  if (lowerName.includes('kinder') && lowerName.includes('주간')) return '2538';
  if (lowerName.includes('junior') && lowerName.includes('주간')) return '2969';
  if (lowerName.includes('times') && lowerName.includes('times') && lowerName.includes('주간') && !lowerName.includes('junior') && !lowerName.includes('kids') && !lowerName.includes('kinder')) return '2544';
  if (lowerName.includes('junior') && lowerName.includes('월간')) return '2546';

  return mapProductNumber(name);
}

function processYes24(data: any[], result: ProcessingResult) {
  data.forEach((row) => {
    const productName = clean(getValue(row, '상품명'));
    
    // Filter: Only include items with "정기구독"
    if (!productName.includes('정기구독')) {
      return;
    }

    const item = createEmptyRow();
    
    // Identity mapping (Yes24 context)
    // Always set as "yes24"
    item["이름(주문)"] = "yes24"; 
    
    // Product Number Mapping using the specialized Yes24 mapper
    item["상품번호"] = mapYes24ProductNumber(productName);

    // Date Mapping: [입금일시] >> [시작일] (연-월-일만)
    const depositTime = clean(getValue(row, '입금일시'));
    if (depositTime) {
      // Input example: "2026-04-21 오전 08:52:06"
      const datePart = depositTime.split(' ')[0]; // Extract YYYY-MM-DD
      const normalizedStart = normalizeDate(datePart);
      item["시작일"] = normalizedStart;

      // End date = Start date + 1 year
      if (normalizedStart) {
        const d = new Date(normalizedStart);
        if (!isNaN(d.getTime())) {
          d.setFullYear(d.getFullYear() + 1);
          item["종료일"] = d.toISOString().split('T')[0];
        }
      }
    }

    // Phone & Zip Mapping
    const tel = clean(getValue(row, '수령자전화')).replace(/--/g, '');
    const mobile = clean(getValue(row, '수령자휴대폰'));
    const zip = clean(getValue(row, '우편번호'));
    
    item["전번(배송)"] = tel;
    item["전번(주문)"] = tel;
    item["휴대폰(배송)"] = mobile;
    item["휴대폰(주문)"] = mobile;
    item["우편번호(배송)"] = zip;
    item["우편번호(주문)"] = zip;

    // Name Mapping
    item["이름(배송)"] = clean(getValue(row, '수령자'));

    // Address Mapping: 4th space split
    const addressFull = clean(getValue(row, '수령자주소(도로명)'));
    if (addressFull) {
      const parts = addressFull.split(' ');
      if (parts.length > 4) {
        const addr1 = parts.slice(0, 4).join(' ');
        const addr2 = parts.slice(4).join(' ');
        item["주소1(배송)"] = addr1;
        item["주소1(주문)"] = addr1;
        item["상세주소(배송)"] = addr2;
        item["상세주소(주문)"] = addr2;
      } else {
        item["주소1(배송)"] = addressFull;
        item["주소1(주문)"] = addressFull;
      }
    }

    result.data.push(item);
  });
}

/**
 * Processes pasted TSV text
 */
export async function processPastedText(text: string): Promise<ProcessingResult> {
  const result: ProcessingResult = { data: [], errors: [] };
  if (!text.trim()) return result;

  try {
    // Basic TSV parsing
    const lines = text.trim().split(/\r?\n/);
    if (lines.length < 2) {
      result.errors.push("데이터가 너무 적거나 형식이 올바르지 않습니다.");
      return result;
    }

    const headers = lines[0].split('\t').map(h => h.trim());
    const dataRows = lines.slice(1).map(line => {
      const cols = line.split('\t');
      const row: any = {};
      headers.forEach((h, i) => {
        row[h] = cols[i] || '';
      });
      return row;
    });

    // Strategy detection
    const hStr = headers.join(' ');
    if (hStr.includes('품목') && (hStr.includes('시작호수') || hStr.includes('만기호수'))) {
      processBookCompass(dataRows, result);
    } else if (hStr.includes('주문상품명') || hStr.includes('상품옵션')) {
      processTheMagazine(dataRows, result);
    } else if (hStr.includes('정간물명') || hStr.includes('구독시작일')) {
      processNiceBook(dataRows, result);
    } else if (hStr.includes('주문번호') && hStr.includes('주문SEQ') && hStr.includes('상품명')) {
      processYes24(dataRows, result);
    } else if (hStr.includes('순번') && hStr.includes('상품명') && hStr.includes('수취인도로명주소')) {
      processSSG(dataRows, result);
    } else {
      result.errors.push("데이터 형식을 실시간으로 판별하지 못했습니다. 첫 줄에 정확한 헤더(품목, 주문상품명, 정간물명, 주문번호 등)가 포함되어 있는지 확인해 주세요.");
    }

    return result;
  } catch (err) {
    result.errors.push("붙여넣은 데이터를 처리하는 중 오류가 발생했습니다.");
    return result;
  }
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
