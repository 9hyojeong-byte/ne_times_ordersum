/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface OrderData {
  id: string; // Internal ID for row management
  "정산구분": string;
  "상품번호": string;
  "시작일": string;
  "종료일": string;
  "이름(주문)": string;
  "메일(주문)": string;
  "전번(주문)": string;
  "휴대폰(주문)": string;
  "우편번호(주문)": string;
  "주소1(주문)": string;
  "상세주소(주문)": string;
  "이름(배송)": string;
  "메일(배송)": string;
  "전번(배송)": string;
  "휴대폰(배송)": string;
  "우편번호(배송)": string;
  "주소1(배송)": string;
  "상세주소(배송)": string;
  "배송방법": string;
  "계산서": string;
  "상담내용 입력": string;
}

export type OrderKey = keyof Omit<OrderData, 'id'>;

export const OUTPUT_COLUMNS: OrderKey[] = [
  "정산구분",
  "상품번호",
  "시작일",
  "종료일",
  "이름(주문)",
  "메일(주문)",
  "전번(주문)",
  "휴대폰(주문)",
  "우편번호(주문)",
  "주소1(주문)",
  "상세주소(주문)",
  "이름(배송)",
  "메일(배송)",
  "전번(배송)",
  "휴대폰(배송)",
  "우편번호(배송)",
  "주소1(배송)",
  "상세주소(배송)",
  "배송방법",
  "계산서",
  "상담내용 입력"
];

export interface ProcessingResult {
  data: OrderData[];
  errors: string[];
}
