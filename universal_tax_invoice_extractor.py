#!/usr/bin/env python3
"""
범용 세금계산서 정보 추출기 (정확한 라벨 기반, 개선된 정규식)
- 공급자/공급받는자 사업자등록번호, 상호 라벨을 기준으로 정확히 추출
"""

import re
import json
import pdfplumber

class UniversalTaxInvoiceExtractor:
    def extract_from_file(self, file_path):
        result = {
            "success": True,
            "supplier": {},
            "buyer": {},
            "items": [],
            "document_info": {},
        }

        try:
            with pdfplumber.open(file_path) as pdf:
                text = "\n".join([page.extract_text() or '' for page in pdf.pages])

            # 공급자 정보
            supplier_reg = re.search(r'공\s*급\s*자[\s\S]*?등록번호\s*(\d{3}-\d{2}-\d{5})', text)
            if supplier_reg:
                result["supplier"]["registration_number"] = supplier_reg.group(1).strip()
            else:
                raise ValueError("공급자 사업자등록번호를 찾을 수 없습니다")

            supplier_name = re.search(r'공\s*급\s*자[\s\S]*?상호[\s\S]*?([^\n]+)', text)
            if supplier_name:
                result["supplier"]["company_name"] = supplier_name.group(1).strip()
            else:
                raise ValueError("공급자 상호를 찾을 수 없습니다")

            # 공급받는자 정보
            buyer_reg = re.search(r'공\s*급\s*받\s*는\s*자[\s\S]*?등록번호\s*(\d{3}-\d{2}-\d{5})', text)
            if buyer_reg:
                result["buyer"]["registration_number"] = buyer_reg.group(1).strip()
            else:
                raise ValueError("공급받는자 사업자등록번호를 찾을 수 없습니다")

            buyer_name = re.search(r'공\s*급\s*받\s*는\s*자[\s\S]*?상호[\s\S]*?([^\n]+)', text)
            if buyer_name:
                result["buyer"]["company_name"] = buyer_name.group(1).strip()
            else:
                raise ValueError("공급받는자 상호를 찾을 수 없습니다")

            # 나머지 정보
            result["supplier"]["address"] = self._find(r'사업장\s*주소\s*([^\n]+)', text)
            result["supplier"]["ceo_name"] = self._find(r'성명\s*([^\n]+)', text)
            result["supplier"]["business_type"] = self._find(r'업태\s*([^\n]+)', text)
            result["supplier"]["item_type"] = self._find(r'종목\s*([^\n]+)', text)

            buyer_addr_match = re.findall(r'사업장\s*주소\s*([^\n]+)', text)
            if len(buyer_addr_match) >= 2:
                result["buyer"]["address"] = buyer_addr_match[1]

            # 품목 정보
            item_pattern = r'(\d{2})\s(\d{2})\s([\w\-/]+)(?:\s([^\n]+))?\s(\d+)\s([0-9,]+)\s([0-9,]+)\s([0-9,]+)'
            items = re.findall(item_pattern, text)
            for it in items:
                month, day, name, spec, qty, unit_price, supply_amt, tax_amt = it
                result["items"].append({
                    "month": month,
                    "day": day,
                    "item_name": name.strip(),
                    "specification": spec.strip() if spec else "",
                    "quantity": int(qty),
                    "unit_price": int(unit_price.replace(',', '')),
                    "supply_amount": int(supply_amt.replace(',', '')),
                    "tax_amount": int(tax_amt.replace(',', '')),
                })

        except Exception as e:
            result["success"] = False
            result["error"] = str(e)

        return result

    def _find(self, pattern, text):
        match = re.search(pattern, text)
        return match.group(1).strip() if match else None


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("input_file", help="PDF 파일 경로")
    parser.add_argument("-o", "--output", help="출력 JSON 파일", default="output.json")
    args = parser.parse_args()

    extractor = UniversalTaxInvoiceExtractor()
    result = extractor.extract_from_file(args.input_file)

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(json.dumps(result, ensure_ascii=False, indent=2))
