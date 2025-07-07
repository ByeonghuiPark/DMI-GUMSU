#!/usr/bin/env python3
"""
HWPX 자동화 통합 GUI 애플리케이션 (타입 오류 해결 버전)
- 세금계산서 정보 추출
- HWPX 텍스트 치환 
- 이미지 삽입
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import threading
from pathlib import Path
import sys

# 기존 모듈들 임포트
try:
    from universal_tax_invoice_extractor import UniversalTaxInvoiceExtractor
    from enhanced_hwpx_processor import EnhancedHWPXProcessor
    from hwpx_image_inserter import HWPXImageInserter
    MODULES_AVAILABLE = True
except ImportError as e:
    MODULES_AVAILABLE = False
    IMPORT_ERROR = str(e)

class HWPXAutomationGUI:
    """HWPX 자동화 통합 GUI"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("HWPX 자동화 도구 v1.0")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # 스타일 설정
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 변수들
        self.current_step = 0
        self.files = {
            'tax_invoice': '',
            'hwpx_template': '', 
            'output_hwpx': '',
            'image_file': '',
            'final_output': ''
        }
        
        self.extracted_data = {}
        self.reference_terms = {}
        
        # GUI 초기화
        self.create_widgets()
        
        # 모듈 가용성 확인
        if not MODULES_AVAILABLE:
            messagebox.showerror("오류", f"필수 모듈을 찾을 수 없습니다:\n{IMPORT_ERROR}")
            self.root.quit()
    
    def create_widgets(self):
        """GUI 위젯 생성"""
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 제목
        title_label = ttk.Label(main_frame, text="HWPX 자동화 통합 도구", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 왼쪽 패널 - 단계별 탭
        left_frame = ttk.Frame(main_frame, padding="5")
        left_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        
        # 오른쪽 패널 - 로그 및 결과
        right_frame = ttk.Frame(main_frame, padding="5")
        right_frame.grid(row=1, column=1, sticky="nsew")
        
        # 탭 생성
        self.notebook = ttk.Notebook(left_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        # 단계별 탭 생성
        self.create_step1_tab()  # 세금계산서 분석
        self.create_step2_tab()  # 참조 데이터 생성
        self.create_step3_tab()  # HWPX 텍스트 치환
        self.create_step4_tab()  # 이미지 삽입
        
        # 오른쪽 패널 - 로그 영역
        self.create_log_panel(right_frame)
        
        # 하단 버튼
        self.create_bottom_buttons(main_frame)
    
    def create_step1_tab(self):
        """1단계: 세금계산서 정보 추출"""
        
        step1_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(step1_frame, text="1단계: 세금계산서 분석")
        
        # 파일 선택
        ttk.Label(step1_frame, text="세금계산서 파일 선택:", 
                 font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        file_frame = ttk.Frame(step1_frame)
        file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)
        
        self.tax_invoice_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.tax_invoice_var, 
                 state='readonly').grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(file_frame, text="파일 선택", 
                  command=self.select_tax_invoice).grid(row=0, column=1)
        
        # 지원 형식 안내
        support_label = ttk.Label(step1_frame, 
                                text="지원 형식: PDF, PNG, JPG, TXT\n"
                                     "PDF는 텍스트 추출, 이미지는 OCR 처리됩니다.",
                                foreground='gray')
        support_label.grid(row=2, column=0, sticky="w", pady=(0, 10))
        
        # 분석 버튼
        self.analyze_btn = ttk.Button(step1_frame, text="세금계산서 분석", 
                                     command=self.analyze_tax_invoice, state='disabled')
        self.analyze_btn.grid(row=3, column=0, pady=10)
        
        # 결과 미리보기
        ttk.Label(step1_frame, text="추출된 정보 미리보기:", 
                 font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky="w", pady=(10, 5))
        
        self.step1_result = scrolledtext.ScrolledText(step1_frame, height=8, width=50)
        self.step1_result.grid(row=5, column=0, sticky="nsew", pady=(0, 10))
        step1_frame.columnconfigure(0, weight=1)
        step1_frame.rowconfigure(5, weight=1)
    
    def create_step2_tab(self):
        """2단계: 참조 데이터 생성"""
        
        step2_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(step2_frame, text="2단계: 참조 데이터")
        
        ttk.Label(step2_frame, text="텍스트 치환을 위한 참조 데이터 설정", 
                 font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 10))
        
        # 자동 생성 버튼
        auto_frame = ttk.Frame(step2_frame)
        auto_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        self.auto_generate_btn = ttk.Button(auto_frame, text="자동 참조 데이터 생성", 
                                           command=self.auto_generate_terms, state='disabled')
        self.auto_generate_btn.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(auto_frame, text="기존 파일 불러오기", 
                  command=self.load_reference_file).grid(row=0, column=1)
        
        # 참조 데이터 편집
        ttk.Label(step2_frame, text="참조 데이터 편집:", 
                 font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky="w", pady=(10, 5))
        
        # 테이블 프레임
        table_frame = ttk.Frame(step2_frame)
        table_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Treeview로 편집 가능한 테이블
        self.terms_tree = ttk.Treeview(table_frame, columns=('original', 'replacement'), 
                                      show='headings', height=10)
        self.terms_tree.heading('original', text='원본 텍스트')
        self.terms_tree.heading('replacement', text='치환될 텍스트')
        self.terms_tree.column('original', width=200)
        self.terms_tree.column('replacement', width=200)
        
        # 스크롤바
        tree_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.terms_tree.yview)
        self.terms_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.terms_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")
        
        # 테이블 조작 버튼
        btn_frame = ttk.Frame(step2_frame)
        btn_frame.grid(row=4, column=0, pady=10)
        
        ttk.Button(btn_frame, text="항목 추가", command=self.add_term).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="항목 삭제", command=self.delete_term).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="저장", command=self.save_terms).grid(row=0, column=2, padx=5)
        
        step2_frame.columnconfigure(0, weight=1)
        step2_frame.rowconfigure(3, weight=1)
    
    def create_step3_tab(self):
        """3단계: HWPX 텍스트 치환"""
        
        step3_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(step3_frame, text="3단계: 텍스트 치환")
        
        # HWPX 템플릿 파일 선택
        ttk.Label(step3_frame, text="HWPX 템플릿 파일:", 
                 font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        hwpx_frame = ttk.Frame(step3_frame)
        hwpx_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        hwpx_frame.columnconfigure(0, weight=1)
        
        self.hwpx_template_var = tk.StringVar()
        ttk.Entry(hwpx_frame, textvariable=self.hwpx_template_var, 
                 state='readonly').grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(hwpx_frame, text="파일 선택", 
                  command=self.select_hwpx_template).grid(row=0, column=1)
        
        # 치환 옵션
        options_frame = ttk.LabelFrame(step3_frame, text="치환 옵션", padding="10")
        options_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        
        self.case_sensitive_var = tk.BooleanVar()
        self.whole_word_var = tk.BooleanVar(value=True)
        self.backup_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(options_frame, text="대소문자 구분", 
                       variable=self.case_sensitive_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(options_frame, text="전체 단어만 매칭", 
                       variable=self.whole_word_var).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(options_frame, text="원본 백업", 
                       variable=self.backup_var).grid(row=0, column=2, sticky="w")
        
        # 출력 파일명
        ttk.Label(step3_frame, text="출력 파일명:", 
                 font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky="w", pady=(10, 5))
        
        output_frame = ttk.Frame(step3_frame)
        output_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_hwpx_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_hwpx_var).grid(
            row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(output_frame, text="위치 선택", 
                  command=self.select_output_location).grid(row=0, column=1)
        
        # 텍스트 치환 실행
        self.process_hwpx_btn = ttk.Button(step3_frame, text="텍스트 치환 실행", 
                                          command=self.process_hwpx_text, state='disabled')
        self.process_hwpx_btn.grid(row=5, column=0, pady=20)
        
        step3_frame.columnconfigure(0, weight=1)
    
    def create_step4_tab(self):
        """4단계: 이미지 삽입"""
        
        step4_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(step4_frame, text="4단계: 이미지 삽입")
        
        # 이미지 파일 선택
        ttk.Label(step4_frame, text="삽입할 이미지 파일:", 
                 font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        img_frame = ttk.Frame(step4_frame)
        img_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        img_frame.columnconfigure(0, weight=1)
        
        self.image_file_var = tk.StringVar()
        ttk.Entry(img_frame, textvariable=self.image_file_var, 
                 state='readonly').grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(img_frame, text="이미지 선택", 
                  command=self.select_image_file).grid(row=0, column=1)
        
        # 위치 설정
        pos_frame = ttk.LabelFrame(step4_frame, text="삽입 위치", padding="10")
        pos_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        
        ttk.Label(pos_frame, text="표 번호:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.table_index_var = tk.StringVar(value="0")
        ttk.Entry(pos_frame, textvariable=self.table_index_var, width=10).grid(
            row=0, column=1, padx=(0, 20))
        
        ttk.Label(pos_frame, text="행 번호:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.row_index_var = tk.StringVar(value="0")
        ttk.Entry(pos_frame, textvariable=self.row_index_var, width=10).grid(
            row=0, column=3, padx=(0, 20))
        
        ttk.Label(pos_frame, text="열 번호:").grid(row=0, column=4, sticky="w", padx=(0, 5))
        self.col_index_var = tk.StringVar(value="0")
        ttk.Entry(pos_frame, textvariable=self.col_index_var, width=10).grid(
            row=0, column=5)
        
        # 표 구조 분석 버튼
        self.analyze_table_btn = ttk.Button(pos_frame, text="표 구조 분석", 
                                           command=self.analyze_table_structure, state='disabled')
        self.analyze_table_btn.grid(row=1, column=0, columnspan=6, pady=10)
        
        # 이미지 옵션
        img_opt_frame = ttk.LabelFrame(step4_frame, text="이미지 옵션", padding="10")
        img_opt_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        
        ttk.Label(img_opt_frame, text="너비(mm):").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.img_width_var = tk.StringVar(value="80")
        ttk.Entry(img_opt_frame, textvariable=self.img_width_var, width=10).grid(
            row=0, column=1, padx=(0, 20))
        
        ttk.Label(img_opt_frame, text="높이(mm):").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.img_height_var = tk.StringVar(value="60")
        ttk.Entry(img_opt_frame, textvariable=self.img_height_var, width=10).grid(
            row=0, column=3, padx=(0, 20))
        
        self.maintain_ratio_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(img_opt_frame, text="비율 유지", 
                       variable=self.maintain_ratio_var).grid(row=0, column=4, sticky="w")
        
        # 정렬 옵션
        ttk.Label(img_opt_frame, text="정렬:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.alignment_var = tk.StringVar(value="center")
        align_combo = ttk.Combobox(img_opt_frame, textvariable=self.alignment_var, 
                                  values=["left", "center", "right"], state="readonly", width=10)
        align_combo.grid(row=1, column=1, pady=(10, 0))
        
        # 최종 출력 파일명
        ttk.Label(step4_frame, text="최종 출력 파일명:", 
                 font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky="w", pady=(10, 5))
        
        final_frame = ttk.Frame(step4_frame)
        final_frame.grid(row=5, column=0, sticky="ew", pady=(0, 10))
        final_frame.columnconfigure(0, weight=1)
        
        self.final_output_var = tk.StringVar()
        ttk.Entry(final_frame, textvariable=self.final_output_var).grid(
            row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(final_frame, text="위치 선택", 
                  command=self.select_final_output).grid(row=0, column=1)
        
        # 이미지 삽입 실행
        self.insert_image_btn = ttk.Button(step4_frame, text="이미지 삽입 실행", 
                                          command=self.insert_image, state='disabled')
        self.insert_image_btn.grid(row=6, column=0, pady=20)
        
        step4_frame.columnconfigure(0, weight=1)
    
    def create_log_panel(self, parent):
        """로그 패널 생성"""
        
        ttk.Label(parent, text="처리 로그", 
                 font=('Arial', 12, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(parent, height=20, width=50)
        self.log_text.grid(row=1, column=0, sticky="nsew")
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
    
    def create_bottom_buttons(self, parent):
        """하단 버튼들 생성"""
        
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="전체 프로세스 실행", 
                  command=self.run_full_process).grid(row=0, column=0, padx=10)
        ttk.Button(btn_frame, text="로그 지우기", 
                  command=self.clear_log).grid(row=0, column=1, padx=10)
        ttk.Button(btn_frame, text="종료", 
                  command=self.root.quit).grid(row=0, column=2, padx=10)
    
    def log(self, message):
        """로그 메시지 추가"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def clear_log(self):
        """로그 지우기"""
        self.log_text.delete(1.0, tk.END)
    
    def select_tax_invoice(self):
        """세금계산서 파일 선택"""
        filetypes = [
            ("모든 지원 형식", "*.pdf;*.png;*.jpg;*.jpeg;*.txt"),
            ("PDF 파일", "*.pdf"),
            ("이미지 파일", "*.png;*.jpg;*.jpeg"),
            ("텍스트 파일", "*.txt"),
            ("모든 파일", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="세금계산서 파일 선택",
            filetypes=filetypes
        )
        
        if filename:
            self.tax_invoice_var.set(filename)
            self.files['tax_invoice'] = filename
            self.analyze_btn.config(state='normal')
            self.log(f"세금계산서 파일 선택: {os.path.basename(filename)}")
    
    def analyze_tax_invoice(self):
        """세금계산서 분석"""
        if not self.files['tax_invoice']:
            messagebox.showerror("오류", "세금계산서 파일을 선택해주세요.")
            return
        
        self.log("세금계산서 분석 시작...")
        
        def analyze_worker():
            try:
                extractor = UniversalTaxInvoiceExtractor()
                result = extractor.extract_from_file(self.files['tax_invoice'])
                
                if result['success']:
                    self.extracted_data = result
                    
                    # 결과 표시
                    self.step1_result.delete(1.0, tk.END)
                    
                    # 주요 정보 요약
                    summary = "=== 추출된 정보 요약 ===\n\n"
                    
                    # 문서 정보
                    doc_info = result.get('document_info', {})
                    if doc_info:
                        summary += "📄 문서 정보:\n"
                        summary += f"  승인번호: {doc_info.get('approval_number', 'N/A')}\n"
                        summary += f"  발행일자: {doc_info.get('issue_date', 'N/A')}\n\n"
                    
                    # 공급자 정보
                    supplier = result.get('supplier', {})
                    if supplier:
                        summary += "🏢 공급자:\n"
                        summary += f"  회사명: {supplier.get('company_name', 'N/A')}\n"
                        summary += f"  등록번호: {supplier.get('registration_number', 'N/A')}\n\n"
                    
                    # 공급받는자 정보
                    buyer = result.get('buyer', {})
                    if buyer:
                        summary += "🏛️ 공급받는자:\n"
                        summary += f"  회사명: {buyer.get('company_name', 'N/A')}\n"
                        summary += f"  등록번호: {buyer.get('registration_number', 'N/A')}\n\n"
                    
                    # 금액 정보
                    amounts = result.get('amounts', {})
                    if amounts:
                        summary += "💰 금액 정보:\n"
                        if amounts.get('total_amount'):
                            summary += f"  총액: {amounts['total_amount']:,}원\n"
                        if amounts.get('supply_amount'):
                            summary += f"  공급가액: {amounts['supply_amount']:,}원\n"
                        if amounts.get('tax_amount'):
                            summary += f"  세액: {amounts['tax_amount']:,}원\n\n"
                    
                    # 품목 정보
                    items = result.get('items', [])
                    if items:
                        summary += f"📦 품목 정보 ({len(items)}개):\n"
                        for i, item in enumerate(items[:3], 1):
                            summary += f"  {i}. {item.get('item_name', 'N/A')}\n"
                        if len(items) > 3:
                            summary += f"  ... 및 {len(items) - 3}개 더\n"
                    
                    self.step1_result.insert(tk.END, summary)
                    
                    # 다음 단계 활성화
                    self.auto_generate_btn.config(state='normal')
                    
                    self.log("세금계산서 분석 완료!")
                    
                else:
                    self.log(f"세금계산서 분석 실패: {result.get('error', '알 수 없는 오류')}")
                    messagebox.showerror("오류", f"분석 실패: {result.get('error', '알 수 없는 오류')}")
                    
            except Exception as e:
                self.log(f"분석 중 오류 발생: {str(e)}")
                messagebox.showerror("오류", f"분석 중 오류 발생: {str(e)}")
        
        # 별도 스레드에서 실행
        thread = threading.Thread(target=analyze_worker)
        thread.daemon = True
        thread.start()
    
    def auto_generate_terms(self):
        """자동 참조 데이터 생성"""
        if not self.extracted_data:
            messagebox.showerror("오류", "먼저 세금계산서를 분석해주세요.")
            return
        
        self.log("참조 데이터 자동 생성 중...")
        
        try:
            # 추출된 데이터를 기반으로 참조 데이터 생성
            terms = {}
            
            # 회사명 치환
            supplier = self.extracted_data.get('supplier', {})
            buyer = self.extracted_data.get('buyer', {})
            
            if supplier.get('company_name'):
                terms["[공급자명]"] = supplier['company_name']
            if supplier.get('registration_number'):
                terms["[공급자등록번호]"] = supplier['registration_number']
            
            if buyer.get('company_name'):
                terms["[공급받는자명]"] = buyer['company_name']
            if buyer.get('registration_number'):
                terms["[공급받는자등록번호]"] = buyer['registration_number']
            
            # 문서 정보
            doc_info = self.extracted_data.get('document_info', {})
            if doc_info.get('approval_number'):
                terms["[승인번호]"] = doc_info['approval_number']
            if doc_info.get('issue_date'):
                terms["[발행일자]"] = doc_info['issue_date']
            
            # 금액 정보
            amounts = self.extracted_data.get('amounts', {})
            if amounts.get('total_amount'):
                terms["[총금액]"] = f"{amounts['total_amount']:,}원"
            if amounts.get('supply_amount'):
                terms["[공급가액]"] = f"{amounts['supply_amount']:,}원"
            if amounts.get('tax_amount'):
                terms["[세액]"] = f"{amounts['tax_amount']:,}원"
            
            # 품목 정보 (첫 번째 품목)
            items = self.extracted_data.get('items', [])
            if items:
                first_item = items[0]
                if first_item.get('item_name'):
                    terms["[주요품목]"] = first_item['item_name']
                if first_item.get('quantity'):
                    terms["[수량]"] = str(first_item['quantity'])
            
            # 연락처 정보
            contacts = self.extracted_data.get('contacts', {})
            if contacts.get('phones'):
                terms["[연락처]"] = contacts['phones'][0]
            if contacts.get('emails'):
                terms["[이메일]"] = contacts['emails'][0]
            
            # 기본 템플릿 용어 추가
            default_terms = {
                "[오늘날짜]": "2025-07-07",
                "[작성자]": "담당자",
                "[부서]": "영업부",
                "[제목]": "세금계산서 관련 문서"
            }
            terms.update(default_terms)
            
            # 참조 데이터 저장
            self.reference_terms = terms
            
            # 테이블에 표시
            self.update_terms_table()
            
            self.log(f"참조 데이터 자동 생성 완료: {len(terms)}개 항목")
            messagebox.showinfo("완료", f"{len(terms)}개의 참조 데이터가 생성되었습니다.")
            
        except Exception as e:
            self.log(f"참조 데이터 생성 중 오류: {str(e)}")
            messagebox.showerror("오류", f"참조 데이터 생성 실패: {str(e)}")
    
    def load_reference_file(self):
        """기존 참조 파일 불러오기"""
        filetypes = [
            ("JSON 파일", "*.json"),
            ("CSV 파일", "*.csv"),
            ("모든 파일", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="참조 데이터 파일 선택",
            filetypes=filetypes
        )
        
        if filename:
            try:
                if filename.endswith('.json'):
                    with open(filename, 'r', encoding='utf-8') as f:
                        self.reference_terms = json.load(f)
                elif filename.endswith('.csv'):
                    import csv
                    self.reference_terms = {}
                    with open(filename, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        next(reader, None)  # 헤더 건너뛰기
                        for row in reader:
                            if len(row) >= 2:
                                self.reference_terms[row[0]] = row[1]
                
                self.update_terms_table()
                self.log(f"참조 데이터 로드 완료: {os.path.basename(filename)}")
                
            except Exception as e:
                messagebox.showerror("오류", f"파일 로드 실패: {str(e)}")
    
    def update_terms_table(self):
        """참조 데이터 테이블 업데이트"""
        # 기존 항목 모두 삭제
        for item in self.terms_tree.get_children():
            self.terms_tree.delete(item)
        
        # 새 항목 추가
        for original, replacement in self.reference_terms.items():
            self.terms_tree.insert('', 'end', values=(original, replacement))
    
    def add_term(self):
        """새 참조 항목 추가"""
        # 간단한 입력 대화상자
        dialog = tk.Toplevel(self.root)
        dialog.title("새 참조 항목 추가")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="원본 텍스트:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        original_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=original_var, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="치환될 텍스트:").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        replacement_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=replacement_var, width=30).grid(row=1, column=1, padx=10, pady=10)
        
        def add_and_close():
            original = original_var.get().strip()
            replacement = replacement_var.get().strip()
            
            if original and replacement:
                self.reference_terms[original] = replacement
                self.terms_tree.insert('', 'end', values=(original, replacement))
                dialog.destroy()
            else:
                messagebox.showerror("오류", "모든 필드를 입력해주세요.")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="추가", command=add_and_close).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="취소", command=dialog.destroy).grid(row=0, column=1, padx=5)
    
    def delete_term(self):
        """선택된 참조 항목 삭제"""
        selected = self.terms_tree.selection()
        if not selected:
            messagebox.showwarning("경고", "삭제할 항목을 선택해주세요.")
            return
        
        for item in selected:
            values = self.terms_tree.item(item, 'values')
            if values:
                original = values[0]
                if original in self.reference_terms:
                    del self.reference_terms[original]
                self.terms_tree.delete(item)
    
    def save_terms(self):
        """참조 데이터 저장"""
        if not self.reference_terms:
            messagebox.showwarning("경고", "저장할 참조 데이터가 없습니다.")
            return
        
        filename = filedialog.asksaveasfilename(
            title="참조 데이터 저장",
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("CSV 파일", "*.csv")]
        )
        
        if filename:
            try:
                if filename.endswith('.json'):
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(self.reference_terms, f, ensure_ascii=False, indent=2)
                elif filename.endswith('.csv'):
                    import csv
                    with open(filename, 'w', encoding='utf-8', newline='') as f:
                        writer = csv.writer(f)
                        writer.writerow(['원본텍스트', '치환텍스트'])
                        for original, replacement in self.reference_terms.items():
                            writer.writerow([original, replacement])
                
                self.log(f"참조 데이터 저장 완료: {os.path.basename(filename)}")
                messagebox.showinfo("완료", "참조 데이터가 저장되었습니다.")
                
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {str(e)}")
    
    def select_hwpx_template(self):
        """HWPX 템플릿 파일 선택"""
        filename = filedialog.askopenfilename(
            title="HWPX 템플릿 파일 선택",
            filetypes=[("HWPX 파일", "*.hwpx"), ("모든 파일", "*.*")]
        )
        
        if filename:
            self.hwpx_template_var.set(filename)
            self.files['hwpx_template'] = filename
            self.process_hwpx_btn.config(state='normal')
            self.analyze_table_btn.config(state='normal')
            
            # 기본 출력 파일명 설정
            base_name = Path(filename).stem
            output_name = f"{base_name}_processed.hwpx"
            self.output_hwpx_var.set(output_name)
            
            self.log(f"HWPX 템플릿 선택: {os.path.basename(filename)}")
    
    def select_output_location(self):
        """출력 파일 위치 선택"""
        filename = filedialog.asksaveasfilename(
            title="출력 파일 저장 위치",
            defaultextension=".hwpx",
            filetypes=[("HWPX 파일", "*.hwpx"), ("모든 파일", "*.*")]
        )
        
        if filename:
            self.output_hwpx_var.set(filename)
    
    def process_hwpx_text(self):
        """HWPX 텍스트 치환 실행"""
        if not self.files['hwpx_template']:
            messagebox.showerror("오류", "HWPX 템플릿 파일을 선택해주세요.")
            return
        
        if not self.reference_terms:
            messagebox.showerror("오류", "참조 데이터가 없습니다.")
            return
        
        output_file = self.output_hwpx_var.get()
        if not output_file:
            messagebox.showerror("오류", "출력 파일명을 입력해주세요.")
            return
        
        self.log("HWPX 텍스트 치환 시작...")
        
        def process_worker():
            try:
                processor = EnhancedHWPXProcessor(log_level="WARNING")
                
                # 치환 옵션
                options = {
                    'case_sensitive': self.case_sensitive_var.get(),
                    'whole_word_only': self.whole_word_var.get(),
                    'backup_original': self.backup_var.get(),
                    'use_advanced_regex': True
                }
                
                result = processor.search_and_replace_text(
                    hwpx_file=self.files['hwpx_template'],
                    reference_data=self.reference_terms,
                    output_file=output_file,
                    replacement_options=options
                )
                
                if result['success']:
                    self.files['output_hwpx'] = output_file
                    self.insert_image_btn.config(state='normal')
                    
                    # 기본 최종 출력 파일명 설정
                    base_name = Path(output_file).stem
                    final_name = f"{base_name}_final.hwpx"
                    self.final_output_var.set(final_name)
                    
                    self.log(f"텍스트 치환 완료! {result['total_replacements']}개 항목 치환")
                    self.log(f"출력 파일: {output_file}")
                    
                    # 치환 내역 표시
                    if result.get('replacements'):
                        self.log("주요 치환 내역:")
                        for replacement in result['replacements'][:5]:
                            if isinstance(replacement, dict):
                                search = replacement['search_term']
                                replace = replacement['replacement_term']
                                count = replacement['count']
                                self.log(f"  '{search}' → '{replace}' ({count}회)")
                    
                else:
                    self.log(f"텍스트 치환 실패: {result.get('error', '알 수 없는 오류')}")
                    messagebox.showerror("오류", f"치환 실패: {result.get('error', '알 수 없는 오류')}")
                    
            except Exception as e:
                self.log(f"치환 중 오류 발생: {str(e)}")
                messagebox.showerror("오류", f"치환 중 오류 발생: {str(e)}")
        
        # 별도 스레드에서 실행
        thread = threading.Thread(target=process_worker)
        thread.daemon = True
        thread.start()
    
    def select_image_file(self):
        """이미지 파일 선택"""
        filetypes = [
            ("이미지 파일", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
            ("PNG 파일", "*.png"),
            ("JPG 파일", "*.jpg;*.jpeg"),
            ("모든 파일", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="삽입할 이미지 선택",
            filetypes=filetypes
        )
        
        if filename:
            self.image_file_var.set(filename)
            self.files['image_file'] = filename
            self.log(f"이미지 파일 선택: {os.path.basename(filename)}")
    
    def analyze_table_structure(self):
        """표 구조 분석"""
        if not self.files['output_hwpx'] and not self.files['hwpx_template']:
            messagebox.showerror("오류", "분석할 HWPX 파일이 없습니다.")
            return
        
        # 치환된 파일이 있으면 그것을, 없으면 템플릿 사용
        hwpx_file = self.files.get('output_hwpx') or self.files['hwpx_template']
        
        self.log(f"표 구조 분석 중: {os.path.basename(hwpx_file)}")
        
        def analyze_worker():
            try:
                inserter = HWPXImageInserter()
                
                # 임시로 로그를 캡처하기 위한 방법
                import io
                import contextlib
                
                log_capture = io.StringIO()
                with contextlib.redirect_stdout(log_capture):
                    inserter.list_tables_in_hwpx(hwpx_file)
                
                captured_output = log_capture.getvalue()
                self.log("표 구조 분석 결과:")
                self.log(captured_output)
                
            except Exception as e:
                self.log(f"표 구조 분석 실패: {str(e)}")
                messagebox.showerror("오류", f"표 구조 분석 실패: {str(e)}")
        
        # 별도 스레드에서 실행
        thread = threading.Thread(target=analyze_worker)
        thread.daemon = True
        thread.start()
    
    def select_final_output(self):
        """최종 출력 파일 위치 선택"""
        filename = filedialog.asksaveasfilename(
            title="최종 출력 파일 저장 위치",
            defaultextension=".hwpx",
            filetypes=[("HWPX 파일", "*.hwpx"), ("모든 파일", "*.*")]
        )
        
        if filename:
            self.final_output_var.set(filename)
    
    def insert_image(self):
        """이미지 삽입 실행"""
        if not self.files.get('output_hwpx'):
            messagebox.showerror("오류", "먼저 텍스트 치환을 완료해주세요.")
            return
        
        if not self.files.get('image_file'):
            messagebox.showerror("오류", "삽입할 이미지를 선택해주세요.")
            return
        
        final_output = self.final_output_var.get()
        if not final_output:
            messagebox.showerror("오류", "최종 출력 파일명을 입력해주세요.")
            return
        
        # 위치 정보 검증
        try:
            table_index = int(self.table_index_var.get())
            row_index = int(self.row_index_var.get())
            col_index = int(self.col_index_var.get())
        except ValueError:
            messagebox.showerror("오류", "표, 행, 열 번호는 숫자여야 합니다.")
            return
        
        self.log("이미지 삽입 시작...")
        
        def insert_worker():
            try:
                inserter = HWPXImageInserter()
                
                # 이미지 옵션
                image_options = {
                    'width': int(self.img_width_var.get()),
                    'height': int(self.img_height_var.get()),
                    'maintain_ratio': self.maintain_ratio_var.get(),
                    'alignment': self.alignment_var.get()
                }
                
                result = inserter.insert_image_to_table(
                    hwpx_file=self.files['output_hwpx'],
                    image_file=self.files['image_file'],
                    output_file=final_output,
                    table_index=table_index,
                    row_index=row_index,
                    col_index=col_index,
                    image_options=image_options
                )
                
                if result['success']:
                    self.files['final_output'] = final_output
                    
                    self.log("이미지 삽입 완료!")
                    self.log(f"위치: {result['position']}")
                    self.log(f"크기: {result['image_size']}")
                    self.log(f"최종 파일: {final_output}")
                    
                    messagebox.showinfo("완료", "모든 작업이 완료되었습니다!")
                    
                else:
                    self.log(f"이미지 삽입 실패: {result.get('error', '알 수 없는 오류')}")
                    messagebox.showerror("오류", f"이미지 삽입 실패: {result.get('error', '알 수 없는 오류')}")
                    
            except Exception as e:
                self.log(f"이미지 삽입 중 오류 발생: {str(e)}")
                messagebox.showerror("오류", f"이미지 삽입 중 오류 발생: {str(e)}")
        
        # 별도 스레드에서 실행
        thread = threading.Thread(target=insert_worker)
        thread.daemon = True
        thread.start()
    
    def run_full_process(self):
        """전체 프로세스 자동 실행"""
        self.log("=== 전체 프로세스 시작 ===")
        
        # 필수 파일 확인
        if not self.files.get('tax_invoice'):
            messagebox.showerror("오류", "세금계산서 파일을 선택해주세요.")
            return
        
        if not self.files.get('hwpx_template'):
            messagebox.showerror("오류", "HWPX 템플릿 파일을 선택해주세요.")
            return
        
        def full_process_worker():
            try:
                # 1단계: 세금계산서 분석
                self.log("1/4 단계: 세금계산서 분석...")
                if not self.extracted_data:
                    extractor = UniversalTaxInvoiceExtractor()
                    result = extractor.extract_from_file(self.files['tax_invoice'])
                    if result['success']:
                        self.extracted_data = result
                        self.log("세금계산서 분석 완료")
                    else:
                        raise Exception(f"세금계산서 분석 실패: {result.get('error')}")
                
                # 2단계: 참조 데이터 생성
                self.log("2/4 단계: 참조 데이터 생성...")
                if not self.reference_terms:
                    self.auto_generate_terms()
                
                # 3단계: 텍스트 치환
                self.log("3/4 단계: 텍스트 치환...")
                processor = EnhancedHWPXProcessor(log_level="WARNING")
                
                output_file = self.output_hwpx_var.get()
                if not output_file:
                    base_name = Path(self.files['hwpx_template']).stem
                    output_file = f"{base_name}_processed.hwpx"
                    self.output_hwpx_var.set(output_file)
                
                options = {
                    'case_sensitive': self.case_sensitive_var.get(),
                    'whole_word_only': self.whole_word_var.get(),
                    'backup_original': self.backup_var.get(),
                    'use_advanced_regex': True
                }
                
                result = processor.search_and_replace_text(
                    hwpx_file=self.files['hwpx_template'],
                    reference_data=self.reference_terms,
                    output_file=output_file,
                    replacement_options=options
                )
                
                if result['success']:
                    self.files['output_hwpx'] = output_file
                    self.log(f"텍스트 치환 완료: {result['total_replacements']}개 치환")
                else:
                    raise Exception(f"텍스트 치환 실패: {result.get('error')}")
                
                # 4단계: 이미지 삽입 (선택사항)
                if self.files.get('image_file'):
                    self.log("4/4 단계: 이미지 삽입...")
                    
                    inserter = HWPXImageInserter()
                    
                    final_output = self.final_output_var.get()
                    if not final_output:
                        base_name = Path(output_file).stem
                        final_output = f"{base_name}_final.hwpx"
                        self.final_output_var.set(final_output)
                    
                    try:
                        table_index = int(self.table_index_var.get())
                        row_index = int(self.row_index_var.get())
                        col_index = int(self.col_index_var.get())
                    except ValueError:
                        table_index = row_index = col_index = 0
                        self.log("위치 정보가 잘못되어 기본값(0,0,0) 사용")
                    
                    image_options = {
                        'width': int(self.img_width_var.get()),
                        'height': int(self.img_height_var.get()),
                        'maintain_ratio': self.maintain_ratio_var.get(),
                        'alignment': self.alignment_var.get()
                    }
                    
                    result = inserter.insert_image_to_table(
                        hwpx_file=self.files['output_hwpx'],
                        image_file=self.files['image_file'],
                        output_file=final_output,
                        table_index=table_index,
                        row_index=row_index,
                        col_index=col_index,
                        image_options=image_options
                    )
                    
                    if result['success']:
                        self.files['final_output'] = final_output
                        self.log(f"이미지 삽입 완료: {final_output}")
                    else:
                        self.log(f"이미지 삽입 실패: {result.get('error')}")
                else:
                    self.log("4/4 단계: 이미지 삽입 건너뜀 (이미지 파일 없음)")
                
                self.log("=== 전체 프로세스 완료! ===")
                messagebox.showinfo("완료", "모든 작업이 성공적으로 완료되었습니다!")
                
            except Exception as e:
                self.log(f"전체 프로세스 실패: {str(e)}")
                messagebox.showerror("오류", f"전체 프로세스 실패: {str(e)}")
        
        # 별도 스레드에서 실행
        thread = threading.Thread(target=full_process_worker)
        thread.daemon = True
        thread.start()
    
    def run(self):
        """GUI 실행"""
        self.root.mainloop()

def main():
    """메인 함수"""
    app = HWPXAutomationGUI()
    app.run()

if __name__ == "__main__":
    main()