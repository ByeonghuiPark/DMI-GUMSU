        # 지원 형식 안내
        support_label = ttk.Label(step1_frame, 
                                text="📄 지원 형식: PDF, PNG, JPG, TXT\n"
                                     "📋 문서 타입: 세금계산서, 사업자등록증\n"
                                     "🔍 PDF는 텍스트 추출, 이미지는 OCR 처리됩니다.",
                                foreground='gray')
        support_label.grid(row=2, column=0, sticky="w", pady=(0, 10))def load_reference_file(self):
    """기존 참조 파일 불러오기 - CSV와 JSON 지원"""
    filetypes = [
        ("모든 지원 형식", "*.json;*.csv"),
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
                # JSON 파일 처리
                with open(filename, 'r', encoding='utf-8') as f:
                    self.reference_terms = json.load(f)
                    
            elif filename.endswith('.csv'):
                # CSV 파일 처리
                import csv
                self.reference_terms = {}
                
                # 인코딩 자동 감지 시도
                encodings = ['utf-8', 'cp949', 'euc-kr']
                loaded = False
                
                for encoding in encodings:
                    try:
                        with open(filename, 'r', encoding=encoding) as f:
                            # CSV 파일의 구분자 자동 감지
                            sample = f.read(1024)
                            f.seek(0)
                            
                            # 가능한 구분자들
                            sniffer = csv.Sniffer()
                            delimiter = sniffer.sniff(sample).delimiter
                            
                            reader = csv.reader(f, delimiter=delimiter)
                            
                            # 첫 번째 행이 헤더인지 확인
                            first_row = next(reader, None)
                            if first_row and len(first_row) >= 2:
                                # 헤더로 보이는 경우 건너뛰기
                                header_keywords = ['검색어', '원본', 'search', 'original', '치환어', '대체', 'replace', 'replacement']
                                is_header = any(keyword in first_row[0].lower() or keyword in first_row[1].lower() 
                                              for keyword in header_keywords)
                                
                                if not is_header:
                                    # 첫 번째 행이 데이터라면 처리
                                    if first_row[0].strip() and first_row[1].strip():
                                        self.reference_terms[first_row[0].strip()] = first_row[1].strip()
                            
                            # 나머지 행 처리
                            for row_num, row in enumerate(reader, start=2):
                                if len(row) >= 2:
                                    original = row[0].strip()
                                    replacement = row[1].strip()
                                    
                                    if original and replacement:
                                        self.reference_terms[original] = replacement
                                    elif original and not replacement:
                                        self.log(f"경고: {row_num}행에서 치환어가 비어있습니다: '{original}'")
                        
                        loaded = True
                        self.log(f"CSV 파일 로드 성공 (인코딩: {encoding})")
                        break
                        
                    except UnicodeDecodeError:
                        continue
                    except Exception as e:
                        self.log(f"CSV 파일 로드 시도 실패 ({encoding}): {str(e)}")
                        continue
                
                if not loaded:
                    raise ValueError("CSV 파일을 읽을 수 없습니다. 인코딩을 확인해주세요.")
            
            else:
                raise ValueError("지원하지 않는 파일 형식입니다. JSON 또는 CSV 파일을 선택해주세요.")
            
            # 테이블 업데이트
            self.update_terms_table()
            
            # 로드 결과 표시
            self.log(f"참조 데이터 로드 완료: {os.path.basename(filename)}")
            self.log(f"총 {len(self.reference_terms)}개 항목 로드됨")
            
            # 처음 몇 개 항목 미리보기
            if self.reference_terms:
                self.log("로드된 항목 미리보기:")
                for i, (key, value) in enumerate(list(self.reference_terms.items())[:3]):
                    self.log(f"  '{key}' → '{value}'")
                if len(self.reference_terms) > 3:
                    self.log(f"  ... 및 {len(self.reference_terms) - 3}개 더")
            
            messagebox.showinfo("완료", 
                               f"참조 데이터 로드 완료!\n"
                               f"파일: {os.path.basename(filename)}\n"
                               f"항목 수: {len(self.reference_terms)}개")
            
        except Exception as e:
            error_msg = f"파일 로드 실패: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("오류", error_msg)

def save_terms(self):
    """참조 데이터 저장 - CSV와 JSON 지원"""
    if not self.reference_terms:
        messagebox.showwarning("경고", "저장할 참조 데이터가 없습니다.")
        return
    
    filename = filedialog.asksaveasfilename(
        title="참조 데이터 저장",
        defaultextension=".json",
        filetypes=[
            ("JSON 파일", "*.json"), 
            ("CSV 파일", "*.csv"),
            ("모든 파일", "*.*")
        ]
    )
    
    if filename:
        try:
            if filename.endswith('.json'):
                # JSON 형식으로 저장
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.reference_terms, f, ensure_ascii=False, indent=2)
                    
            elif filename.endswith('.csv'):
                # CSV 형식으로 저장
                import csv
                with open(filename, 'w', encoding='utf-8', newline='') as f:
                    writer = csv.writer(f)
                    # 헤더 작성
                    writer.writerow(['원본텍스트', '치환텍스트'])
                    # 데이터 작성
                    for original, replacement in self.reference_terms.items():
                        writer.writerow([original, replacement])
            
            else:
                raise ValueError("지원하지 않는 파일 형식입니다.")
            
            self.log(f"참조 데이터 저장 완료: {os.path.basename(filename)}")
            messagebox.showinfo("완료", 
                               f"참조 데이터가 저장되었습니다.\n"
                               f"파일: {os.path.basename(filename)}")
            
        except Exception as e:
            error_msg = f"저장 실패: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("오류", error_msg)